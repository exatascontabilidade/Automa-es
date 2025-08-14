# extratos_pro.py
# Autor: Gemini AI
# Data: 13 de agosto de 2025
# Descrição: Script completo para extrair, limpar e enriquecer dados de extratos bancários
# de múltiplos formatos, usando IA (OpenAI) com fallback para heurística local.

import os
import re
import glob
import sys
import json
import argparse
import warnings
import traceback
import logging
from typing import List, Optional, Tuple, Dict, Any

import pandas as pd
import yaml
from dotenv import load_dotenv

# --- Configurações Iniciais ---
warnings.filterwarnings("ignore", category=UserWarning)
logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s", handlers=[logging.StreamHandler(sys.stdout)])

# Carrega as variáveis do arquivo .env para o ambiente
load_dotenv()

def _try_import(module_name: str, package_name: str = None):
    """Tenta importar um módulo, fornecendo uma mensagem clara em caso de falha."""
    try:
        return __import__(module_name)
    except ImportError:
        pkg = package_name or module_name
        logging.warning(f"Módulo opcional '{module_name}' não encontrado. Para habilitar esta funcionalidade, instale com: pip install {pkg}")
        return None

dateparser = _try_import("dateparser", "python-dateutil")
pdfplumber = _try_import("pdfplumber")
ofxparse = _try_import("ofxparse")
camelot = _try_import("camelot", '"camelot-py[cv]"')
pytesseract = _try_import("pytesseract")
openai_mod = _try_import("openai")
from PIL import Image


class ExtratoProcessor:
    """
    Processa arquivos de extratos bancários (PDF, CSV, Excel, OFX) para extrair,
    limpar e enriquecer dados de transações.
    """
    def __init__(self, config_path: str):
        self.config = self._load_config(config_path)
        self.NEG_PARENS = re.compile(r"^\(\s*(.+?)\s*\)$")
        self.CURRENCY = re.compile(r"^(R\$|US\$|\$|€)?\s*(.+)$", re.I)
        self.LIXO_PAT = re.compile(f"(?i)\\b({'|'.join(self.config['patterns']['lixo'])})\\b")
        self.LINE_DATE_PAT = r"(?:\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)"
        self.LINE_VAL_PAT = r"(?:-?\s*\(?\s*R?\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}\)?-?"

    def _load_config(self, path: str) -> Dict[str, Any]:
        """Carrega a configuração de um arquivo YAML."""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                logging.info(f"Usando arquivo de configuração: {path}")
                return yaml.safe_load(f)
        except FileNotFoundError:
            logging.error(f"FATAL: Arquivo de configuração '{path}' não encontrado. Crie um ou use --config para especificar o caminho.")
            sys.exit(1)
        except Exception as e:
            logging.error(f"FATAL: Erro ao carregar ou parsear o arquivo de configuração: {e}")
            sys.exit(1)

    # ---------- Funções Utilitárias ----------
    def normalize_colname(self, s: str) -> str:
        s = (str(s) or "").strip().lower().replace("\n", " ").replace("\r", " ")
        s = re.sub(r"\s+", " ", s)
        s = s.translate(str.maketrans("çáàãâéêíóôõú", "caaaaeeiooou"))
        return s.replace(".", "").replace(":", "")

    def parse_brl_number(self, text: Any) -> Optional[float]:
        if text is None: return None
        s = str(text).strip().replace("\u00A0", "").replace(" ", "")
        if not s: return None
        m = self.CURRENCY.match(s)
        if m: s = m.group(2).strip()
        tail_neg = s.endswith(("-", "D", "DB")); s = s.rstrip("+-CD B")
        m = self.NEG_PARENS.match(s); neg_paren = bool(m)
        if m: s = m.group(1).strip()
        if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
        elif s.count(",") == 1 and s.count(".") == 0: s = s.replace(",", ".")
        elif s.count(".") > 1 and s.count(",") == 0: s = s.replace(".", "")
        try:
            val = float(s)
            if tail_neg or neg_paren: val = -abs(val)
            return val
        except (ValueError, TypeError): return None

    def clean_description(self, series: pd.Series) -> pd.Series:
        return series.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    def find_first_column(self, columns: List[str], candidates: List[str]) -> Optional[str]:
        norm_map = {c: self.normalize_colname(c) for c in columns}
        cand_norms = [self.normalize_colname(c) for c in candidates]
        for col, ncol in norm_map.items():
            for nc in cand_norms:
                if ncol == nc or (nc in ncol) or (ncol in nc): return col
        return None

    def infer_year_from_filename(self, path: str) -> Optional[int]:
        m = re.search(r"(20\d{2})", os.path.basename(path))
        return int(m.group(1)) if m else None

    def normalize_dates(self, series: pd.Series, default_year: Optional[int] = None) -> pd.Series:
        def _parse(x):
            if pd.isna(x): return pd.NaT
            s = str(x).strip()
            if re.fullmatch(r"\d{1,2}/\d{1,2}", s) and default_year: s = f"{s}/{default_year}"
            if dateparser:
                dt = dateparser.parse(s, languages=["pt", "en"])
                if dt: return dt
            return pd.to_datetime(s, errors="coerce", dayfirst=True)
        return series.apply(_parse)

    def dedupe(self, df: pd.DataFrame) -> pd.DataFrame:
        key = (df["Data"].astype(str) + "|" +
               df["Descrição"].str.strip().str.lower() + "|" +
               df["Valor"].round(2).astype(str))
        return df.loc[~key.duplicated()].reset_index(drop=True)

    # ---------- Leitores de Arquivo ----------
    def read_csv_like(self, path: str) -> Optional[pd.DataFrame]:
        for sep in [",", ";", "\t", "|"]:
            for enc in ["utf-8-sig", "latin1", "cp1252"]:
                try:
                    df = pd.read_csv(path, sep=sep, encoding=enc, dtype=str, on_bad_lines='skip')
                    if df.shape[1] >= 2 and not df.empty: return df
                except Exception: continue
        return None

    def read_excel_like(self, path: str) -> Optional[pd.DataFrame]:
        try:
            xls = pd.ExcelFile(path)
            best_sheet = None
            for sh in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sh, dtype=str)
                if df is not None and not df.empty:
                    if best_sheet is None or df.size > best_sheet.size: best_sheet = df
            return best_sheet
        except Exception as e: logging.debug(f"Falha ao ler Excel {path}: {e}"); return None

    def read_ofx_like(self, path: str) -> Optional[pd.DataFrame]:
        if not ofxparse: return None
        try:
            with open(path, "r", encoding="latin1") as f: ofx = ofxparse.OfxParser.parse(f)
            rows = [{"data": tx.date, "descricao": tx.memo, "valor": tx.amount}
                    for acct in ofx.accounts for tx in getattr(acct, "statement", type('obj', (object,), {'transactions': []})()).transactions]
            return pd.DataFrame(rows) if rows else None
        except Exception as e: logging.debug(f"Falha ao ler OFX {path}: {e}"); return None

    def read_pdf_tables(self, path: str) -> List[pd.DataFrame]:
        frames = []
        if camelot:
            try:
                tables = camelot.read_pdf(path, pages="all", flavor="lattice", line_scale=40)
                frames.extend(t.df for t in tables if t.df.shape[0] > 1)
                if not frames:
                     tables = camelot.read_pdf(path, pages="all", flavor="stream")
                     frames.extend(t.df for t in tables if t.df.shape[0] > 1)
            except Exception: pass
        if frames: return frames
        if pdfplumber:
            try:
                with pdfplumber.open(path) as pdf:
                    for p in pdf.pages:
                        tables = p.extract_tables() or []
                        frames.extend(pd.DataFrame(t) for t in tables if t and len(t) > 1)
            except Exception: pass
        return frames
    
    def pdf_to_text(self, path: str, do_ocr: bool) -> str:
        text_content = ""
        if not pdfplumber: return ""
        try:
            with pdfplumber.open(path) as pdf:
                # 1. Tentar extração de texto nativo
                for page in pdf.pages: text_content += page.extract_text() or ""
                # 2. Se o texto for insignificante e OCR for permitido, usar OCR
                if len(text_content.strip()) < 100 and do_ocr and pytesseract:
                    logging.info("PDF com pouco texto, tentando OCR...")
                    text_content = ""
                    for page in pdf.pages:
                        img = page.to_image(resolution=300).original
                        text_content += pytesseract.image_to_string(img, lang="por+eng")
        except Exception as e:
            logging.error(f"Erro ao processar PDF {os.path.basename(path)}: {e}")
        return text_content
    
    # ---------- Lógica de Parsing e Construção do DataFrame ----------
    def rows_from_text_lines(self, text: str, year_hint: int) -> Optional[pd.DataFrame]:
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if not line: continue
            m_date = re.match(rf"^\s*({self.LINE_DATE_PAT})", line)
            if not m_date: continue
            
            rest_of_line = line[m_date.end():]
            vals_found = list(re.finditer(self.LINE_VAL_PAT, rest_of_line))
            if not vals_found: continue
            
            last_val_match = vals_found[-1]
            desc = rest_of_line[:last_val_match.start()].strip()
            val_txt = last_val_match.group(0).strip()
            
            # Tenta extrair crédito/débito da última parte
            if val_txt.endswith(('C', 'CR')): val_sign = 1
            elif val_txt.endswith(('D', 'DB')): val_sign = -1
            else: val_sign = 1 # Padrão

            valor = self.parse_brl_number(val_txt)
            if valor is not None: valor = abs(valor) * val_sign
            
            rows.append({"Data": m_date.group(1), "Descrição": desc, "Valor": valor})
        
        if not rows: return None
        df = pd.DataFrame(rows)
        df["Data"] = self.normalize_dates(df["Data"], default_year=year_hint)
        df["Descrição"] = self.clean_description(df["Descrição"])
        df = df.dropna(subset=["Data", "Valor"])
        df = df[~df["Descrição"].str.contains(self.LIXO_PAT, na=False, case=False)]
        return df if not df.empty else None

    def build_final_frame(self, df: pd.DataFrame, year_hint: Optional[int]) -> Optional[pd.DataFrame]:
        df = df.dropna(axis=1, how="all").reset_index(drop=True)
        if df.empty: return None

        # Renomeia colunas para o padrão
        cc = self.config['column_candidates']
        rename_map = {}
        for col in df.columns:
            if not rename_map.get("Data") and self.find_first_column([col], cc['date']): rename_map[col] = "Data"
            elif not rename_map.get("Descrição") and self.find_first_column([col], cc['desc']): rename_map[col] = "Descrição"
            elif not rename_map.get("Valor") and self.find_first_column([col], cc['value']): rename_map[col] = "Valor"
            elif not rename_map.get("Credito") and self.find_first_column([col], cc['credit']): rename_map[col] = "Credito"
            elif not rename_map.get("Debito") and self.find_first_column([col], cc['debit']): rename_map[col] = "Debito"
            elif not rename_map.get("Protocolo") and self.find_first_column([col], cc['protocol']): rename_map[col] = "Protocolo"
        
        df.rename(columns=rename_map, inplace=True)
        
        # Processa colunas de Débito/Crédito se Valor não existir
        if "Valor" not in df.columns and ("Debito" in df.columns or "Credito" in df.columns):
            db = df["Debito"].apply(self.parse_brl_number).fillna(0) if "Debito" in df else 0
            cr = df["Credito"].apply(self.parse_brl_number).fillna(0) if "Credito" in df else 0
            df["Valor"] = cr - db
        elif "Valor" in df.columns:
            df["Valor"] = df["Valor"].apply(self.parse_brl_number)

        # Normaliza colunas-chave
        if "Data" in df.columns: df["Data"] = self.normalize_dates(df["Data"], default_year=year_hint)
        else: return None # Data é obrigatória
        
        if "Descrição" not in df.columns:
            # Fallback: Acha a coluna com mais texto
            text_cols = [c for c in df.columns if df[c].dtype == 'object' and c != 'Data']
            if text_cols: df['Descrição'] = df[text_cols[0]]
            else: df['Descrição'] = ""

        df["Descrição"] = self.clean_description(df["Descrição"])
        
        final_cols = ["Data", "Descrição", "Valor", "Protocolo"]
        out_df = df[[c for c in final_cols if c in df.columns]].copy()
        
        out_df = out_df.dropna(subset=["Data", "Valor"])
        out_df = out_df[out_df["Descrição"].str.strip() != '']
        out_df = out_df[~out_df["Descrição"].str.contains(self.LIXO_PAT, na=False, case=False)]
        
        return out_df if not out_df.empty else None

    # ---------- Enriquecimento (IA e Heurística) ----------
    def ai_enrich(self, df: pd.DataFrame) -> pd.DataFrame:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key or not openai_mod:
            logging.warning("API Key da OpenAI não definida ou módulo 'openai' não instalado. Usando heurística local.")
            return self.heuristic_enrich(df)
        try: client = openai_mod.OpenAI(api_key=api_key)
        except Exception as e:
            logging.error(f"Falha ao inicializar o cliente OpenAI: {e}. Usando heurística local.")
            return self.heuristic_enrich(df)

        ai_cfg = self.config['ai']
        batch_size = ai_cfg['batch_size']
        model = ai_cfg['model']
        categories_str = ", ".join(f'"{cat}"' for cat in ai_cfg['categories'])
        prompt_template = ai_cfg['prompt'].format(categories=categories_str)
        
        rows_to_process = df.to_dict(orient="records")
        enriched_rows = []
        
        logging.info(f"Iniciando enriquecimento com IA ({model}). Total de {len(rows_to_process)} linhas em lotes de {batch_size}.")
        for i in range(0, len(rows_to_process), batch_size):
            chunk = rows_to_process[i:i+batch_size]
            logging.info(f"Processando lote {i//batch_size + 1}...")
            content = json.dumps(chunk, ensure_ascii=False, indent=2)
            messages = [{"role": "system", "content": prompt_template}, {"role": "user", "content": content}]
            
            try:
                resp = client.chat.completions.create(
                    model=model, messages=messages, temperature=0.1, response_format={"type": "json_object"}
                )
                response_json = json.loads(resp.choices[0].message.content)
                items = response_json.get("transacoes", [])
                if not isinstance(items, list) or len(items) != len(chunk):
                    raise ValueError(f"Resposta da IA com tamanho inesperado. Recebido: {len(items)}, Esperado: {len(chunk)}")
                for original, enriched in zip(chunk, items):
                    merged = original.copy()
                    merged.update({"Categoria": enriched.get("Categoria", "Outros"), "Contraparte": enriched.get("Contraparte"), "Nota": enriched.get("Nota")})
                    enriched_rows.append(merged)
            except Exception as e:
                logging.error(f"Falha no lote {i//batch_size + 1}: {e}. Usando heurística local para este lote.")
                fallback_df = self.heuristic_enrich(pd.DataFrame(chunk))
                enriched_rows.extend(fallback_df.to_dict(orient="records"))

        final_df = pd.DataFrame(enriched_rows)
        for col in ["Categoria", "Contraparte", "Nota"]:
            if col not in final_df.columns: final_df[col] = None
        return final_df

    def heuristic_enrich(self, df: pd.DataFrame) -> pd.DataFrame:
        logging.info("Usando método de enriquecimento por heurística local.")
        df_out = df.copy()
        def cat(desc: str, valor: float) -> str:
            d = str(desc or "").lower(); valor = valor or 0
            if "pix" in d: return "PIX Enviado" if valor < 0 else "PIX Recebido"
            if "boleto" in d or "pagto" in d: return "Pagamento de Boleto"
            if any(k in d for k in ["tarifa", "taxa", "cesta"]): return "Tarifa Bancária"
            if any(k in d for k in ["compra", "debito", "credito"]): return "Compra no Cartão"
            if any(k in d for k in ["salario", "vencimento"]): return "Salário"
            if any(k in d for k in ["imposto", "darf", "gps"]): return "Pagamento de Impostos"
            if any(k in d for k in ["transfer", "ted", "doc"]): return "Transferência"
            return "Entrada" if valor > 0 else "Saída"
        df_out["Categoria"] = df_out.apply(lambda row: cat(row.get("Descrição"), row.get("Valor")), axis=1)
        df_out["Contraparte"] = None
        df_out["Nota"] = "Enriquecido por heurística"
        return df_out

    # ---------- Orquestrador de Arquivo Único ----------
    def process_file(self, path: str, ocr: bool) -> Optional[pd.DataFrame]:
        ext = os.path.splitext(path)[1].lower()
        year_hint = self.infer_year_from_filename(path)
        df_raw = None

        logging.info(f"Analisando {os.path.basename(path)}...")
        if ext in (".csv", ".txt"): df_raw = self.read_csv_like(path)
        elif ext in (".xls", ".xlsx"): df_raw = self.read_excel_like(path)
        elif ext in (".ofx", ".ofc"): df_raw = self.read_ofx_like(path)
        elif ext == ".pdf":
            pdf_tables = self.read_pdf_tables(path)
            if pdf_tables:
                # Concatena todas as tabelas encontradas no PDF
                df_raw = pd.concat([self.build_final_frame(t, year_hint) for t in pdf_tables], ignore_index=True)
            else: # Se não houver tabelas, processa como texto
                pdf_text = self.pdf_to_text(path, do_ocr=ocr)
                return self.rows_from_text_lines(pdf_text, year_hint)
        
        if df_raw is None: return None
        return self.build_final_frame(df_raw, year_hint)


def save_output(df: pd.DataFrame, base_path: str, fmt: str):
    """Salva o DataFrame no formato especificado."""
    path = f"{base_path}.{fmt}"
    if fmt == "csv": df.to_csv(path, index=False, sep=";", encoding="utf-8-sig")
    else: df.to_excel(path, index=False)
    logging.info(f"✅ Arquivo salvo: {os.path.basename(path)}")


def main():
    parser = argparse.ArgumentParser(description="Processador de Extratos Bancários com IA.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("-i", "--input", required=True, help="Arquivo ou pasta de entrada.")
    parser.add_argument("-o", "--outdir", help="Diretório de saída. Padrão: mesmo do arquivo de origem.")
    parser.add_argument("-f", "--format", choices=["xlsx", "csv"], default="xlsx", help="Formato de saída tabular.")
    parser.add_argument("--ocr", action="store_true", help="Forçar OCR em PDFs (pode ser lento).")
    parser.add_argument("--no-ai", action="store_true", help="Desativa o enriquecimento com IA, usa apenas a heurística local.")
    parser.add_argument("-c", "--config", default="config.yaml", help="Caminho para o arquivo de configuração YAML.")
    args = parser.parse_args()

    processor = ExtratoProcessor(config_path=args.config)
    
    target_path = args.input
    if os.path.isdir(target_path):
        files = [f for ext in ("*.pdf", "*.csv", "*.txt", "*.xls", "*.xlsx", "*.ofx", "*.ofc")
                   for f in glob.glob(os.path.join(target_path, ext))]
    elif os.path.isfile(target_path):
        files = [target_path]
    else:
        logging.error(f"Caminho não encontrado: {target_path}"); sys.exit(1)

    if not files: logging.warning("Nenhum arquivo compatível encontrado no caminho especificado."); return

    all_dfs = []
    for path in sorted(files):
        try:
            df = processor.process_file(path, ocr=args.ocr)
            if df is None or df.empty:
                logging.warning(f"Nenhuma transação extraída de {os.path.basename(path)}.")
                continue
            
            logging.info(f"Extraídas {len(df)} transações de {os.path.basename(path)}.")
            
            df_sorted = df.sort_values(by="Data").reset_index(drop=True)
            df_deduped = processor.dedupe(df_sorted)

            if args.no_ai:
                df_enriched = processor.heuristic_enrich(df_deduped)
            else:
                df_enriched = processor.ai_enrich(df_deduped)
            
            out_dir = args.outdir or os.path.dirname(path)
            os.makedirs(out_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(path))[0]
            save_output(df_enriched, os.path.join(out_dir, f"{base_name}_processado"), args.format)
            
            df_enriched["__arquivo_origem"] = os.path.basename(path)
            all_dfs.append(df_enriched)
            
        except Exception as e:
            logging.error(f"Erro fatal ao processar {os.path.basename(path)}: {e}")
            traceback.print_exc()

    if len(all_dfs) > 1:
        logging.info("Consolidando todos os resultados...")
        df_consolidado = pd.concat(all_dfs, ignore_index=True).sort_values(by=["__arquivo_origem", "Data"])
        
        out_dir = args.outdir or (target_path if os.path.isdir(target_path) else os.path.dirname(target_path))
        save_output(df_consolidado, os.path.join(out_dir, "CONSOLIDADO"), args.format)

    logging.info("Processamento concluído.")

if __name__ == "__main__":
    main()