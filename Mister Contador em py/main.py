# extratos_ai_txt.py (versão robusta)
# -----------------------------------------------------------------------------
# 1) Extrai lançamentos (Data, Descrição, Valor, Protocolo) de extratos:
#    - PDF (texto / escaneado com OCR opcional)
#    - CSV/TXT (vários separadores/encodes)
#    - XLS/XLSX
#    - OFX/OFC
# 2) Enriquecimento por IA (opcional):
#    - Se OPENAI_API_KEY estiver definido e --no-ai NÃO for usado, chama a API
#    - Caso contrário, usa heurísticas locais
# 3) Gera arquivos no MESMO diretório do script ou em --outdir:
#    - {nome}_limpo.xlsx/.csv/.parquet
#    - {nome}_limpo.txt  (lista “plana” dos lançamentos)
#    - consolidado.* e consolidado.txt (se processar uma PASTA)
#
# Uso rápido:
#   python extratos_ai_txt.py
#   python extratos_ai_txt.py -i "C:\extratos" --format csv
#   python extratos_ai_txt.py -i "arquivo.pdf" --ocr
# -----------------------------------------------------------------------------

import os, re, glob, sys, json, argparse, warnings, traceback
from typing import List, Optional, Tuple, Dict
from datetime import datetime
import pandas as pd
warnings.filterwarnings("ignore", category=UserWarning)

# ---------- imports opcionais ----------
def _try_import(m):
    try: return __import__(m)
    except Exception: return None

dateparser  = _try_import("dateparser")
pdfplumber  = _try_import("pdfplumber")
ofxparse    = _try_import("ofxparse")
camelot     = _try_import("camelot")     # opcional para tabelas em PDF
pytesseract = _try_import("pytesseract")
openai_mod  = _try_import("openai")      # IA opcional (pip install openai)
from PIL import Image  # pillow

# ---------- utilidades ----------
def get_script_dir():
    if getattr(sys,"frozen",False) and hasattr(sys,"_MEIPASS"):
        return os.path.dirname(sys.executable)
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.path.dirname(os.path.abspath(sys.argv[0] or "."))

def normalize_colname(s: str) -> str:
    s = (str(s) or "").strip().lower()
    s = s.replace("\n"," ").replace("\r"," ")
    s = re.sub(r"\s+"," ", s)
    s = s.translate(str.maketrans("çáàãâéêíóôõú","caaaaeeiooou"))
    return s.replace(".","").replace(":","")

NEG_PARENS = re.compile(r"^\(\s*(.+?)\s*\)$")
CURRENCY   = re.compile(r"^(R\$|US\$|\$|€)?\s*(.+)$", re.I)
LIXO_PAT   = re.compile(
    r"(?i)\b("
    r"saldo\s*(anterior|final|do\s+dia)|total\s*geral|p[aá]g\.\s*\d+|"
    r"via\s+do\s+cliente|comprovante|ag[eê]ncia|conta|extrato\s+(?:banc[aá]rio|simplificado)|"
    r"n[oº]\s*doc(?:umento)?|autentica[cç][aã]o\s*eletr[oô]nica"
    r")\b"
)

def parse_brl_number(text) -> Optional[float]:
    if text is None: return None
    s = str(text).strip().replace("\u00A0","").replace(" ", "")
    if s == "": return None
    m = CURRENCY.match(s)
    if m: s = m.group(2).strip()
    tail_neg = s.endswith("-"); s = s.rstrip("+-")
    m = NEG_PARENS.match(s); neg_paren = bool(m)
    if m: s = m.group(1).strip()

    # Formatos: 1.234,56 | 1234,56 | 1,234.56 | 1234.56 | -1.234,56 | (1.234,56)
    if "," in s and "." in s:
        # assume . milhar e , decimal
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    elif s.count(".") > 1 and s.count(",") == 0:
        s = s.replace(".", "")  # inteiro com pontos de milhar

    try:
        val = float(s)
        if tail_neg or neg_paren: val = -abs(val)
        return val
    except:
        return None

def clean_description(series: pd.Series) -> pd.Series:
    return (series.astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.replace(r"\u00A0", " ", regex=True)
            .str.strip())

def find_first_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {c: normalize_colname(c) for c in columns}
    cand_norms = [normalize_colname(c) for c in candidates]
    # match exato ou inclusão bi-direcional
    for col, ncol in norm_map.items():
        for nc in cand_norms:
            if ncol==nc or (nc in ncol) or (ncol in nc):
                return col
    # reforço: contém qualquer um dos candidatos
    for col in columns:
        ncol = normalize_colname(col)
        if any(nc in ncol for nc in cand_norms):
            return col
    return None

COL_DATE = ["data","date","dt","competencia","competência","movimento","lançamento","lancamento"]
COL_DESC = ["historico","histórico","descricao","descrição","description","documento",
            "complemento","favorecido","detalhe","detalhes","observacao","observação","texto","hist"]
COL_VAL  = ["valor","amount","vl","val","valor (r$)","r$","valorfinal","total","mov.","movimento"]
COL_DB   = ["debito","débito","db","saída","saida","debitos","deb."]
COL_CR   = ["credito","crédito","cr","entrada","creditos","cred."]
COL_DC   = ["d/c","dc","tipo","natureza"]

def pick_value_column(df: pd.DataFrame):
    cols = list(df.columns)
    c_val = find_first_column(cols, COL_VAL)
    c_cr  = find_first_column(cols, COL_CR)
    c_db  = find_first_column(cols, COL_DB)
    c_dc  = find_first_column(cols, COL_DC)
    return c_val, c_cr, c_db, c_dc

def infer_year_from_filename(path: str) -> Optional[int]:
    m = re.search(r"(20\d{2})", os.path.basename(path))
    return int(m.group(1)) if m else None

def normalize_dates(series: pd.Series, default_year: Optional[int]=None) -> pd.Series:
    def _p(x):
        if pd.isna(x): return pd.NaT
        s = str(x).strip()
        # datas sem ano (ex: 05/01) -> injeta ano padrão
        if re.fullmatch(r"\d{1,2}/\d{1,2}", s) and default_year:
            s = f"{s}/{default_year}"
        if dateparser:
            dt = dateparser.parse(s, languages=["pt","pt-br","en"])
            return dt
        return pd.to_datetime(s, errors="coerce", dayfirst=True)
    return series.apply(_p)

def apply_dc_sign(df: pd.DataFrame, col_dc: Optional[str], vals: pd.Series) -> pd.Series:
    if not col_dc or col_dc not in df.columns: return vals
    dc = df[col_dc].astype(str).str.strip().str.upper()
    out = vals.copy()
    is_deb = dc.str.startswith(("D","DEB"))
    out[is_deb] = -out[is_deb].abs()
    return out

def dedupe(df: pd.DataFrame) -> pd.DataFrame:
    key = (df["Data"].astype(str)+"|"+
           df["Descrição"].astype(str).str.strip().str.lower()+"|"+
           df["Valor"].round(2).astype(str))
    return df.loc[~key.duplicated()].reset_index(drop=True)

# ---------- leitores de arquivo ----------
def read_csv_like(path: str) -> Optional[pd.DataFrame]:
    for sep in [",",";","\t","|"]:
        for enc in ["utf-8-sig","latin1","cp1252","utf-8"]:
            try:
                df = pd.read_csv(path, sep=sep, encoding=enc, dtype=str)
                if df.shape[1] >= 2 and not df.empty: return df
            except: pass
    return None

def read_excel_like(path: str) -> Optional[pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        best=None
        for sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh, dtype=str)
            if df is None or df.empty: continue
            if best is None or df.size > best.size: best = df
        return best
    except: return None

def read_ofx_like(path: str) -> Optional[pd.DataFrame]:
    if not ofxparse: return None
    try:
        with open(path,"r",encoding="latin1") as f:
            ofx = ofxparse.OfxParser.parse(f)
        rows=[]
        for acct in ofx.accounts:
            st = getattr(acct, "statement", None) or getattr(acct, "stmt", None)
            if not st: 
                continue
            for tx in st.transactions:
                rows.append({
                    "data": tx.date.strftime("%d/%m/%Y") if tx.date else None,
                    "descricao": tx.memo or tx.payee or "",
                    "valor": tx.amount
                })
        return pd.DataFrame(rows)
    except: return None

def read_pdf_tables(path: str) -> List[pd.DataFrame]:
    frames=[]
    # camelot primeiro
    if camelot:
        try:
            for flavor in ["lattice","stream"]:
                tables = camelot.read_pdf(path, pages="all", flavor=flavor)
                for t in tables:
                    df = t.df
                    if df is not None and df.shape[0] >= 2:
                        header=df.iloc[0].tolist()
                        if any(isinstance(x,str) and re.search(r"data|descri|hist|valor|deb|cred", normalize_colname(x)) for x in header):
                            df.columns = header; df = df.iloc[1:].reset_index(drop=True)
                        frames.append(df)
                if frames: return frames
        except: pass
    # pdfplumber tabelas
    if pdfplumber:
        try:
            with pdfplumber.open(path) as pdf:
                for p in pdf.pages:
                    try:
                        for tb in (p.extract_tables() or []):
                            df = pd.DataFrame(tb)
                            if df.shape[0] >= 2:
                                header=df.iloc[0].tolist()
                                if any(isinstance(x,str) and re.search(r"data|descri|hist|valor|deb|cred", normalize_colname(x)) for x in header):
                                    df.columns = header; df = df.iloc[1:].reset_index(drop=True)
                            frames.append(df)
                    except: pass
        except: pass
    return frames

def pdf_text_lines(path: str) -> List[str]:
    if not pdfplumber: return []
    lines=[]
    try:
        with pdfplumber.open(path) as pdf:
            for p in pdf.pages:
                txt = p.extract_text() or ""
                for ln in txt.splitlines():
                    ln = ln.strip()
                    if ln: lines.append(ln)
    except: pass
    return lines

def is_scanned_pdf(path: str) -> bool:
    if not pdfplumber: return False
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages: return False
            return len(pdf.pages[0].chars) == 0
    except: return False

def ocr_pdf_to_text(path: str, dpi=300) -> List[str]:
    if not (pdfplumber and pytesseract): return []
    texts=[]
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                img = page.to_image(resolution=dpi).original  # PIL.Image
                txt = pytesseract.image_to_string(img, lang="por+eng")
                texts.append(txt or "")
    except: pass
    return texts

# ---------- parser de linhas (PDF texto/ocr) ----------
LINE_DATE = r"(?:\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)"
LINE_VAL  = r"(?:\(?-?\d{1,3}(?:\.\d{3})*,\d{2}\)?-?)"

def rows_from_text_lines(text: str) -> pd.DataFrame:
    lines=[ln for ln in text.splitlines() if ln.strip()]
    rows=[]
    for ln in lines:
        m_date = re.match(rf"^\s*({LINE_DATE})\s+(.*)", ln)
        if not m_date: continue
        rest = m_date.group(2)
        m_val = list(re.finditer(LINE_VAL, rest))
        if not m_val: continue
        val_txt = m_val[-1].group(0)
        desc = rest[:m_val[-1].start()].strip()
        # protocolo: maior sequência de 9+ dígitos antes do valor
        protos = list(re.finditer(r"\b(\d{9,})\b", desc))
        protocolo = protos[-1].group(1) if protos else None
        if protocolo:
            # remove protocolo da descrição
            idx = desc.rfind(protocolo)
            if idx >= 0:
                desc = (desc[:idx] + desc[idx+len(protocolo):]).strip()
        rows.append({"data": m_date.group(1), "descricao": desc, "protocolo": protocolo, "valor_txt": val_txt})
    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows, dtype=str)
    # normaliza
    yy = df["data"].str.extract(r"(\d{4})")[0]
    default_year = int(yy.iloc[0]) if yy.notna().any() else None
    data = normalize_dates(df["data"], default_year=default_year)
    out = pd.DataFrame({
        "Data": data.dt.strftime("%Y-%m-%d"),
        "Descrição": clean_description(df["descricao"]),
        "Protocolo": df["protocolo"],
        "Valor": df["valor_txt"].apply(parse_brl_number)
    })
    out = out[~out["Descrição"].str.contains(LIXO_PAT, na=False)]
    # ordena
    try:
        d = pd.to_datetime(out["Data"], errors="coerce")
        out = out.assign(_d=d).sort_values("_d").drop(columns=["_d"])
    except: pass
    if out["Valor"].notna().any(): out["Valor"] = out["Valor"].round(2)
    return out

# ---------- montagem do DF final (genérico) ----------
def build_final_frame(df: pd.DataFrame, year_hint: Optional[int]=None) -> Optional[pd.DataFrame]:
    if df is None or df.empty: return None
    df = df.dropna(axis=1, how="all")
    if df.empty: return None

    col_data = find_first_column(df.columns.tolist(), COL_DATE)
    col_desc = find_first_column(df.columns.tolist(), COL_DESC)
    c_val, c_cr, c_db, c_dc = pick_value_column(df)

    if col_data is None:
        for c in df.columns:
            if df[c].astype(str).str.contains(r"\b\d{1,2}/\d{1,2}(/\d{2,4})?\b").any():
                col_data = c; break
    if col_desc is None:
        txt = {c: df[c].astype(str).str.count(r"[A-Za-zÀ-ú]").sum() for c in df.columns}
        if txt: col_desc = max(txt, key=txt.get)

    if c_val:
        v = df[c_val].astype(str).apply(parse_brl_number)
        v = apply_dc_sign(df, c_dc, v)
    elif c_cr or c_db:
        cr = df[c_cr].astype(str).apply(parse_brl_number) if c_cr else pd.Series([0.0]*len(df))
        db = df[c_db].astype(str).apply(parse_brl_number) if c_db else pd.Series([0.0]*len(df))
        v = (cr.fillna(0)-db.fillna(0))
    else:
        cand=[]
        for c in df.columns:
            vals = df[c].astype(str).apply(parse_brl_number)
            if vals.notna().sum() >= max(3, len(df)//4): cand.append((c, vals))
        if cand: _, v = sorted(cand, key=lambda x: x[1].notna().sum(), reverse=True)[0]
        else: v = pd.Series([None]*len(df))

    year_hint = year_hint or infer_year_from_filename("<unknown>")
    data_norm = normalize_dates(df[col_data], default_year=year_hint) if col_data else pd.Series([pd.NaT]*len(df))
    desc_norm = clean_description(df[col_desc]) if col_desc else pd.Series([""]*len(df))

    out = pd.DataFrame({"Data": data_norm.dt.strftime("%Y-%m-%d"),
                        "Descrição": desc_norm,
                        "Valor": v})
    # protocolo: tenta extrair campo caso exista
    for cand in ["protocolo","numero","número","nr","ref","ref."]:
        if cand in [normalize_colname(c) for c in df.columns]:
            # encontra coluna original que bate no cand
            src = [c for c in df.columns if normalize_colname(c)==cand][0]
            out["Protocolo"] = df[src].astype(str).str.extract(r"(\d{6,})")[0]
            break
    if "Protocolo" not in out.columns:
        out["Protocolo"] = None

    out = out[~out["Descrição"].str.contains(LIXO_PAT, na=False)]
    if out["Valor"].notna().any(): out["Valor"]=out["Valor"].round(2)
    out = out[["Data","Descrição","Valor","Protocolo"]]
    out = out.dropna(subset=["Data","Descrição"], how="all")
    out = out[out["Data"].notna() | out["Descrição"].ne("")]
    return out if not out.empty else None

# ---------- pipeline de um arquivo ----------
def process_file(path: str, ocr: bool=False) -> Tuple[Optional[pd.DataFrame], Dict]:
    ext = os.path.splitext(path)[1].lower()
    meta = {"arquivo": os.path.basename(path), "ok": False, "metodo": None, "linhas": 0, "erro": None}

    try:
        if ext in (".csv",".txt"):
            df = read_csv_like(path); meta["metodo"]="csv/txt"
            if df is not None:
                out = build_final_frame(df, year_hint=infer_year_from_filename(path))
                if out is not None: meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            return None, meta

        if ext in (".xls",".xlsx"):
            df = read_excel_like(path); meta["metodo"]="excel"
            if df is not None:
                out = build_final_frame(df, year_hint=infer_year_from_filename(path))
                if out is not None: meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            return None, meta

        if ext in (".ofx",".ofc"):
            df = read_ofx_like(path); meta["metodo"]="ofx/ofc"
            if df is not None:
                out = build_final_frame(df, year_hint=infer_year_from_filename(path))
                if out is not None: meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            return None, meta

        if ext == ".pdf":
            # 1) tentar tabelas
            frames = read_pdf_tables(path)
            for f in frames:
                out = build_final_frame(f, year_hint=infer_year_from_filename(path))
                if out is not None and not out.empty:
                    meta["metodo"]="pdf-tabela"; meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            # 2) texto linha-a-linha
            lines = pdf_text_lines(path)
            if lines:
                out = rows_from_text_lines("\n".join(lines))
                if out is not None and not out.empty:
                    meta["metodo"]="pdf-texto"; meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            # 3) OCR se pedido
            if ocr and pytesseract:
                texts = ocr_pdf_to_text(path)
                big = "\n".join(texts)
                out = rows_from_text_lines(big)
                if out is not None and not out.empty:
                    meta["metodo"]="pdf-ocr"; meta["ok"]=True; meta["linhas"]=len(out); return out, meta
            return None, meta

        meta["metodo"]="desconhecido"
        return None, meta

    except Exception as e:
        meta["erro"] = f"{type(e).__name__}: {e}"
        return None, meta

# ---------- IA opcional ----------
IA_PROMPT = """Você é um assistente contábil.
Receberá uma lista JSON de lançamentos no formato:
[{"Data":"YYYY-MM-DD","Descrição":"texto","Valor":numero,"Protocolo":"opcional"}]

Tarefa:
- Classifique cada linha em uma das categorias: ["Entrada","Saída","Tarifa","Transferência","Pix","Boleto","Cartão","Salário","Impostos","Outros"].
- Extraia "Contraparte" (nome de quem pagou/recebeu) quando houver.
- Gere "Nota" curta se algo não estiver claro.
Retorne APENAS JSON com a mesma quantidade de itens e os campos extras: Categoria, Contraparte, Nota.
"""

def ai_enrich(df: pd.DataFrame, batch_size: int=40, model: str="gpt-4o-mini") -> pd.DataFrame:
    """
    Tenta usar OpenAI (se OPENAI_API_KEY existir). Caso contrário, devolve heurístico.
    """
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key or not openai_mod:
        return heuristic_enrich(df)

    # OpenAI Python SDK v1
    try:
        client = openai_mod.OpenAI(api_key=api_key)
    except Exception:
        # se falhar instância, usa heurística
        return heuristic_enrich(df)

    rows = df.to_dict(orient="records")
    enriched = []
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i+batch_size]
        content = json.dumps(chunk, ensure_ascii=False)
        messages = [{"role":"system","content":IA_PROMPT},
                    {"role":"user","content":content}]
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.2,
                response_format={"type":"json_object"}
            )
            data = resp.choices[0].message.content
            j = json.loads(data)
            items = j.get("items", None)
            if items is None and isinstance(j, list): items = j
            if items is None: items = j
            if not isinstance(items, list): items = []
        except Exception:
            items = heuristic_enrich(pd.DataFrame(chunk)).to_dict(orient="records")

        # mescla base + extras
        for base, extra in zip(chunk, items):
            enriched.append({
                "Data": base.get("Data"),
                "Descrição": base.get("Descrição"),
                "Protocolo": base.get("Protocolo"),
                "Valor": base.get("Valor"),
                "Categoria": (extra.get("Categoria") or "Outros"),
                "Contraparte": extra.get("Contraparte"),
                "Nota": extra.get("Nota")
            })
    return pd.DataFrame(enriched)

def heuristic_enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Classificação simples local, sem IA, com extração de contraparte básica."""
    def cat(desc: str, valor: float) -> str:
        d = (desc or "").lower()
        if "pix" in d: return "Pix"
        if "boleto" in d or "c[oó]digo de barras" in d: return "Boleto"
        if "tarifa" in d or "taxa" in d or "cesta" in d: return "Tarifa"
        if "cart[aã]o" in d or "cr[eé]dito" in d or "d[eé]bito" in d: return "Cartão"
        if "sal[aá]rio" in d or "pro labore" in d or "pr[oó]-labore" in d: return "Salário"
        if "impost" in d or "dar" in d or "gps " in d or "gnre" in d or "darf" in d: return "Impostos"
        if any(k in d for k in ["transfer", "ted", "doc"]): return "Transferência"
        return "Entrada" if (valor is not None and valor > 0) else "Saída"

    def contraparte(desc: str) -> Optional[str]:
        if not desc: return None
        # pega bloco de letras/números “longo”
        m = re.search(r"([A-Za-zÀ-ú0-9][A-Za-zÀ-ú0-9\s\.\-&]{5,})", desc)
        if not m: return None
        c = m.group(1).strip()
        # remove ruídos comuns
        c = re.sub(r"(?i)\b(pix|boleto|tarifa|taxa|transfer[eê]ncia|ted|doc|cart[aã]o|visa|master|debit[o0]|credit[o0])\b","",c).strip()
        return c if len(c) >= 4 else None

    enriched = []
    for _, row in df.iterrows():
        dsc = str(row.get("Descrição") or "")
        val = row.get("Valor")
        enriched.append({
            "Data": row.get("Data"),
            "Descrição": dsc,
            "Protocolo": row.get("Protocolo"),
            "Valor": val,
            "Categoria": cat(dsc, val),
            "Contraparte": contraparte(dsc),
            "Nota": None
        })
    return pd.DataFrame(enriched)

# ---------- saída ----------
def save_outputs(df: pd.DataFrame, base_out: str, fmt: str="xlsx", also_txt: bool=True):
    os.makedirs(os.path.dirname(base_out) or ".", exist_ok=True)
    path_tab = f"{base_out}.{fmt.lower()}"
    if fmt.lower() == "xlsx":
        df.to_excel(path_tab, index=False)
    elif fmt.lower() == "csv":
        df.to_csv(path_tab, index=False, encoding="utf-8-sig", sep=";")
    elif fmt.lower() == "parquet":
        df.to_parquet(path_tab, index=False)
    else:
        raise ValueError("Formato inválido. Use: xlsx|csv|parquet")

    if also_txt:
        path_txt = f"{base_out}.txt"
        with open(path_txt, "w", encoding="utf-8") as f:
            for _, r in df.iterrows():
                linha = f'{r.get("Data")}\t{r.get("Descrição")}\t{r.get("Valor")}\t{r.get("Protocolo") or ""}'
                f.write(linha.strip()+"\n")
    return path_tab

# ---------- orquestração ----------
def process_path(input_path: Optional[str], ocr: bool, use_ai: bool, model: str, batch_size: int,
                 fmt: str, outdir: Optional[str], also_txt: bool=True) -> int:
    script_dir = get_script_dir()
    target = input_path or script_dir
    files=[]
    if os.path.isdir(target):
        for ext in ("*.pdf","*.csv","*.txt","*.xls","*.xlsx","*.ofx","*.ofc"):
            files.extend(glob.glob(os.path.join(target, ext)))
    elif os.path.isfile(target):
        files=[target]
    else:
        print(f"❌ Caminho não encontrado: {target}")
        return 2

    if not files:
        print("⚠️ Nenhum arquivo encontrado.")
        return 1

    metas=[]
    consolidados=[]
    for path in sorted(files):
        print(f"🔎 Processando: {os.path.basename(path)}")
        df, meta = process_file(path, ocr=ocr)
        if df is None or df.empty:
            print(f"   → Falha ({meta.get('metodo')}), linhas=0")
            metas.append(meta)
            continue
        # dedup e ordenação
        df = dedupe(df)
        try:
            d = pd.to_datetime(df["Data"], errors="coerce")
            df = df.assign(_d=d).sort_values("_d").drop(columns=["_d"])
        except: pass

        # enriquecimento (IA ou heurística)
        if use_ai:
            df_enr = ai_enrich(df, batch_size=batch_size, model=model)
            # garante ordem/colunas
            order = ["Data","Descrição","Valor","Protocolo","Categoria","Contraparte","Nota"]
            for c in order:
                if c not in df_enr.columns: df_enr[c]=None
            df_final = df_enr[order]
        else:
            df_final = heuristic_enrich(df)

        # salvar
        base_name = os.path.splitext(os.path.basename(path))[0] + "_limpo"
        out_base_dir = outdir or os.path.dirname(path) or "."
        base_out = os.path.join(out_base_dir, base_name)
        try:
            path_tab = save_outputs(df_final, base_out, fmt=fmt, also_txt=also_txt)
            print(f"   → OK: {meta.get('metodo')} | linhas={len(df_final)} | salvo: {os.path.basename(path_tab)}")
            meta["ok"]=True; meta["linhas"]=len(df_final)
            consolidados.append(df_final.assign(__arquivo=os.path.basename(path)))
        except Exception as e:
            meta["erro"] = f"Salvar: {type(e).__name__}: {e}"
            print(f"   → Erro ao salvar: {meta['erro']}")
        metas.append(meta)

    # consolidado (quando >1)
    if len(consolidados) > 1:
        dfc = pd.concat(consolidados, ignore_index=True)
        # ordena
        try:
            d = pd.to_datetime(dfc["Data"], errors="coerce")
            dfc = dfc.assign(_d=d).sort_values(["__arquivo","_d"]).drop(columns=["_d"])
        except: pass

        base_out = os.path.join(outdir or (input_path or script_dir), "consolidado")
        try:
            path_tab = save_outputs(dfc, base_out, fmt=fmt, also_txt=also_txt)
            print(f"📦 Consolidado: {os.path.basename(path_tab)} ({len(dfc)} linhas)")
        except Exception as e:
            print(f"❌ Erro ao salvar consolidado: {type(e).__name__}: {e}")

    # resumo
    ok = sum(1 for m in metas if m.get("ok"))
    fail = len(metas) - ok
    print("\n===== RESUMO =====")
    for m in metas:
        status = "OK" if m.get("ok") else "FALHA"
        extra = f" | linhas={m.get('linhas')}" if m.get("ok") else ""
        if m.get("erro"):
            extra += f" | erro={m.get('erro')}"
        print(f"- {m.get('arquivo')}: {status} ({m.get('metodo')}){extra}")
    print(f"Total: {len(metas)} | OK: {ok} | Falhas: {fail}")
    return 0 if ok>0 else 1

# ---------- main ----------
def main():
    parser = argparse.ArgumentParser(
        description="Extrai lançamentos de extratos (PDF/CSV/Excel/OFX) e gera arquivos limpos + TXT."
    )
    parser.add_argument("-i","--input", help="Arquivo ou pasta. Se omitido, usa a pasta do script.", default=None)
    parser.add_argument("--ocr", help="Força OCR em PDFs escaneados.", action="store_true")
    parser.add_argument("--format", help="Formato de saída tabular.", choices=["xlsx","csv","parquet"], default="xlsx")
    parser.add_argument("--outdir", help="Diretório de saída. Padrão: mesmo do arquivo de origem.", default=None)
    parser.add_argument("--no-txt", help="Não gerar arquivo .txt paralelo.", action="store_true")
    parser.add_argument("--no-ai", help="Não usar IA (sempre heurística local).", action="store_true")
    parser.add_argument("--model", help="Modelo OpenAI (se IA ativa).", default="gpt-4o-mini")
    parser.add_argument("--batch-size", type=int, default=40, help="Tamanho do lote para IA.")
    args = parser.parse_args()

    use_ai = not args.no_ai
    if use_ai and not os.getenv("OPENAI_API_KEY"):
        print("⚠️  IA ativada, mas OPENAI_API_KEY não está definida. Usarei heurística local.")
        use_ai = False

    rc = process_path(
        input_path=args.input,
        ocr=args.ocr,
        use_ai=use_ai,
        model=args.model,
        batch_size=args.batch_size,
        fmt=args.format,
        outdir=args.outdir,
        also_txt=not args.no_txt
    )
    sys.exit(rc)

if __name__ == "__main__":
    main()
