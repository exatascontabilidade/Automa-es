Attribute VB_Name = "clsRegistrosSPED"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Contrib As Boolean

' --- Funções de Carregamento do Bloco 0 ---

Public Sub CarregarDadosRegistro0000(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0000, reg0000, CamposChave)
    Set dtoTitSPED.t0000 = Util.MapearTitulos(reg0000, 3)
End Sub

Public Sub CarregarDadosRegistro0000_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0000_Contr, reg0000_Contr, CamposChave)
    Set dtoTitSPED.t0000_Contr = Util.MapearTitulos(reg0000_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro0001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0001, reg0001, CamposChave)
    Set dtoTitSPED.t0001 = Util.MapearTitulos(reg0001, 3)
End Sub

Public Sub CarregarDadosRegistro0002(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0002, reg0002, CamposChave)
    Set dtoTitSPED.t0002 = Util.MapearTitulos(reg0002, 3)
End Sub

Public Sub CarregarDadosRegistro0005(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0005, reg0005, CamposChave)
    Set dtoTitSPED.t0005 = Util.MapearTitulos(reg0005, 3)
End Sub

Public Sub CarregarDadosRegistro0015(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0015, reg0015, CamposChave)
    Set dtoTitSPED.t0015 = Util.MapearTitulos(reg0015, 3)
End Sub

Public Sub CarregarDadosRegistro0035(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0035, reg0035, CamposChave)
    Set dtoTitSPED.t0035 = Util.MapearTitulos(reg0035, 3)
End Sub

Public Sub CarregarDadosRegistro0100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0100, reg0100, CamposChave)
    Set dtoTitSPED.t0100 = Util.MapearTitulos(reg0100, 3)
End Sub

Public Sub CarregarDadosRegistro0110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0110, reg0110, CamposChave)
    Set dtoTitSPED.t0110 = Util.MapearTitulos(reg0110, 3)
End Sub

Public Sub CarregarDadosRegistro0111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0111, reg0111, CamposChave)
    Set dtoTitSPED.t0111 = Util.MapearTitulos(reg0111, 3)
End Sub

Public Sub CarregarDadosRegistro0120(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0120, reg0120, CamposChave)
    Set dtoTitSPED.t0120 = Util.MapearTitulos(reg0120, 3)
End Sub

Public Sub CarregarDadosRegistro0140(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0140, reg0140, CamposChave)
    Set dtoTitSPED.t0140 = Util.MapearTitulos(reg0140, 3)
End Sub

Public Sub CarregarDadosRegistro0145(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0145, reg0145, CamposChave)
    Set dtoTitSPED.t0145 = Util.MapearTitulos(reg0145, 3)
End Sub

Public Sub CarregarDadosRegistro0150(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0150, reg0150, CamposChave)
    Set dtoTitSPED.t0150 = Util.MapearTitulos(reg0150, 3)
End Sub

Public Sub CarregarDadosRegistro0175(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0175, reg0175, CamposChave)
    Set dtoTitSPED.t0175 = Util.MapearTitulos(reg0175, 3)
End Sub

Public Sub CarregarDadosRegistro0190(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0190, reg0190, CamposChave)
    Set dtoTitSPED.t0190 = Util.MapearTitulos(reg0190, 3)
End Sub

Public Sub CarregarDadosRegistro0200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0200, reg0200, CamposChave)
    Set dtoTitSPED.t0200 = Util.MapearTitulos(reg0200, 3)
End Sub

Public Sub CarregarDadosRegistro0205(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0205, reg0205, CamposChave)
    Set dtoTitSPED.t0205 = Util.MapearTitulos(reg0205, 3)
End Sub

Public Sub CarregarDadosRegistro0206(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0206, reg0206, CamposChave)
    Set dtoTitSPED.t0206 = Util.MapearTitulos(reg0206, 3)
End Sub

Public Sub CarregarDadosRegistro0208(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0208, reg0208, CamposChave)
    Set dtoTitSPED.t0208 = Util.MapearTitulos(reg0208, 3)
End Sub

Public Sub CarregarDadosRegistro0210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0210, reg0210, CamposChave)
    Set dtoTitSPED.t0210 = Util.MapearTitulos(reg0210, 3)
End Sub

Public Sub CarregarDadosRegistro0220(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0220, reg0220, CamposChave)
    Set dtoTitSPED.t0220 = Util.MapearTitulos(reg0220, 3)
End Sub

Public Sub CarregarDadosRegistro0221(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0221, reg0221, CamposChave)
    Set dtoTitSPED.t0221 = Util.MapearTitulos(reg0221, 3)
End Sub

Public Sub CarregarDadosRegistro0300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0300, reg0300, CamposChave)
    Set dtoTitSPED.t0300 = Util.MapearTitulos(reg0300, 3)
End Sub

Public Sub CarregarDadosRegistro0305(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0305, reg0305, CamposChave)
    Set dtoTitSPED.t0305 = Util.MapearTitulos(reg0305, 3)
End Sub

Public Sub CarregarDadosRegistro0400(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0400, reg0400, CamposChave)
    Set dtoTitSPED.t0400 = Util.MapearTitulos(reg0400, 3)
End Sub

Public Sub CarregarDadosRegistro0450(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0450, reg0450, CamposChave)
    Set dtoTitSPED.t0450 = Util.MapearTitulos(reg0450, 3)
End Sub

Public Sub CarregarDadosRegistro0460(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0460, reg0460, CamposChave)
    Set dtoTitSPED.t0460 = Util.MapearTitulos(reg0460, 3)
End Sub

Public Sub CarregarDadosRegistro0500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0500, reg0500, CamposChave)
    Set dtoTitSPED.t0500 = Util.MapearTitulos(reg0500, 3)
End Sub

Public Sub CarregarDadosRegistro0600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0600, reg0600, CamposChave)
    Set dtoTitSPED.t0600 = Util.MapearTitulos(reg0600, 3)
End Sub

Public Sub CarregarDadosRegistro0900(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0900, reg0900, CamposChave)
    Set dtoTitSPED.t0900 = Util.MapearTitulos(reg0900, 3)
End Sub

Public Sub CarregarDadosRegistro0990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r0990, reg0990, CamposChave)
    Set dtoTitSPED.t0990 = Util.MapearTitulos(reg0990, 3)
End Sub


' --- Funções de Carregamento do Bloco A ---

Public Sub CarregarDadosRegistroA001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA001, regA001, CamposChave)
    Set dtoTitSPED.tA001 = Util.MapearTitulos(regA001, 3)
End Sub

Public Sub CarregarDadosRegistroA010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA010, regA010, CamposChave)
    Set dtoTitSPED.tA010 = Util.MapearTitulos(regA010, 3)
End Sub

Public Sub CarregarDadosRegistroA100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA100, regA100, CamposChave)
    Set dtoTitSPED.tA100 = Util.MapearTitulos(regA100, 3)
End Sub

Public Sub CarregarDadosRegistroA110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA110, regA110, CamposChave)
    Set dtoTitSPED.tA110 = Util.MapearTitulos(regA110, 3)
End Sub

Public Sub CarregarDadosRegistroA111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA111, regA111, CamposChave)
    Set dtoTitSPED.tA111 = Util.MapearTitulos(regA111, 3)
End Sub

Public Sub CarregarDadosRegistroA120(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA120, regA120, CamposChave)
    Set dtoTitSPED.tA120 = Util.MapearTitulos(regA120, 3)
End Sub

Public Sub CarregarDadosRegistroA170(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA170, regA170, CamposChave)
    Set dtoTitSPED.tA170 = Util.MapearTitulos(regA170, 3)
End Sub

Public Sub CarregarDadosRegistroA990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rA990, regA990, CamposChave)
    Set dtoTitSPED.tA990 = Util.MapearTitulos(regA990, 3)
End Sub


' --- Funções de Carregamento do Bloco B ---

Public Sub CarregarDadosRegistroB001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB001, regB001, CamposChave)
    Set dtoTitSPED.tB001 = Util.MapearTitulos(regB001, 3)
End Sub

Public Sub CarregarDadosRegistroB020(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB020, regB020, CamposChave)
    Set dtoTitSPED.tB020 = Util.MapearTitulos(regB020, 3)
End Sub

Public Sub CarregarDadosRegistroB025(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB025, regB025, CamposChave)
    Set dtoTitSPED.tB025 = Util.MapearTitulos(regB025, 3)
End Sub

Public Sub CarregarDadosRegistroB030(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB030, regB030, CamposChave)
    Set dtoTitSPED.tB030 = Util.MapearTitulos(regB030, 3)
End Sub

Public Sub CarregarDadosRegistroB035(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB035, regB035, CamposChave)
    Set dtoTitSPED.tB035 = Util.MapearTitulos(regB035, 3)
End Sub

Public Sub CarregarDadosRegistroB350(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB350, regB350, CamposChave)
    Set dtoTitSPED.tB350 = Util.MapearTitulos(regB350, 3)
End Sub

Public Sub CarregarDadosRegistroB420(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB420, regB420, CamposChave)
    Set dtoTitSPED.tB420 = Util.MapearTitulos(regB420, 3)
End Sub

Public Sub CarregarDadosRegistroB440(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB440, regB440, CamposChave)
    Set dtoTitSPED.tB440 = Util.MapearTitulos(regB440, 3)
End Sub

Public Sub CarregarDadosRegistroB460(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB460, regB460, CamposChave)
    Set dtoTitSPED.tB460 = Util.MapearTitulos(regB460, 3)
End Sub

Public Sub CarregarDadosRegistroB470(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB470, regB470, CamposChave)
    Set dtoTitSPED.tB470 = Util.MapearTitulos(regB470, 3)
End Sub

Public Sub CarregarDadosRegistroB500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB500, regB500, CamposChave)
    Set dtoTitSPED.tB500 = Util.MapearTitulos(regB500, 3)
End Sub

Public Sub CarregarDadosRegistroB510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB510, regB510, CamposChave)
    Set dtoTitSPED.tB510 = Util.MapearTitulos(regB510, 3)
End Sub

Public Sub CarregarDadosRegistroB990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rB990, regB990, CamposChave)
    Set dtoTitSPED.tB990 = Util.MapearTitulos(regB990, 3)
End Sub


' --- Funções de Carregamento do Bloco C ---

Public Sub CarregarDadosRegistroC001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC001, regC001, CamposChave)
    Set dtoTitSPED.tC001 = Util.MapearTitulos(regC001, 3)
End Sub

Public Sub CarregarDadosRegistroC010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC010, regC010, CamposChave)
    Set dtoTitSPED.tC010 = Util.MapearTitulos(regC010, 3)
End Sub

Public Sub CarregarDadosRegistroC100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC100, regC100, CamposChave)
    Set dtoTitSPED.tC100 = Util.MapearTitulos(regC100, 3)
End Sub

Public Sub CarregarDadosRegistroC101(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC101, regC101, CamposChave)
    Set dtoTitSPED.tC101 = Util.MapearTitulos(regC101, 3)
End Sub

Public Sub CarregarDadosRegistroC105(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC105, regC105, CamposChave)
    Set dtoTitSPED.tC105 = Util.MapearTitulos(regC105, 3)
End Sub

Public Sub CarregarDadosRegistroC110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC110, regC110, CamposChave)
    Set dtoTitSPED.tC110 = Util.MapearTitulos(regC110, 3)
End Sub

Public Sub CarregarDadosRegistroC111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC111, regC111, CamposChave)
    Set dtoTitSPED.tC111 = Util.MapearTitulos(regC111, 3)
End Sub

Public Sub CarregarDadosRegistroC112(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC112, regC112, CamposChave)
    Set dtoTitSPED.tC112 = Util.MapearTitulos(regC112, 3)
End Sub

Public Sub CarregarDadosRegistroC113(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC113, regC113, CamposChave)
    Set dtoTitSPED.tC113 = Util.MapearTitulos(regC113, 3)
End Sub

Public Sub CarregarDadosRegistroC114(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC114, regC114, CamposChave)
    Set dtoTitSPED.tC114 = Util.MapearTitulos(regC114, 3)
End Sub

Public Sub CarregarDadosRegistroC115(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC115, regC115, CamposChave)
    Set dtoTitSPED.tC115 = Util.MapearTitulos(regC115, 3)
End Sub

Public Sub CarregarDadosRegistroC116(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC116, regC116, CamposChave)
    Set dtoTitSPED.tC116 = Util.MapearTitulos(regC116, 3)
End Sub

Public Sub CarregarDadosRegistroC120(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC120, regC120, CamposChave)
    Set dtoTitSPED.tC120 = Util.MapearTitulos(regC120, 3)
End Sub

Public Sub CarregarDadosRegistroC130(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC130, regC130, CamposChave)
    Set dtoTitSPED.tC130 = Util.MapearTitulos(regC130, 3)
End Sub

Public Sub CarregarDadosRegistroC140(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC140, regC140, CamposChave)
    Set dtoTitSPED.tC140 = Util.MapearTitulos(regC140, 3)
End Sub

Public Sub CarregarDadosRegistroC141(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC141, regC141, CamposChave)
    Set dtoTitSPED.tC141 = Util.MapearTitulos(regC141, 3)
End Sub

Public Sub CarregarDadosRegistroC160(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC160, regC160, CamposChave)
    Set dtoTitSPED.tC160 = Util.MapearTitulos(regC160, 3)
End Sub

Public Sub CarregarDadosRegistroC165(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC165, regC165, CamposChave)
    Set dtoTitSPED.tC165 = Util.MapearTitulos(regC165, 3)
End Sub

Public Sub CarregarDadosRegistroC170(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC170, regC170, CamposChave)
    Set dtoTitSPED.tC170 = Util.MapearTitulos(regC170, 3)
End Sub

Public Sub CarregarDadosRegistroC171(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC171, regC171, CamposChave)
    Set dtoTitSPED.tC171 = Util.MapearTitulos(regC171, 3)
End Sub

Public Sub CarregarDadosRegistroC172(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC172, regC172, CamposChave)
    Set dtoTitSPED.tC172 = Util.MapearTitulos(regC172, 3)
End Sub

Public Sub CarregarDadosRegistroC173(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC173, regC173, CamposChave)
    Set dtoTitSPED.tC173 = Util.MapearTitulos(regC173, 3)
End Sub

Public Sub CarregarDadosRegistroC174(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC174, regC174, CamposChave)
    Set dtoTitSPED.tC174 = Util.MapearTitulos(regC174, 3)
End Sub

Public Sub CarregarDadosRegistroC175(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC175, regC175, CamposChave)
    Set dtoTitSPED.tC175 = Util.MapearTitulos(regC175, 3)
End Sub

Public Sub CarregarDadosRegistroC175_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC175_Contr, regC175_Contr, CamposChave)
    Set dtoTitSPED.tC175_Contr = Util.MapearTitulos(regC175_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC176(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC176, regC176, CamposChave)
    Set dtoTitSPED.tC176 = Util.MapearTitulos(regC176, 3)
End Sub

Public Sub CarregarDadosRegistroC177(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC177, regC177, CamposChave)
    Set dtoTitSPED.tC177 = Util.MapearTitulos(regC177, 3)
End Sub

Public Sub CarregarDadosRegistroC178(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC178, regC178, CamposChave)
    Set dtoTitSPED.tC178 = Util.MapearTitulos(regC178, 3)
End Sub

Public Sub CarregarDadosRegistroC179(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC179, regC179, CamposChave)
    Set dtoTitSPED.tC179 = Util.MapearTitulos(regC179, 3)
End Sub

Public Sub CarregarDadosRegistroC180(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC180, regC180, CamposChave)
    Set dtoTitSPED.tC180 = Util.MapearTitulos(regC180, 3)
End Sub

Public Sub CarregarDadosRegistroC180_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC180_Contr, regC180_Contr, CamposChave)
    Set dtoTitSPED.tC180_Contr = Util.MapearTitulos(regC180_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC181(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC181, regC181, CamposChave)
    Set dtoTitSPED.tC181 = Util.MapearTitulos(regC181, 3)
End Sub

Public Sub CarregarDadosRegistroC181_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC181_Contr, regC181_Contr, CamposChave)
    Set dtoTitSPED.tC181_Contr = Util.MapearTitulos(regC181_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC185(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC185, regC185, CamposChave)
    Set dtoTitSPED.tC185 = Util.MapearTitulos(regC185, 3)
End Sub

Public Sub CarregarDadosRegistroC185_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC185_Contr, regC185_Contr, CamposChave)
    Set dtoTitSPED.tC185_Contr = Util.MapearTitulos(regC185_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC186(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC186, regC186, CamposChave)
    Set dtoTitSPED.tC186 = Util.MapearTitulos(regC186, 3)
End Sub

Public Sub CarregarDadosRegistroC188(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC188, regC188, CamposChave)
    Set dtoTitSPED.tC188 = Util.MapearTitulos(regC188, 3)
End Sub

Public Sub CarregarDadosRegistroC190(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC190, regC190, CamposChave)
    Set dtoTitSPED.tC190 = Util.MapearTitulos(regC190, 3)
End Sub

Public Sub CarregarDadosRegistroC190_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC190_Contr, regC190_Contr, CamposChave)
    Set dtoTitSPED.tC190_Contr = Util.MapearTitulos(regC190_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC191(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC191, regC191, CamposChave)
    Set dtoTitSPED.tC191 = Util.MapearTitulos(regC191, 3)
End Sub

Public Sub CarregarDadosRegistroC191_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC191_Contr, regC191_Contr, CamposChave)
    Set dtoTitSPED.tC191_Contr = Util.MapearTitulos(regC191_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC195(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC195, regC195, CamposChave)
    Set dtoTitSPED.tC195 = Util.MapearTitulos(regC195, 3)
End Sub

Public Sub CarregarDadosRegistroC195_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC195_Contr, regC195_Contr, CamposChave)
    Set dtoTitSPED.tC195_Contr = Util.MapearTitulos(regC195_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC197(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC197, regC197, CamposChave)
    Set dtoTitSPED.tC197 = Util.MapearTitulos(regC197, 3)
End Sub

Public Sub CarregarDadosRegistroC198(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC198, regC198, CamposChave)
    Set dtoTitSPED.tC198 = Util.MapearTitulos(regC198, 3)
End Sub

Public Sub CarregarDadosRegistroC199(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC199, regC199, CamposChave)
    Set dtoTitSPED.tC199 = Util.MapearTitulos(regC199, 3)
End Sub

Public Sub CarregarDadosRegistroC300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC300, regC300, CamposChave)
    Set dtoTitSPED.tC300 = Util.MapearTitulos(regC300, 3)
End Sub

Public Sub CarregarDadosRegistroC310(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC310, regC310, CamposChave)
    Set dtoTitSPED.tC310 = Util.MapearTitulos(regC310, 3)
End Sub

Public Sub CarregarDadosRegistroC320(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC320, regC320, CamposChave)
    Set dtoTitSPED.tC320 = Util.MapearTitulos(regC320, 3)
End Sub

Public Sub CarregarDadosRegistroC321(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC321, regC321, CamposChave)
    Set dtoTitSPED.tC321 = Util.MapearTitulos(regC321, 3)
End Sub

Public Sub CarregarDadosRegistroC330(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC330, regC330, CamposChave)
    Set dtoTitSPED.tC330 = Util.MapearTitulos(regC330, 3)
End Sub

Public Sub CarregarDadosRegistroC350(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC350, regC350, CamposChave)
    Set dtoTitSPED.tC350 = Util.MapearTitulos(regC350, 3)
End Sub

Public Sub CarregarDadosRegistroC370(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC370, regC370, CamposChave)
    Set dtoTitSPED.tC370 = Util.MapearTitulos(regC370, 3)
End Sub

Public Sub CarregarDadosRegistroC380(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC380, regC380, CamposChave)
    Set dtoTitSPED.tC380 = Util.MapearTitulos(regC380, 3)
End Sub

Public Sub CarregarDadosRegistroC380_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC380_Contr, regC380_Contr, CamposChave)
    Set dtoTitSPED.tC380_Contr = Util.MapearTitulos(regC380_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC381(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC381, regC381, CamposChave)
    Set dtoTitSPED.tC381 = Util.MapearTitulos(regC381, 3)
End Sub

Public Sub CarregarDadosRegistroC385(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC385, regC385, CamposChave)
    Set dtoTitSPED.tC385 = Util.MapearTitulos(regC385, 3)
End Sub

Public Sub CarregarDadosRegistroC390(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC390, regC390, CamposChave)
    Set dtoTitSPED.tC390 = Util.MapearTitulos(regC390, 3)
End Sub

Public Sub CarregarDadosRegistroC395(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC395, regC395, CamposChave)
    Set dtoTitSPED.tC395 = Util.MapearTitulos(regC395, 3)
End Sub

Public Sub CarregarDadosRegistroC396(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC396, regC396, CamposChave)
    Set dtoTitSPED.tC396 = Util.MapearTitulos(regC396, 3)
End Sub

Public Sub CarregarDadosRegistroC400(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC400, regC400, CamposChave)
    Set dtoTitSPED.tC400 = Util.MapearTitulos(regC400, 3)
End Sub

Public Sub CarregarDadosRegistroC405(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC405, regC405, CamposChave)
    Set dtoTitSPED.tC405 = Util.MapearTitulos(regC405, 3)
End Sub

Public Sub CarregarDadosRegistroC405_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC405_Contr, regC405_Contr, CamposChave)
    Set dtoTitSPED.tC405_Contr = Util.MapearTitulos(regC405_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC410(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC410, regC410, CamposChave)
    Set dtoTitSPED.tC410 = Util.MapearTitulos(regC410, 3)
End Sub

Public Sub CarregarDadosRegistroC420(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC420, regC420, CamposChave)
    Set dtoTitSPED.tC420 = Util.MapearTitulos(regC420, 3)
End Sub

Public Sub CarregarDadosRegistroC425(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC425, regC425, CamposChave)
    Set dtoTitSPED.tC425 = Util.MapearTitulos(regC425, 3)
End Sub

Public Sub CarregarDadosRegistroC430(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC430, regC430, CamposChave)
    Set dtoTitSPED.tC430 = Util.MapearTitulos(regC430, 3)
End Sub

Public Sub CarregarDadosRegistroC460(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC460, regC460, CamposChave)
    Set dtoTitSPED.tC460 = Util.MapearTitulos(regC460, 3)
End Sub

Public Sub CarregarDadosRegistroC465(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC465, regC465, CamposChave)
    Set dtoTitSPED.tC465 = Util.MapearTitulos(regC465, 3)
End Sub

Public Sub CarregarDadosRegistroC470(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC470, regC470, CamposChave)
    Set dtoTitSPED.tC470 = Util.MapearTitulos(regC470, 3)
End Sub

Public Sub CarregarDadosRegistroC480(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC480, regC480, CamposChave)
    Set dtoTitSPED.tC480 = Util.MapearTitulos(regC480, 3)
End Sub

Public Sub CarregarDadosRegistroC481(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC481, regC481, CamposChave)
    Set dtoTitSPED.tC481 = Util.MapearTitulos(regC481, 3)
End Sub

Public Sub CarregarDadosRegistroC485(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC485, regC485, CamposChave)
    Set dtoTitSPED.tC485 = Util.MapearTitulos(regC485, 3)
End Sub

Public Sub CarregarDadosRegistroC489(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC489, regC489, CamposChave)
    Set dtoTitSPED.tC489 = Util.MapearTitulos(regC489, 3)
End Sub

Public Sub CarregarDadosRegistroC490(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC490, regC490, CamposChave)
    Set dtoTitSPED.tC490 = Util.MapearTitulos(regC490, 3)
End Sub

Public Sub CarregarDadosRegistroC490_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC490_Contr, regC490_Contr, CamposChave)
    Set dtoTitSPED.tC490_Contr = Util.MapearTitulos(regC490_Contr, 3)
End Sub


Public Sub CarregarDadosRegistroC491(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC491, regC491, CamposChave)
    Set dtoTitSPED.tC491 = Util.MapearTitulos(regC491, 3)
End Sub

Public Sub CarregarDadosRegistroC495(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC495, regC495, CamposChave)
    Set dtoTitSPED.tC495 = Util.MapearTitulos(regC495, 3)
End Sub

Public Sub CarregarDadosRegistroC495_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC495_Contr, regC495_Contr, CamposChave)
    Set dtoTitSPED.tC495_Contr = Util.MapearTitulos(regC495_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC499(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC499, regC499, CamposChave)
    Set dtoTitSPED.tC499 = Util.MapearTitulos(regC499, 3)
End Sub

Public Sub CarregarDadosRegistroC500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC500, regC500, CamposChave)
    Set dtoTitSPED.tC500 = Util.MapearTitulos(regC500, 3)
End Sub

Public Sub CarregarDadosRegistroC500_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC500_Contr, regC500_Contr, CamposChave)
    Set dtoTitSPED.tC500_Contr = Util.MapearTitulos(regC500_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC501(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC501, regC501, CamposChave)
    Set dtoTitSPED.tC501 = Util.MapearTitulos(regC501, 3)
End Sub

Public Sub CarregarDadosRegistroC505(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC505, regC505, CamposChave)
    Set dtoTitSPED.tC505 = Util.MapearTitulos(regC505, 3)
End Sub

Public Sub CarregarDadosRegistroC509(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC509, regC509, CamposChave)
    Set dtoTitSPED.tC509 = Util.MapearTitulos(regC509, 3)
End Sub

Public Sub CarregarDadosRegistroC510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC510, regC510, CamposChave)
    Set dtoTitSPED.tC510 = Util.MapearTitulos(regC510, 3)
End Sub

Public Sub CarregarDadosRegistroC590(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC590, regC590, CamposChave)
    Set dtoTitSPED.tC590 = Util.MapearTitulos(regC590, 3)
End Sub

Public Sub CarregarDadosRegistroC591(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC591, regC591, CamposChave)
    Set dtoTitSPED.tC591 = Util.MapearTitulos(regC591, 3)
End Sub

Public Sub CarregarDadosRegistroC595(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC595, regC595, CamposChave)
    Set dtoTitSPED.tC595 = Util.MapearTitulos(regC595, 3)
End Sub

Public Sub CarregarDadosRegistroC597(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC597, regC597, CamposChave)
    Set dtoTitSPED.tC597 = Util.MapearTitulos(regC597, 3)
End Sub

Public Sub CarregarDadosRegistroC600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC600, regC600, CamposChave)
    Set dtoTitSPED.tC600 = Util.MapearTitulos(regC600, 3)
End Sub

Public Sub CarregarDadosRegistroC601(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC601, regC601, CamposChave)
    Set dtoTitSPED.tC601 = Util.MapearTitulos(regC601, 3)
End Sub

Public Sub CarregarDadosRegistroC601_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC601_Contr, regC601_Contr, CamposChave)
    Set dtoTitSPED.tC601_Contr = Util.MapearTitulos(regC601_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC605(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC605, regC605, CamposChave)
    Set dtoTitSPED.tC605 = Util.MapearTitulos(regC605, 3)
End Sub

Public Sub CarregarDadosRegistroC609(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC609, regC609, CamposChave)
    Set dtoTitSPED.tC609 = Util.MapearTitulos(regC609, 3)
End Sub

Public Sub CarregarDadosRegistroC610(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC610, regC610, CamposChave)
    Set dtoTitSPED.tC610 = Util.MapearTitulos(regC610, 3)
End Sub

Public Sub CarregarDadosRegistroC690(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC690, regC690, CamposChave)
    Set dtoTitSPED.tC690 = Util.MapearTitulos(regC690, 3)
End Sub

Public Sub CarregarDadosRegistroC700(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC700, regC700, CamposChave)
    Set dtoTitSPED.tC700 = Util.MapearTitulos(regC700, 3)
End Sub

Public Sub CarregarDadosRegistroC790(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC790, regC790, CamposChave)
    Set dtoTitSPED.tC790 = Util.MapearTitulos(regC790, 3)
End Sub

Public Sub CarregarDadosRegistroC791(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC791, regC791, CamposChave)
    Set dtoTitSPED.tC791 = Util.MapearTitulos(regC791, 3)
End Sub

Public Sub CarregarDadosRegistroC800(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC800, regC800, CamposChave)
    Set dtoTitSPED.tC800 = Util.MapearTitulos(regC800, 3)
End Sub

Public Sub CarregarDadosRegistroC810(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC810, regC810, CamposChave)
    Set dtoTitSPED.tC810 = Util.MapearTitulos(regC810, 3)
End Sub

Public Sub CarregarDadosRegistroC815(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC815, regC815, CamposChave)
    Set dtoTitSPED.tC815 = Util.MapearTitulos(regC815, 3)
End Sub

Public Sub CarregarDadosRegistroC820(ParamArray CamposChave() As Variant)
    'Descontinuado pelo SPED Contribuições
    'Call CarregarDadosRegistro(dtoRegSPED.rC820, regC820, CamposChave)
    'Set dtoTitSPED.tC820 = Util.MapearTitulos(regC820, 3)
End Sub

Public Sub CarregarDadosRegistroC830(ParamArray CamposChave() As Variant)
    'Descontinuado pelo SPED Contribuições
    'Call CarregarDadosRegistro(dtoRegSPED.rC830, regC830, CamposChave)
    'Set dtoTitSPED.tC830 = Util.MapearTitulos(regC830, 3)
End Sub

Public Sub CarregarDadosRegistroC850(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC850, regC850, CamposChave)
    Set dtoTitSPED.tC850 = Util.MapearTitulos(regC850, 3)
End Sub

Public Sub CarregarDadosRegistroC855(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC855, regC855, CamposChave)
    Set dtoTitSPED.tC855 = Util.MapearTitulos(regC855, 3)
End Sub

Public Sub CarregarDadosRegistroC857(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC857, regC857, CamposChave)
    Set dtoTitSPED.tC857 = Util.MapearTitulos(regC857, 3)
End Sub

Public Sub CarregarDadosRegistroC860(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC860, regC860, CamposChave)
    Set dtoTitSPED.tC860 = Util.MapearTitulos(regC860, 3)
End Sub

Public Sub CarregarDadosRegistroC870(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC870, regC870, CamposChave)
    Set dtoTitSPED.tC870 = Util.MapearTitulos(regC870, 3)
End Sub

Public Sub CarregarDadosRegistroC870_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC870_Contr, regC870_Contr, CamposChave)
    Set dtoTitSPED.tC870_Contr = Util.MapearTitulos(regC870_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC880(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC880, regC880, CamposChave)
    Set dtoTitSPED.tC880 = Util.MapearTitulos(regC880, 3)
End Sub

Public Sub CarregarDadosRegistroC880_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC880_Contr, regC880_Contr, CamposChave)
    Set dtoTitSPED.tC880_Contr = Util.MapearTitulos(regC880_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC890(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC890, regC890, CamposChave)
    Set dtoTitSPED.tC890 = Util.MapearTitulos(regC890, 3)
End Sub

Public Sub CarregarDadosRegistroC890_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC890_Contr, regC890_Contr, CamposChave)
    Set dtoTitSPED.tC890_Contr = Util.MapearTitulos(regC890_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroC895(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC895, regC895, CamposChave)
    Set dtoTitSPED.tC895 = Util.MapearTitulos(regC895, 3)
End Sub

Public Sub CarregarDadosRegistroC897(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC897, regC897, CamposChave)
    Set dtoTitSPED.tC897 = Util.MapearTitulos(regC897, 3)
End Sub

Public Sub CarregarDadosRegistroC990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rC990, regC990, CamposChave)
    Set dtoTitSPED.tC990 = Util.MapearTitulos(regC990, 3)
End Sub

' --- Funções de Carregamento do Bloco D ---

Public Sub CarregarDadosRegistroD001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD001, regD001, CamposChave)
    Set dtoTitSPED.tD001 = Util.MapearTitulos(regD001, 3)
End Sub

Public Sub CarregarDadosRegistroD010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD010, regD010, CamposChave)
    Set dtoTitSPED.tD010 = Util.MapearTitulos(regD010, 3)
End Sub

Public Sub CarregarDadosRegistroD100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD100, regD100, CamposChave)
    Set dtoTitSPED.tD100 = Util.MapearTitulos(regD100, 3)
End Sub

Public Sub CarregarDadosRegistroD101_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD101_Contr, regD101_Contr, CamposChave)
    Set dtoTitSPED.tD101_Contr = Util.MapearTitulos(regD101_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroD101(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD101, regD101, CamposChave)
    Set dtoTitSPED.tD101 = Util.MapearTitulos(regD101, 3)
End Sub

Public Sub CarregarDadosRegistroD105(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD105, regD105, CamposChave)
    Set dtoTitSPED.tD105 = Util.MapearTitulos(regD105, 3)
End Sub

Public Sub CarregarDadosRegistroD110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD110, regD110, CamposChave)
    Set dtoTitSPED.tD110 = Util.MapearTitulos(regD110, 3)
End Sub

Public Sub CarregarDadosRegistroD111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD111, regD111, CamposChave)
    Set dtoTitSPED.tD111 = Util.MapearTitulos(regD111, 3)
End Sub

Public Sub CarregarDadosRegistroD120(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD120, regD120, CamposChave)
    Set dtoTitSPED.tD120 = Util.MapearTitulos(regD120, 3)
End Sub

Public Sub CarregarDadosRegistroD130(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD130, regD130, CamposChave)
    Set dtoTitSPED.tD130 = Util.MapearTitulos(regD130, 3)
End Sub

Public Sub CarregarDadosRegistroD140(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD140, regD140, CamposChave)
    Set dtoTitSPED.tD140 = Util.MapearTitulos(regD140, 3)
End Sub

Public Sub CarregarDadosRegistroD150(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD150, regD150, CamposChave)
    Set dtoTitSPED.tD150 = Util.MapearTitulos(regD150, 3)
End Sub

Public Sub CarregarDadosRegistroD160(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD160, regD160, CamposChave)
    Set dtoTitSPED.tD160 = Util.MapearTitulos(regD160, 3)
End Sub

Public Sub CarregarDadosRegistroD161(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD161, regD161, CamposChave)
    Set dtoTitSPED.tD161 = Util.MapearTitulos(regD161, 3)
End Sub

Public Sub CarregarDadosRegistroD162(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD162, regD162, CamposChave)
    Set dtoTitSPED.tD162 = Util.MapearTitulos(regD162, 3)
End Sub

Public Sub CarregarDadosRegistroD170(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD170, regD170, CamposChave)
    Set dtoTitSPED.tD170 = Util.MapearTitulos(regD170, 3)
End Sub

Public Sub CarregarDadosRegistroD180(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD180, regD180, CamposChave)
    Set dtoTitSPED.tD180 = Util.MapearTitulos(regD180, 3)
End Sub

Public Sub CarregarDadosRegistroD190(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD190, regD190, CamposChave)
    Set dtoTitSPED.tD190 = Util.MapearTitulos(regD190, 3)
End Sub

Public Sub CarregarDadosRegistroD195(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD195, regD195, CamposChave)
    Set dtoTitSPED.tD195 = Util.MapearTitulos(regD195, 3)
End Sub

Public Sub CarregarDadosRegistroD197(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD197, regD197, CamposChave)
    Set dtoTitSPED.tD197 = Util.MapearTitulos(regD197, 3)
End Sub

Public Sub CarregarDadosRegistroD200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD200, regD200, CamposChave)
    Set dtoTitSPED.tD200 = Util.MapearTitulos(regD200, 3)
End Sub

Public Sub CarregarDadosRegistroD201(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD201, regD201, CamposChave)
    Set dtoTitSPED.tD201 = Util.MapearTitulos(regD201, 3)
End Sub

Public Sub CarregarDadosRegistroD205(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD205, regD205, CamposChave)
    Set dtoTitSPED.tD205 = Util.MapearTitulos(regD205, 3)
End Sub

Public Sub CarregarDadosRegistroD209(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD209, regD209, CamposChave)
    Set dtoTitSPED.tD209 = Util.MapearTitulos(regD209, 3)
End Sub

Public Sub CarregarDadosRegistroD300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD300, regD300, CamposChave)
    Set dtoTitSPED.tD300 = Util.MapearTitulos(regD300, 3)
End Sub

Public Sub CarregarDadosRegistroD300_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD300_Contr, regD300_Contr, CamposChave)
    Set dtoTitSPED.tD300_Contr = Util.MapearTitulos(regD300_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroD301(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD301, regD301, CamposChave)
    Set dtoTitSPED.tD301 = Util.MapearTitulos(regD301, 3)
End Sub

Public Sub CarregarDadosRegistroD309(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD309, regD309, CamposChave)
    Set dtoTitSPED.tD309 = Util.MapearTitulos(regD309, 3)
End Sub

Public Sub CarregarDadosRegistroD310(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD310, regD310, CamposChave)
    Set dtoTitSPED.tD310 = Util.MapearTitulos(regD310, 3)
End Sub

Public Sub CarregarDadosRegistroD350(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD350, regD350, CamposChave)
    Set dtoTitSPED.tD350 = Util.MapearTitulos(regD350, 3)
End Sub

Public Sub CarregarDadosRegistroD350_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD350_Contr, regD350_Contr, CamposChave)
    Set dtoTitSPED.tD350_Contr = Util.MapearTitulos(regD350_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroD355(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD355, regD355, CamposChave)
    Set dtoTitSPED.tD355 = Util.MapearTitulos(regD355, 3)
End Sub

Public Sub CarregarDadosRegistroD359(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD359, regD359, CamposChave)
    Set dtoTitSPED.tD359 = Util.MapearTitulos(regD359, 3)
End Sub

Public Sub CarregarDadosRegistroD360(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD360, regD360, CamposChave)
    Set dtoTitSPED.tD360 = Util.MapearTitulos(regD360, 3)
End Sub

Public Sub CarregarDadosRegistroD365(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD365, regD365, CamposChave)
    Set dtoTitSPED.tD365 = Util.MapearTitulos(regD365, 3)
End Sub

Public Sub CarregarDadosRegistroD370(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD370, regD370, CamposChave)
    Set dtoTitSPED.tD370 = Util.MapearTitulos(regD370, 3)
End Sub

Public Sub CarregarDadosRegistroD390(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD390, regD390, CamposChave)
    Set dtoTitSPED.tD390 = Util.MapearTitulos(regD390, 3)
End Sub

Public Sub CarregarDadosRegistroD400(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD400, regD400, CamposChave)
    Set dtoTitSPED.tD400 = Util.MapearTitulos(regD400, 3)
End Sub

Public Sub CarregarDadosRegistroD410(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD410, regD410, CamposChave)
    Set dtoTitSPED.tD410 = Util.MapearTitulos(regD410, 3)
End Sub

Public Sub CarregarDadosRegistroD411(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD411, regD411, CamposChave)
    Set dtoTitSPED.tD411 = Util.MapearTitulos(regD411, 3)
End Sub

Public Sub CarregarDadosRegistroD420(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD420, regD420, CamposChave)
    Set dtoTitSPED.tD420 = Util.MapearTitulos(regD420, 3)
End Sub

Public Sub CarregarDadosRegistroD500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD500, regD500, CamposChave)
    Set dtoTitSPED.tD500 = Util.MapearTitulos(regD500, 3)
End Sub

Public Sub CarregarDadosRegistroD501(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD501, regD501, CamposChave)
    Set dtoTitSPED.tD501 = Util.MapearTitulos(regD501, 3)
End Sub

Public Sub CarregarDadosRegistroD505(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD505, regD505, CamposChave)
    Set dtoTitSPED.tD505 = Util.MapearTitulos(regD505, 3)
End Sub

Public Sub CarregarDadosRegistroD509(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD509, regD509, CamposChave)
    Set dtoTitSPED.tD509 = Util.MapearTitulos(regD509, 3)
End Sub

Public Sub CarregarDadosRegistroD510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD510, regD510, CamposChave)
    Set dtoTitSPED.tD510 = Util.MapearTitulos(regD510, 3)
End Sub

Public Sub CarregarDadosRegistroD530(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD530, regD530, CamposChave)
    Set dtoTitSPED.tD530 = Util.MapearTitulos(regD530, 3)
End Sub

Public Sub CarregarDadosRegistroD590(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD590, regD590, CamposChave)
    Set dtoTitSPED.tD590 = Util.MapearTitulos(regD590, 3)
End Sub

Public Sub CarregarDadosRegistroD600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD600, regD600, CamposChave)
    Set dtoTitSPED.tD600 = Util.MapearTitulos(regD600, 3)
End Sub

Public Sub CarregarDadosRegistroD600_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD600_Contr, regD600_Contr, CamposChave)
    Set dtoTitSPED.tD600_Contr = Util.MapearTitulos(regD600_Contr, 3)
End Sub

Public Sub CarregarDadosRegistroD601(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD601, regD601, CamposChave)
    Set dtoTitSPED.tD601 = Util.MapearTitulos(regD601, 3)
End Sub

Public Sub CarregarDadosRegistroD605(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD605, regD605, CamposChave)
    Set dtoTitSPED.tD605 = Util.MapearTitulos(regD605, 3)
End Sub

Public Sub CarregarDadosRegistroD609(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD609, regD609, CamposChave)
    Set dtoTitSPED.tD609 = Util.MapearTitulos(regD609, 3)
End Sub

Public Sub CarregarDadosRegistroD610(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD610, regD610, CamposChave)
    Set dtoTitSPED.tD610 = Util.MapearTitulos(regD610, 3)
End Sub

Public Sub CarregarDadosRegistroD690(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD690, regD690, CamposChave)
    Set dtoTitSPED.tD690 = Util.MapearTitulos(regD690, 3)
End Sub

Public Sub CarregarDadosRegistroD695(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD695, regD695, CamposChave)
    Set dtoTitSPED.tD695 = Util.MapearTitulos(regD695, 3)
End Sub

Public Sub CarregarDadosRegistroD696(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD696, regD696, CamposChave)
    Set dtoTitSPED.tD696 = Util.MapearTitulos(regD696, 3)
End Sub

Public Sub CarregarDadosRegistroD697(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD697, regD697, CamposChave)
    Set dtoTitSPED.tD697 = Util.MapearTitulos(regD697, 3)
End Sub

Public Sub CarregarDadosRegistroD700(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD700, regD700, CamposChave)
    Set dtoTitSPED.tD700 = Util.MapearTitulos(regD700, 3)
End Sub

Public Sub CarregarDadosRegistroD730(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD730, regD730, CamposChave)
    Set dtoTitSPED.tD730 = Util.MapearTitulos(regD730, 3)
End Sub

Public Sub CarregarDadosRegistroD731(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD731, regD731, CamposChave)
    Set dtoTitSPED.tD731 = Util.MapearTitulos(regD731, 3)
End Sub

Public Sub CarregarDadosRegistroD735(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD735, regD735, CamposChave)
    Set dtoTitSPED.tD735 = Util.MapearTitulos(regD735, 3)
End Sub

Public Sub CarregarDadosRegistroD737(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD737, regD737, CamposChave)
    Set dtoTitSPED.tD737 = Util.MapearTitulos(regD737, 3)
End Sub

Public Sub CarregarDadosRegistroD750(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD750, regD750, CamposChave)
    Set dtoTitSPED.tD750 = Util.MapearTitulos(regD750, 3)
End Sub

Public Sub CarregarDadosRegistroD760(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD760, regD760, CamposChave)
    Set dtoTitSPED.tD760 = Util.MapearTitulos(regD760, 3)
End Sub

Public Sub CarregarDadosRegistroD761(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD761, regD761, CamposChave)
    Set dtoTitSPED.tD761 = Util.MapearTitulos(regD761, 3)
End Sub

Public Sub CarregarDadosRegistroD990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rD990, regD990, CamposChave)
    Set dtoTitSPED.tD990 = Util.MapearTitulos(regD990, 3)
End Sub


' --- Funções de Carregamento do Bloco E ---

Public Sub CarregarDadosRegistroE001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE001, regE001, CamposChave)
    Set dtoTitSPED.tE001 = Util.MapearTitulos(regE001, 3)
End Sub

Public Sub CarregarDadosRegistroE100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE100, regE100, CamposChave)
    Set dtoTitSPED.tE100 = Util.MapearTitulos(regE100, 3)
End Sub

Public Sub CarregarDadosRegistroE110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE110, regE110, CamposChave)
    Set dtoTitSPED.tE110 = Util.MapearTitulos(regE110, 3)
End Sub

Public Sub CarregarDadosRegistroE111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE111, regE111, CamposChave)
    Set dtoTitSPED.tE111 = Util.MapearTitulos(regE111, 3)
End Sub

Public Sub CarregarDadosRegistroE112(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE112, regE112, CamposChave)
    Set dtoTitSPED.tE112 = Util.MapearTitulos(regE112, 3)
End Sub

Public Sub CarregarDadosRegistroE113(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE113, regE113, CamposChave)
    Set dtoTitSPED.tE113 = Util.MapearTitulos(regE113, 3)
End Sub

Public Sub CarregarDadosRegistroE115(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE115, regE115, CamposChave)
    Set dtoTitSPED.tE115 = Util.MapearTitulos(regE115, 3)
End Sub

Public Sub CarregarDadosRegistroE116(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE116, regE116, CamposChave)
    Set dtoTitSPED.tE116 = Util.MapearTitulos(regE116, 3)
End Sub

Public Sub CarregarDadosRegistroE200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE200, regE200, CamposChave)
    Set dtoTitSPED.tE200 = Util.MapearTitulos(regE200, 3)
End Sub

Public Sub CarregarDadosRegistroE210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE210, regE210, CamposChave)
    Set dtoTitSPED.tE210 = Util.MapearTitulos(regE210, 3)
End Sub

Public Sub CarregarDadosRegistroE220(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE220, regE220, CamposChave)
    Set dtoTitSPED.tE220 = Util.MapearTitulos(regE220, 3)
End Sub

Public Sub CarregarDadosRegistroE230(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE230, regE230, CamposChave)
    Set dtoTitSPED.tE230 = Util.MapearTitulos(regE230, 3)
End Sub

Public Sub CarregarDadosRegistroE240(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE240, regE240, CamposChave)
    Set dtoTitSPED.tE240 = Util.MapearTitulos(regE240, 3)
End Sub

Public Sub CarregarDadosRegistroE250(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE250, regE250, CamposChave)
    Set dtoTitSPED.tE250 = Util.MapearTitulos(regE250, 3)
End Sub

Public Sub CarregarDadosRegistroE300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE300, regE300, CamposChave)
    Set dtoTitSPED.tE300 = Util.MapearTitulos(regE300, 3)
End Sub

Public Sub CarregarDadosRegistroE310(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE310, regE310, CamposChave)
    Set dtoTitSPED.tE310 = Util.MapearTitulos(regE310, 3)
End Sub

Public Sub CarregarDadosRegistroE311(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE311, regE311, CamposChave)
    Set dtoTitSPED.tE311 = Util.MapearTitulos(regE311, 3)
End Sub

Public Sub CarregarDadosRegistroE312(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE312, regE312, CamposChave)
    Set dtoTitSPED.tE312 = Util.MapearTitulos(regE312, 3)
End Sub

Public Sub CarregarDadosRegistroE313(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE313, regE313, CamposChave)
    Set dtoTitSPED.tE313 = Util.MapearTitulos(regE313, 3)
End Sub

Public Sub CarregarDadosRegistroE316(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE316, regE316, CamposChave)
    Set dtoTitSPED.tE316 = Util.MapearTitulos(regE316, 3)
End Sub

Public Sub CarregarDadosRegistroE500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE500, regE500, CamposChave)
    Set dtoTitSPED.tE500 = Util.MapearTitulos(regE500, 3)
End Sub

Public Sub CarregarDadosRegistroE510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE510, regE510, CamposChave)
    Set dtoTitSPED.tE510 = Util.MapearTitulos(regE510, 3)
End Sub

Public Sub CarregarDadosRegistroE520(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE520, regE520, CamposChave)
    Set dtoTitSPED.tE520 = Util.MapearTitulos(regE520, 3)
End Sub

Public Sub CarregarDadosRegistroE530(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE530, regE530, CamposChave)
    Set dtoTitSPED.tE530 = Util.MapearTitulos(regE530, 3)
End Sub

Public Sub CarregarDadosRegistroE531(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE531, regE531, CamposChave)
    Set dtoTitSPED.tE531 = Util.MapearTitulos(regE531, 3)
End Sub

Public Sub CarregarDadosRegistroE990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rE990, regE990, CamposChave)
    Set dtoTitSPED.tE990 = Util.MapearTitulos(regE990, 3)
End Sub


' --- Funções de Carregamento do Bloco F ---

Public Sub CarregarDadosRegistroF001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF001, regF001, CamposChave)
    Set dtoTitSPED.tF001 = Util.MapearTitulos(regF001, 3)
End Sub

Public Sub CarregarDadosRegistroF010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF010, regF010, CamposChave)
    Set dtoTitSPED.tF010 = Util.MapearTitulos(regF010, 3)
End Sub

Public Sub CarregarDadosRegistroF100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF100, regF100, CamposChave)
    Set dtoTitSPED.tF100 = Util.MapearTitulos(regF100, 3)
End Sub

Public Sub CarregarDadosRegistroF111(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF111, regF111, CamposChave)
    Set dtoTitSPED.tF111 = Util.MapearTitulos(regF111, 3)
End Sub

Public Sub CarregarDadosRegistroF120(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF120, regF120, CamposChave)
    Set dtoTitSPED.tF120 = Util.MapearTitulos(regF120, 3)
End Sub

Public Sub CarregarDadosRegistroF129(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF129, regF129, CamposChave)
    Set dtoTitSPED.tF129 = Util.MapearTitulos(regF129, 3)
End Sub

Public Sub CarregarDadosRegistroF130(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF130, regF130, CamposChave)
    Set dtoTitSPED.tF130 = Util.MapearTitulos(regF130, 3)
End Sub

Public Sub CarregarDadosRegistroF139(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF139, regF139, CamposChave)
    Set dtoTitSPED.tF139 = Util.MapearTitulos(regF139, 3)
End Sub

Public Sub CarregarDadosRegistroF150(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF150, regF150, CamposChave)
    Set dtoTitSPED.tF150 = Util.MapearTitulos(regF150, 3)
End Sub

Public Sub CarregarDadosRegistroF200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF200, regF200, CamposChave)
    Set dtoTitSPED.tF200 = Util.MapearTitulos(regF200, 3)
End Sub

Public Sub CarregarDadosRegistroF205(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF205, regF205, CamposChave)
    Set dtoTitSPED.tF205 = Util.MapearTitulos(regF205, 3)
End Sub

Public Sub CarregarDadosRegistroF210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF210, regF210, CamposChave)
    Set dtoTitSPED.tF210 = Util.MapearTitulos(regF210, 3)
End Sub

Public Sub CarregarDadosRegistroF211(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF211, regF211, CamposChave)
    Set dtoTitSPED.tF211 = Util.MapearTitulos(regF211, 3)
End Sub

Public Sub CarregarDadosRegistroF500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF500, regF500, CamposChave)
    Set dtoTitSPED.tF500 = Util.MapearTitulos(regF500, 3)
End Sub

Public Sub CarregarDadosRegistroF509(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF509, regF509, CamposChave)
    Set dtoTitSPED.tF509 = Util.MapearTitulos(regF509, 3)
End Sub

Public Sub CarregarDadosRegistroF510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF510, regF510, CamposChave)
    Set dtoTitSPED.tF510 = Util.MapearTitulos(regF510, 3)
End Sub

Public Sub CarregarDadosRegistroF519(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF519, regF519, CamposChave)
    Set dtoTitSPED.tF519 = Util.MapearTitulos(regF519, 3)
End Sub

Public Sub CarregarDadosRegistroF525(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF525, regF525, CamposChave)
    Set dtoTitSPED.tF525 = Util.MapearTitulos(regF525, 3)
End Sub

Public Sub CarregarDadosRegistroF550(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF550, regF550, CamposChave)
    Set dtoTitSPED.tF550 = Util.MapearTitulos(regF550, 3)
End Sub

Public Sub CarregarDadosRegistroF559(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF559, regF559, CamposChave)
    Set dtoTitSPED.tF559 = Util.MapearTitulos(regF559, 3)
End Sub

Public Sub CarregarDadosRegistroF560(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF560, regF560, CamposChave)
    Set dtoTitSPED.tF560 = Util.MapearTitulos(regF560, 3)
End Sub

Public Sub CarregarDadosRegistroF569(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF569, regF569, CamposChave)
    Set dtoTitSPED.tF569 = Util.MapearTitulos(regF569, 3)
End Sub

Public Sub CarregarDadosRegistroF600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF600, regF600, CamposChave)
    Set dtoTitSPED.tF600 = Util.MapearTitulos(regF600, 3)
End Sub

Public Sub CarregarDadosRegistroF700(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF700, regF700, CamposChave)
    Set dtoTitSPED.tF700 = Util.MapearTitulos(regF700, 3)
End Sub

Public Sub CarregarDadosRegistroF800(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF800, regF800, CamposChave)
    Set dtoTitSPED.tF800 = Util.MapearTitulos(regF800, 3)
End Sub

Public Sub CarregarDadosRegistroF990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rF990, regF990, CamposChave)
    Set dtoTitSPED.tF990 = Util.MapearTitulos(regF990, 3)
End Sub


' --- Funções de Carregamento do Bloco G ---

Public Sub CarregarDadosRegistroG001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG001, regG001, CamposChave)
    Set dtoTitSPED.tG001 = Util.MapearTitulos(regG001, 3)
End Sub

Public Sub CarregarDadosRegistroG110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG110, regG110, CamposChave)
    Set dtoTitSPED.tG110 = Util.MapearTitulos(regG110, 3)
End Sub

Public Sub CarregarDadosRegistroG125(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG125, regG125, CamposChave)
    Set dtoTitSPED.tG125 = Util.MapearTitulos(regG125, 3)
End Sub

Public Sub CarregarDadosRegistroG126(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG126, regG126, CamposChave)
    Set dtoTitSPED.tG126 = Util.MapearTitulos(regG126, 3)
End Sub

Public Sub CarregarDadosRegistroG130(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG130, regG130, CamposChave)
    Set dtoTitSPED.tG130 = Util.MapearTitulos(regG130, 3)
End Sub

Public Sub CarregarDadosRegistroG140(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG140, regG140, CamposChave)
    Set dtoTitSPED.tG140 = Util.MapearTitulos(regG140, 3)
End Sub

Public Sub CarregarDadosRegistroG990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rG990, regG990, CamposChave)
    Set dtoTitSPED.tG990 = Util.MapearTitulos(regG990, 3)
End Sub


' --- Funções de Carregamento do Bloco H ---

Public Sub CarregarDadosRegistroH001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH001, regH001, CamposChave)
    Set dtoTitSPED.tH001 = Util.MapearTitulos(regH001, 3)
End Sub

Public Sub CarregarDadosRegistroH005(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH005, regH005, CamposChave)
    Set dtoTitSPED.tH005 = Util.MapearTitulos(regH005, 3)
End Sub

Public Sub CarregarDadosRegistroH010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH010, regH010, CamposChave)
    Set dtoTitSPED.tH010 = Util.MapearTitulos(regH010, 3)
End Sub

Public Sub CarregarDadosRegistroH020(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH020, regH020, CamposChave)
    Set dtoTitSPED.tH020 = Util.MapearTitulos(regH020, 3)
End Sub

Public Sub CarregarDadosRegistroH030(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH030, regH030, CamposChave)
    Set dtoTitSPED.tH030 = Util.MapearTitulos(regH030, 3)
End Sub

Public Sub CarregarDadosRegistroH990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rH990, regH990, CamposChave)
    Set dtoTitSPED.tH990 = Util.MapearTitulos(regH990, 3)
End Sub


' --- Funções de Carregamento do Bloco I ---

Public Sub CarregarDadosRegistroI001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI001, regI001, CamposChave)
    Set dtoTitSPED.tI001 = Util.MapearTitulos(regI001, 3)
End Sub

Public Sub CarregarDadosRegistroI010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI010, regI010, CamposChave)
    Set dtoTitSPED.tI010 = Util.MapearTitulos(regI010, 3)
End Sub

Public Sub CarregarDadosRegistroI100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI100, regI100, CamposChave)
    Set dtoTitSPED.tI100 = Util.MapearTitulos(regI100, 3)
End Sub

Public Sub CarregarDadosRegistroI199(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI199, regI199, CamposChave)
    Set dtoTitSPED.tI199 = Util.MapearTitulos(regI199, 3)
End Sub

Public Sub CarregarDadosRegistroI200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI200, regI200, CamposChave)
    Set dtoTitSPED.tI200 = Util.MapearTitulos(regI200, 3)
End Sub

Public Sub CarregarDadosRegistroI299(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI299, regI299, CamposChave)
    Set dtoTitSPED.tI299 = Util.MapearTitulos(regI299, 3)
End Sub

Public Sub CarregarDadosRegistroI300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI300, regI300, CamposChave)
    Set dtoTitSPED.tI300 = Util.MapearTitulos(regI300, 3)
End Sub

Public Sub CarregarDadosRegistroI399(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI399, regI399, CamposChave)
    Set dtoTitSPED.tI399 = Util.MapearTitulos(regI399, 3)
End Sub

Public Sub CarregarDadosRegistroI990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rI990, regI990, CamposChave)
    Set dtoTitSPED.tI990 = Util.MapearTitulos(regI990, 3)
End Sub


' --- Funções de Carregamento do Bloco K ---

Public Sub CarregarDadosRegistroK001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK001, regK001, CamposChave)
    Set dtoTitSPED.tK001 = Util.MapearTitulos(regK001, 3)
End Sub

Public Sub CarregarDadosRegistroK010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK010, regK010, CamposChave)
    Set dtoTitSPED.tK010 = Util.MapearTitulos(regK010, 3)
End Sub

Public Sub CarregarDadosRegistroK100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK100, regK100, CamposChave)
    Set dtoTitSPED.tK100 = Util.MapearTitulos(regK100, 3)
End Sub

Public Sub CarregarDadosRegistroK200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK200, regK200, CamposChave)
    Set dtoTitSPED.tK200 = Util.MapearTitulos(regK200, 3)
End Sub

Public Sub CarregarDadosRegistroK210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK210, regK210, CamposChave)
    Set dtoTitSPED.tK210 = Util.MapearTitulos(regK210, 3)
End Sub

Public Sub CarregarDadosRegistroK215(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK215, regK215, CamposChave)
    Set dtoTitSPED.tK215 = Util.MapearTitulos(regK215, 3)
End Sub

Public Sub CarregarDadosRegistroK220(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK220, regK220, CamposChave)
    Set dtoTitSPED.tK220 = Util.MapearTitulos(regK220, 3)
End Sub

Public Sub CarregarDadosRegistroK230(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK230, regK230, CamposChave)
    Set dtoTitSPED.tK230 = Util.MapearTitulos(regK230, 3)
End Sub

Public Sub CarregarDadosRegistroK235(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK235, regK235, CamposChave)
    Set dtoTitSPED.tK235 = Util.MapearTitulos(regK235, 3)
End Sub

Public Sub CarregarDadosRegistroK250(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK250, regK250, CamposChave)
    Set dtoTitSPED.tK250 = Util.MapearTitulos(regK250, 3)
End Sub

Public Sub CarregarDadosRegistroK255(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK255, regK255, CamposChave)
    Set dtoTitSPED.tK255 = Util.MapearTitulos(regK255, 3)
End Sub

Public Sub CarregarDadosRegistroK260(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK260, regK260, CamposChave)
    Set dtoTitSPED.tK260 = Util.MapearTitulos(regK260, 3)
End Sub

Public Sub CarregarDadosRegistroK265(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK265, regK265, CamposChave)
    Set dtoTitSPED.tK265 = Util.MapearTitulos(regK265, 3)
End Sub

Public Sub CarregarDadosRegistroK270(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK270, regK270, CamposChave)
    Set dtoTitSPED.tK270 = Util.MapearTitulos(regK270, 3)
End Sub

Public Sub CarregarDadosRegistroK275(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK275, regK275, CamposChave)
    Set dtoTitSPED.tK275 = Util.MapearTitulos(regK275, 3)
End Sub

Public Sub CarregarDadosRegistroK280(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK280, regK280, CamposChave)
    Set dtoTitSPED.tK280 = Util.MapearTitulos(regK280, 3)
End Sub

Public Sub CarregarDadosRegistroK290(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK290, regK290, CamposChave)
    Set dtoTitSPED.tK290 = Util.MapearTitulos(regK290, 3)
End Sub

Public Sub CarregarDadosRegistroK291(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK291, regK291, CamposChave)
    Set dtoTitSPED.tK291 = Util.MapearTitulos(regK291, 3)
End Sub

Public Sub CarregarDadosRegistroK292(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK292, regK292, CamposChave)
    Set dtoTitSPED.tK292 = Util.MapearTitulos(regK292, 3)
End Sub

Public Sub CarregarDadosRegistroK300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK300, regK300, CamposChave)
    Set dtoTitSPED.tK300 = Util.MapearTitulos(regK300, 3)
End Sub

Public Sub CarregarDadosRegistroK301(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK301, regK301, CamposChave)
    Set dtoTitSPED.tK301 = Util.MapearTitulos(regK301, 3)
End Sub

Public Sub CarregarDadosRegistroK302(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK302, regK302, CamposChave)
    Set dtoTitSPED.tK302 = Util.MapearTitulos(regK302, 3)
End Sub

Public Sub CarregarDadosRegistroK990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rK990, regK990, CamposChave)
    Set dtoTitSPED.tK990 = Util.MapearTitulos(regK990, 3)
End Sub


' --- Funções de Carregamento do Bloco M ---

Public Sub CarregarDadosRegistroM001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM001, regM001, CamposChave)
    Set dtoTitSPED.tM001 = Util.MapearTitulos(regM001, 3)
End Sub

Public Sub CarregarDadosRegistroM100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM100, regM100, CamposChave)
    Set dtoTitSPED.tM100 = Util.MapearTitulos(regM100, 3)
End Sub

Public Sub CarregarDadosRegistroM105(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM105, regM105, CamposChave)
    Set dtoTitSPED.tM105 = Util.MapearTitulos(regM105, 3)
End Sub

Public Sub CarregarDadosRegistroM110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM110, regM110, CamposChave)
    Set dtoTitSPED.tM110 = Util.MapearTitulos(regM110, 3)
End Sub

Public Sub CarregarDadosRegistroM115(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM115, regM115, CamposChave)
    Set dtoTitSPED.tM115 = Util.MapearTitulos(regM115, 3)
End Sub

Public Sub CarregarDadosRegistroM200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM200, regM200, CamposChave)
    Set dtoTitSPED.tM200 = Util.MapearTitulos(regM200, 3)
End Sub

Public Sub CarregarDadosRegistroM205(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM205, regM205, CamposChave)
    Set dtoTitSPED.tM205 = Util.MapearTitulos(regM205, 3)
End Sub

Public Sub CarregarDadosRegistroM210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM210, regM210, CamposChave)
    Set dtoTitSPED.tM210 = Util.MapearTitulos(regM210, 3)
End Sub

Public Sub CarregarDadosRegistroM211(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM211, regM211, CamposChave)
    Set dtoTitSPED.tM211 = Util.MapearTitulos(regM211, 3)
End Sub

Public Sub CarregarDadosRegistroM215(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM215, regM215, CamposChave)
    Set dtoTitSPED.tM215 = Util.MapearTitulos(regM215, 3)
End Sub

Public Sub CarregarDadosRegistroM220(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM220, regM220, CamposChave)
    Set dtoTitSPED.tM220 = Util.MapearTitulos(regM220, 3)
End Sub

Public Sub CarregarDadosRegistroM225(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM225, regM225, CamposChave)
    Set dtoTitSPED.tM225 = Util.MapearTitulos(regM225, 3)
End Sub

Public Sub CarregarDadosRegistroM230(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM230, regM230, CamposChave)
    Set dtoTitSPED.tM230 = Util.MapearTitulos(regM230, 3)
End Sub

Public Sub CarregarDadosRegistroM300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM300, regM300, CamposChave)
    Set dtoTitSPED.tM300 = Util.MapearTitulos(regM300, 3)
End Sub

Public Sub CarregarDadosRegistroM350(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM350, regM350, CamposChave)
    Set dtoTitSPED.tM350 = Util.MapearTitulos(regM350, 3)
End Sub

Public Sub CarregarDadosRegistroM400(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM400, regM400, CamposChave)
    Set dtoTitSPED.tM400 = Util.MapearTitulos(regM400, 3)
End Sub

Public Sub CarregarDadosRegistroM410(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM410, regM410, CamposChave)
    Set dtoTitSPED.tM410 = Util.MapearTitulos(regM410, 3)
End Sub

Public Sub CarregarDadosRegistroM500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM500, regM500, CamposChave)
    Set dtoTitSPED.tM500 = Util.MapearTitulos(regM500, 3)
End Sub

Public Sub CarregarDadosRegistroM505(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM505, regM505, CamposChave)
    Set dtoTitSPED.tM505 = Util.MapearTitulos(regM505, 3)
End Sub

Public Sub CarregarDadosRegistroM510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM510, regM510, CamposChave)
    Set dtoTitSPED.tM510 = Util.MapearTitulos(regM510, 3)
End Sub

Public Sub CarregarDadosRegistroM515(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM515, regM515, CamposChave)
    Set dtoTitSPED.tM515 = Util.MapearTitulos(regM515, 3)
End Sub

Public Sub CarregarDadosRegistroM600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM600, regM600, CamposChave)
    Set dtoTitSPED.tM600 = Util.MapearTitulos(regM600, 3)
End Sub

Public Sub CarregarDadosRegistroM605(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM605, regM605, CamposChave)
    Set dtoTitSPED.tM605 = Util.MapearTitulos(regM605, 3)
End Sub

Public Sub CarregarDadosRegistroM610(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM610, regM610, CamposChave)
    Set dtoTitSPED.tM610 = Util.MapearTitulos(regM610, 3)
End Sub

Public Sub CarregarDadosRegistroM611(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM611, regM611, CamposChave)
    Set dtoTitSPED.tM611 = Util.MapearTitulos(regM611, 3)
End Sub

Public Sub CarregarDadosRegistroM615(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM615, regM615, CamposChave)
    Set dtoTitSPED.tM615 = Util.MapearTitulos(regM615, 3)
End Sub

Public Sub CarregarDadosRegistroM620(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM620, regM620, CamposChave)
    Set dtoTitSPED.tM620 = Util.MapearTitulos(regM620, 3)
End Sub

Public Sub CarregarDadosRegistroM625(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM625, regM625, CamposChave)
    Set dtoTitSPED.tM625 = Util.MapearTitulos(regM625, 3)
End Sub

Public Sub CarregarDadosRegistroM630(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM630, regM630, CamposChave)
    Set dtoTitSPED.tM630 = Util.MapearTitulos(regM630, 3)
End Sub

Public Sub CarregarDadosRegistroM700(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM700, regM700, CamposChave)
    Set dtoTitSPED.tM700 = Util.MapearTitulos(regM700, 3)
End Sub

Public Sub CarregarDadosRegistroM800(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM800, regM800, CamposChave)
    Set dtoTitSPED.tM800 = Util.MapearTitulos(regM800, 3)
End Sub

Public Sub CarregarDadosRegistroM810(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM810, regM810, CamposChave)
    Set dtoTitSPED.tM810 = Util.MapearTitulos(regM810, 3)
End Sub

Public Sub CarregarDadosRegistroM990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rM990, regM990, CamposChave)
    Set dtoTitSPED.tM990 = Util.MapearTitulos(regM990, 3)
End Sub


' --- Funções de Carregamento do Bloco P ---

Public Sub CarregarDadosRegistroP001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP001, regP001, CamposChave)
    Set dtoTitSPED.tP001 = Util.MapearTitulos(regP001, 3)
End Sub

Public Sub CarregarDadosRegistroP010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP010, regP010, CamposChave)
    Set dtoTitSPED.tP010 = Util.MapearTitulos(regP010, 3)
End Sub

Public Sub CarregarDadosRegistroP100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP100, regP100, CamposChave)
    Set dtoTitSPED.tP100 = Util.MapearTitulos(regP100, 3)
End Sub

Public Sub CarregarDadosRegistroP110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP110, regP110, CamposChave)
    Set dtoTitSPED.tP110 = Util.MapearTitulos(regP110, 3)
End Sub

Public Sub CarregarDadosRegistroP199(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP199, regP199, CamposChave)
    Set dtoTitSPED.tP199 = Util.MapearTitulos(regP199, 3)
End Sub

Public Sub CarregarDadosRegistroP200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP200, regP200, CamposChave)
    Set dtoTitSPED.tP200 = Util.MapearTitulos(regP200, 3)
End Sub

Public Sub CarregarDadosRegistroP210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP210, regP210, CamposChave)
    Set dtoTitSPED.tP210 = Util.MapearTitulos(regP210, 3)
End Sub

Public Sub CarregarDadosRegistroP990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.rP990, regP990, CamposChave)
    Set dtoTitSPED.tP990 = Util.MapearTitulos(regP990, 3)
End Sub


' --- Funções de Carregamento do Bloco 1 ---

Public Sub CarregarDadosRegistro1001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1001, reg1001, CamposChave)
    Set dtoTitSPED.t1001 = Util.MapearTitulos(reg1001, 3)
End Sub

Public Sub CarregarDadosRegistro1010(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1010, reg1010, CamposChave)
    Set dtoTitSPED.t1010 = Util.MapearTitulos(reg1010, 3)
End Sub

Public Sub CarregarDadosRegistro1010_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1010_Contr, reg1010_Contr, CamposChave)
    Set dtoTitSPED.t1010_Contr = Util.MapearTitulos(reg1010_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1011(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1011, reg1011, CamposChave)
    Set dtoTitSPED.t1011 = Util.MapearTitulos(reg1011, 3)
End Sub

Public Sub CarregarDadosRegistro1020(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1020, reg1020, CamposChave)
    Set dtoTitSPED.t1020 = Util.MapearTitulos(reg1020, 3)
End Sub

Public Sub CarregarDadosRegistro1050(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1050, reg1050, CamposChave)
    Set dtoTitSPED.t1050 = Util.MapearTitulos(reg1050, 3)
End Sub

Public Sub CarregarDadosRegistro1100(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1100, reg1100, CamposChave)
    Set dtoTitSPED.t1100 = Util.MapearTitulos(reg1100, 3)
End Sub

Public Sub CarregarDadosRegistro1100_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1100_Contr, reg1100_Contr, CamposChave)
    Set dtoTitSPED.t1100_Contr = Util.MapearTitulos(reg1100_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1101(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1101, reg1101, CamposChave)
    Set dtoTitSPED.t1101 = Util.MapearTitulos(reg1101, 3)
End Sub

Public Sub CarregarDadosRegistro1102(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1102, reg1102, CamposChave)
    Set dtoTitSPED.t1102 = Util.MapearTitulos(reg1102, 3)
End Sub

Public Sub CarregarDadosRegistro1105(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1105, reg1105, CamposChave)
    Set dtoTitSPED.t1105 = Util.MapearTitulos(reg1105, 3)
End Sub

Public Sub CarregarDadosRegistro1110(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1110, reg1110, CamposChave)
    Set dtoTitSPED.t1110 = Util.MapearTitulos(reg1110, 3)
End Sub

Public Sub CarregarDadosRegistro1200(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1200, reg1200, CamposChave)
    Set dtoTitSPED.t1200 = Util.MapearTitulos(reg1200, 3)
End Sub

Public Sub CarregarDadosRegistro1210(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1210, reg1210, CamposChave)
    Set dtoTitSPED.t1210 = Util.MapearTitulos(reg1210, 3)
End Sub

Public Sub CarregarDadosRegistro1220(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1220, reg1220, CamposChave)
    Set dtoTitSPED.t1220 = Util.MapearTitulos(reg1220, 3)
End Sub

Public Sub CarregarDadosRegistro1250(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1250, reg1250, CamposChave)
    Set dtoTitSPED.t1250 = Util.MapearTitulos(reg1250, 3)
End Sub

Public Sub CarregarDadosRegistro1255(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1255, reg1255, CamposChave)
    Set dtoTitSPED.t1255 = Util.MapearTitulos(reg1255, 3)
End Sub

Public Sub CarregarDadosRegistro1300(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1300, reg1300, CamposChave)
    Set dtoTitSPED.t1300 = Util.MapearTitulos(reg1300, 3)
End Sub

Public Sub CarregarDadosRegistro1300_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1300_Contr, reg1300_Contr, CamposChave)
    Set dtoTitSPED.t1300_Contr = Util.MapearTitulos(reg1300_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1310(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1310, reg1310, CamposChave)
    Set dtoTitSPED.t1310 = Util.MapearTitulos(reg1310, 3)
End Sub

Public Sub CarregarDadosRegistro1320(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1320, reg1320, CamposChave)
    Set dtoTitSPED.t1320 = Util.MapearTitulos(reg1320, 3)
End Sub

Public Sub CarregarDadosRegistro1350(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1350, reg1350, CamposChave)
    Set dtoTitSPED.t1350 = Util.MapearTitulos(reg1350, 3)
End Sub

Public Sub CarregarDadosRegistro1360(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1360, reg1360, CamposChave)
    Set dtoTitSPED.t1360 = Util.MapearTitulos(reg1360, 3)
End Sub

Public Sub CarregarDadosRegistro1370(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1370, reg1370, CamposChave)
    Set dtoTitSPED.t1370 = Util.MapearTitulos(reg1370, 3)
End Sub

Public Sub CarregarDadosRegistro1390(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1390, reg1390, CamposChave)
    Set dtoTitSPED.t1390 = Util.MapearTitulos(reg1390, 3)
End Sub

Public Sub CarregarDadosRegistro1391(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1391, reg1391, CamposChave)
    Set dtoTitSPED.t1391 = Util.MapearTitulos(reg1391, 3)
End Sub

Public Sub CarregarDadosRegistro1400(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1400, reg1400, CamposChave)
    Set dtoTitSPED.t1400 = Util.MapearTitulos(reg1400, 3)
End Sub

Public Sub CarregarDadosRegistro1500(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1500, reg1500, CamposChave)
    Set dtoTitSPED.t1500 = Util.MapearTitulos(reg1500, 3)
End Sub

Public Sub CarregarDadosRegistro1500_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1500_Contr, reg1500_Contr, CamposChave)
    Set dtoTitSPED.t1500_Contr = Util.MapearTitulos(reg1500_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1501(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1501, reg1501, CamposChave)
    Set dtoTitSPED.t1501 = Util.MapearTitulos(reg1501, 3)
End Sub

Public Sub CarregarDadosRegistro1502(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1502, reg1502, CamposChave)
    Set dtoTitSPED.t1502 = Util.MapearTitulos(reg1502, 3)
End Sub

Public Sub CarregarDadosRegistro1510(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1510, reg1510, CamposChave)
    Set dtoTitSPED.t1510 = Util.MapearTitulos(reg1510, 3)
End Sub

Public Sub CarregarDadosRegistro1600(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1600, reg1600, CamposChave)
    Set dtoTitSPED.t1600 = Util.MapearTitulos(reg1600, 3)
End Sub

Public Sub CarregarDadosRegistro1600_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1600_Contr, reg1600_Contr, CamposChave)
    Set dtoTitSPED.t1600_Contr = Util.MapearTitulos(reg1600_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1601(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1601, reg1601, CamposChave)
    Set dtoTitSPED.t1601 = Util.MapearTitulos(reg1601, 3)
End Sub

Public Sub CarregarDadosRegistro1610(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1610, reg1610, CamposChave)
    Set dtoTitSPED.t1610 = Util.MapearTitulos(reg1610, 3)
End Sub

Public Sub CarregarDadosRegistro1620(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1620, reg1620, CamposChave)
    Set dtoTitSPED.t1620 = Util.MapearTitulos(reg1620, 3)
End Sub

Public Sub CarregarDadosRegistro1700(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1700, reg1700, CamposChave)
    Set dtoTitSPED.t1700 = Util.MapearTitulos(reg1700, 3)
End Sub

Public Sub CarregarDadosRegistro1700_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1700_Contr, reg1700_Contr, CamposChave)
    Set dtoTitSPED.t1700_Contr = Util.MapearTitulos(reg1700_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1710(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1710, reg1710, CamposChave)
    Set dtoTitSPED.t1710 = Util.MapearTitulos(reg1710, 3)
End Sub

Public Sub CarregarDadosRegistro1800(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1800, reg1800, CamposChave)
    Set dtoTitSPED.t1800 = Util.MapearTitulos(reg1800, 3)
End Sub

Public Sub CarregarDadosRegistro1800_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1800_Contr, reg1800_Contr, CamposChave)
    Set dtoTitSPED.t1800_Contr = Util.MapearTitulos(reg1800_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1809(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1809, reg1809, CamposChave)
    Set dtoTitSPED.t1809 = Util.MapearTitulos(reg1809, 3)
End Sub

Public Sub CarregarDadosRegistro1900(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1900, reg1900, CamposChave)
    Set dtoTitSPED.t1900 = Util.MapearTitulos(reg1900, 3)
End Sub

Public Sub CarregarDadosRegistro1900_Contr(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1900_Contr, reg1900_Contr, CamposChave)
    Set dtoTitSPED.t1900_Contr = Util.MapearTitulos(reg1900_Contr, 3)
End Sub

Public Sub CarregarDadosRegistro1910(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1910, reg1910, CamposChave)
    Set dtoTitSPED.t1910 = Util.MapearTitulos(reg1910, 3)
End Sub

Public Sub CarregarDadosRegistro1920(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1920, reg1920, CamposChave)
    Set dtoTitSPED.t1920 = Util.MapearTitulos(reg1920, 3)
End Sub

Public Sub CarregarDadosRegistro1921(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1921, reg1921, CamposChave)
    Set dtoTitSPED.t1921 = Util.MapearTitulos(reg1921, 3)
End Sub

Public Sub CarregarDadosRegistro1922(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1922, reg1922, CamposChave)
    Set dtoTitSPED.t1922 = Util.MapearTitulos(reg1922, 3)
End Sub

Public Sub CarregarDadosRegistro1923(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1923, reg1923, CamposChave)
    Set dtoTitSPED.t1923 = Util.MapearTitulos(reg1923, 3)
End Sub

Public Sub CarregarDadosRegistro1925(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1925, reg1925, CamposChave)
    Set dtoTitSPED.t1925 = Util.MapearTitulos(reg1925, 3)
End Sub

Public Sub CarregarDadosRegistro1926(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1926, reg1926, CamposChave)
    Set dtoTitSPED.t1926 = Util.MapearTitulos(reg1926, 3)
End Sub

Public Sub CarregarDadosRegistro1960(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1960, reg1960, CamposChave)
    Set dtoTitSPED.t1960 = Util.MapearTitulos(reg1960, 3)
End Sub

Public Sub CarregarDadosRegistro1970(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1970, reg1970, CamposChave)
    Set dtoTitSPED.t1970 = Util.MapearTitulos(reg1970, 3)
End Sub

Public Sub CarregarDadosRegistro1975(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1975, reg1975, CamposChave)
    Set dtoTitSPED.t1975 = Util.MapearTitulos(reg1975, 3)
End Sub

Public Sub CarregarDadosRegistro1980(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1980, reg1980, CamposChave)
    Set dtoTitSPED.t1980 = Util.MapearTitulos(reg1980, 3)
End Sub

Public Sub CarregarDadosRegistro1990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r1990, reg1990, CamposChave)
    Set dtoTitSPED.t1990 = Util.MapearTitulos(reg1990, 3)
End Sub


' --- Funções de Carregamento do Bloco 9 ---

Public Sub CarregarDadosRegistro9001(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r9001, reg9001, CamposChave)
    Set dtoTitSPED.t9001 = Util.MapearTitulos(reg9001, 3)
End Sub

Public Sub CarregarDadosRegistro9900(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r9900, reg9900, CamposChave)
    Set dtoTitSPED.t9900 = Util.MapearTitulos(reg9900, 3)
End Sub

Public Sub CarregarDadosRegistro9990(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r9990, reg9990, CamposChave)
    Set dtoTitSPED.t9990 = Util.MapearTitulos(reg9990, 3)
End Sub

Public Sub CarregarDadosRegistro9999(ParamArray CamposChave() As Variant)
    Call CarregarDadosRegistro(dtoRegSPED.r9999, reg9999, CamposChave)
    Set dtoTitSPED.t9999 = Util.MapearTitulos(reg9999, 3)
End Sub
' --- Funções Auxiliares ---

Private Sub CarregarDadosRegistro(ByRef dicDados As Dictionary, ByRef Plan As Worksheet, ByVal Campos As Variant)
    
    If UBound(Campos) >= LBound(Campos) Then _
        Set dicDados = ExtrairDadosRegistro(Plan, Campos) Else _
            Set dicDados = ExtrairDadosRegistro(Plan, Array())
            
End Sub

Private Function ExtrairDadosRegistro(ByVal Plan As Worksheet, Optional CamposChave As Variant) As Dictionary

Dim Titulos As Variant, Registro, Campos, CampoChave, CamposTexto, Campo
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Chave As String
Dim b As Long
    
    If UBound(CamposChave) = -1 Then CamposChave = Array("CHV_REG")
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then
        
        Set ExtrairDadosRegistro = New Dictionary
        Exit Function
        
    End If
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando dados do " & Plan.name, Dados.Rows.Count, Comeco)
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            For Each CampoChave In CamposChave
                arrCamposChave.Add Campos(dicTitulos(CampoChave))
            Next CampoChave
            
            Chave = VBA.Replace(VBA.Join(arrCamposChave.toArray()), " ", "")

            dicDados(Chave) = Campos
            arrCamposChave.Clear
            
         End If
         
    Next Linha
    
    Set ExtrairDadosRegistro = dicDados
    
End Function
