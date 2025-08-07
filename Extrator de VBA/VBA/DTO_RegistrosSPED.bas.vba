Attribute VB_Name = "DTO_RegistrosSPED"
Option Explicit

Public dtoRegSPED As ListaRegistrosSPED
Public dtoTitSPED As ListaTitulosSPED

Public Type ListaRegistrosSPED
    
    'Dados do Bloco 0
    r0000 As Dictionary
    r0000_Contr As Dictionary
    r0001 As Dictionary
    r0002 As Dictionary
    r0005 As Dictionary
    r0015 As Dictionary
    r0035 As Dictionary
    r0100 As Dictionary
    r0110 As Dictionary
    r0111 As Dictionary
    r0120 As Dictionary
    r0140 As Dictionary
    r0145 As Dictionary
    r0150 As Dictionary
    r0175 As Dictionary
    r0190 As Dictionary
    r0200 As Dictionary
    r0205 As Dictionary
    r0206 As Dictionary
    r0208 As Dictionary
    r0210 As Dictionary
    r0220 As Dictionary
    r0221 As Dictionary
    r0300 As Dictionary
    r0305 As Dictionary
    r0400 As Dictionary
    r0450 As Dictionary
    r0460 As Dictionary
    r0500 As Dictionary
    r0600 As Dictionary
    r0900 As Dictionary
    r0990 As Dictionary

    'Dados do Bloco A
    rA001 As Dictionary
    rA010 As Dictionary
    rA100 As Dictionary
    rA110 As Dictionary
    rA111 As Dictionary
    rA120 As Dictionary
    rA170 As Dictionary
    rA990 As Dictionary

    'Dados do Bloco B
    rB001 As Dictionary
    rB020 As Dictionary
    rB025 As Dictionary
    rB030 As Dictionary
    rB035 As Dictionary
    rB350 As Dictionary
    rB420 As Dictionary
    rB440 As Dictionary
    rB460 As Dictionary
    rB470 As Dictionary
    rB500 As Dictionary
    rB510 As Dictionary
    rB990 As Dictionary
    
    'Dados do Bloco C
    rC001 As Dictionary
    rC010 As Dictionary
    rC100 As Dictionary
    rC101 As Dictionary
    rC105 As Dictionary
    rC110 As Dictionary
    rC111 As Dictionary
    rC112 As Dictionary
    rC113 As Dictionary
    rC114 As Dictionary
    rC115 As Dictionary
    rC116 As Dictionary
    rC120 As Dictionary
    rC130 As Dictionary
    rC140 As Dictionary
    rC141 As Dictionary
    rC160 As Dictionary
    rC165 As Dictionary
    rC170 As Dictionary
    rC171 As Dictionary
    rC172 As Dictionary
    rC173 As Dictionary
    rC174 As Dictionary
    rC175 As Dictionary
    rC175_Contr As Dictionary
    rC176 As Dictionary
    rC177 As Dictionary
    rC178 As Dictionary
    rC179 As Dictionary
    rC180 As Dictionary
    rC180_Contr As Dictionary
    rC181 As Dictionary
    rC181_Contr As Dictionary
    rC185 As Dictionary
    rC185_Contr As Dictionary
    rC186 As Dictionary
    rC188 As Dictionary
    rC190 As Dictionary
    rC190_Contr As Dictionary
    rC191 As Dictionary
    rC191_Contr As Dictionary
    rC195 As Dictionary
    rC195_Contr As Dictionary
    rC197 As Dictionary
    rC198 As Dictionary
    rC199 As Dictionary
    rC300 As Dictionary
    rC310 As Dictionary
    rC320 As Dictionary
    rC321 As Dictionary
    rC330 As Dictionary
    rC350 As Dictionary
    rC370 As Dictionary
    rC380 As Dictionary
    rC380_Contr As Dictionary
    rC381 As Dictionary
    rC385 As Dictionary
    rC390 As Dictionary
    rC395 As Dictionary
    rC396 As Dictionary
    rC400 As Dictionary
    rC405 As Dictionary
    rC405_Contr As Dictionary
    rC410 As Dictionary
    rC420 As Dictionary
    rC425 As Dictionary
    rC430 As Dictionary
    rC460 As Dictionary
    rC465 As Dictionary
    rC470 As Dictionary
    rC480 As Dictionary
    rC481 As Dictionary
    rC485 As Dictionary
    rC489 As Dictionary
    rC490 As Dictionary
    rC490_Contr As Dictionary
    rC491 As Dictionary
    rC495 As Dictionary
    rC495_Contr As Dictionary
    rC499 As Dictionary
    rC500 As Dictionary
    rC500_Contr As Dictionary
    rC501 As Dictionary
    rC505 As Dictionary
    rC509 As Dictionary
    rC510 As Dictionary
    rC590 As Dictionary
    rC591 As Dictionary
    rC595 As Dictionary
    rC597 As Dictionary
    rC600 As Dictionary
    rC601 As Dictionary
    rC601_Contr As Dictionary
    rC605 As Dictionary
    rC609 As Dictionary
    rC610 As Dictionary
    rC690 As Dictionary
    rC700 As Dictionary
    rC790 As Dictionary
    rC791 As Dictionary
    rC800 As Dictionary
    rC810 As Dictionary
    rC815 As Dictionary
    rC820 As Dictionary
    rC830 As Dictionary
    rC850 As Dictionary
    rC855 As Dictionary
    rC857 As Dictionary
    rC860 As Dictionary
    rC870 As Dictionary
    rC870_Contr As Dictionary
    rC880 As Dictionary
    rC880_Contr As Dictionary
    rC890 As Dictionary
    rC890_Contr As Dictionary
    rC895 As Dictionary
    rC897 As Dictionary
    rC990 As Dictionary
    
    'Dados do Bloco D
    rD001 As Dictionary
    rD010 As Dictionary
    rD100 As Dictionary
    rD101 As Dictionary
    rD101_Contr As Dictionary
    rD105 As Dictionary
    rD110 As Dictionary
    rD111 As Dictionary
    rD120 As Dictionary
    rD130 As Dictionary
    rD140 As Dictionary
    rD150 As Dictionary
    rD160 As Dictionary
    rD161 As Dictionary
    rD162 As Dictionary
    rD170 As Dictionary
    rD180 As Dictionary
    rD190 As Dictionary
    rD195 As Dictionary
    rD197 As Dictionary
    rD200 As Dictionary
    rD201 As Dictionary
    rD205 As Dictionary
    rD209 As Dictionary
    rD300 As Dictionary
    rD300_Contr As Dictionary
    rD301 As Dictionary
    rD309 As Dictionary
    rD310 As Dictionary
    rD350 As Dictionary
    rD350_Contr As Dictionary
    rD355 As Dictionary
    rD359 As Dictionary
    rD360 As Dictionary
    rD365 As Dictionary
    rD370 As Dictionary
    rD390 As Dictionary
    rD400 As Dictionary
    rD410 As Dictionary
    rD411 As Dictionary
    rD420 As Dictionary
    rD500 As Dictionary
    rD501 As Dictionary
    rD505 As Dictionary
    rD509 As Dictionary
    rD510 As Dictionary
    rD530 As Dictionary
    rD590 As Dictionary
    rD600 As Dictionary
    rD600_Contr As Dictionary
    rD601 As Dictionary
    rD605 As Dictionary
    rD609 As Dictionary
    rD610 As Dictionary
    rD690 As Dictionary
    rD695 As Dictionary
    rD696 As Dictionary
    rD697 As Dictionary
    rD700 As Dictionary
    rD730 As Dictionary
    rD731 As Dictionary
    rD735 As Dictionary
    rD737 As Dictionary
    rD750 As Dictionary
    rD760 As Dictionary
    rD761 As Dictionary
    rD990 As Dictionary
    
    'Dados do Bloco E
    rE001 As Dictionary
    rE100 As Dictionary
    rE110 As Dictionary
    rE111 As Dictionary
    rE112 As Dictionary
    rE113 As Dictionary
    rE115 As Dictionary
    rE116 As Dictionary
    rE200 As Dictionary
    rE210 As Dictionary
    rE220 As Dictionary
    rE230 As Dictionary
    rE240 As Dictionary
    rE250 As Dictionary
    rE300 As Dictionary
    rE310 As Dictionary
    rE311 As Dictionary
    rE312 As Dictionary
    rE313 As Dictionary
    rE316 As Dictionary
    rE500 As Dictionary
    rE510 As Dictionary
    rE520 As Dictionary
    rE530 As Dictionary
    rE531 As Dictionary
    rE990 As Dictionary
    
    'Dados do Bloco F
    rF001 As Dictionary
    rF010 As Dictionary
    rF100 As Dictionary
    rF111 As Dictionary
    rF120 As Dictionary
    rF129 As Dictionary
    rF130 As Dictionary
    rF139 As Dictionary
    rF150 As Dictionary
    rF200 As Dictionary
    rF205 As Dictionary
    rF210 As Dictionary
    rF211 As Dictionary
    rF500 As Dictionary
    rF509 As Dictionary
    rF510 As Dictionary
    rF519 As Dictionary
    rF525 As Dictionary
    rF550 As Dictionary
    rF559 As Dictionary
    rF560 As Dictionary
    rF569 As Dictionary
    rF600 As Dictionary
    rF700 As Dictionary
    rF800 As Dictionary
    rF990 As Dictionary
    
    'Dados do Bloco G
    rG001 As Dictionary
    rG110 As Dictionary
    rG125 As Dictionary
    rG126 As Dictionary
    rG130 As Dictionary
    rG140 As Dictionary
    rG990 As Dictionary
    
    'Dados do Bloco H
    rH001 As Dictionary
    rH005 As Dictionary
    rH010 As Dictionary
    rH020 As Dictionary
    rH030 As Dictionary
    rH990 As Dictionary
    
    'Dados do Bloco I
    rI001 As Dictionary
    rI010 As Dictionary
    rI100 As Dictionary
    rI199 As Dictionary
    rI200 As Dictionary
    rI299 As Dictionary
    rI300 As Dictionary
    rI399 As Dictionary
    rI990 As Dictionary
    
    'Dados do Bloco K
    rK001 As Dictionary
    rK010 As Dictionary
    rK100 As Dictionary
    rK200 As Dictionary
    rK210 As Dictionary
    rK215 As Dictionary
    rK220 As Dictionary
    rK230 As Dictionary
    rK235 As Dictionary
    rK250 As Dictionary
    rK255 As Dictionary
    rK260 As Dictionary
    rK265 As Dictionary
    rK270 As Dictionary
    rK275 As Dictionary
    rK280 As Dictionary
    rK290 As Dictionary
    rK291 As Dictionary
    rK292 As Dictionary
    rK300 As Dictionary
    rK301 As Dictionary
    rK302 As Dictionary
    rK990 As Dictionary
    
    'Dados do Bloco M
    rM001 As Dictionary
    rM100 As Dictionary
    rM105 As Dictionary
    rM110 As Dictionary
    rM115 As Dictionary
    rM200 As Dictionary
    rM205 As Dictionary
    rM210 As Dictionary
    rM210_INI As Dictionary
    rM211 As Dictionary
    rM215 As Dictionary
    rM220 As Dictionary
    rM225 As Dictionary
    rM230 As Dictionary
    rM300 As Dictionary
    rM350 As Dictionary
    rM400 As Dictionary
    rM410 As Dictionary
    rM500 As Dictionary
    rM505 As Dictionary
    rM510 As Dictionary
    rM515 As Dictionary
    rM600 As Dictionary
    rM605 As Dictionary
    rM610 As Dictionary
    rM610_INI As Dictionary
    rM611 As Dictionary
    rM615 As Dictionary
    rM620 As Dictionary
    rM625 As Dictionary
    rM630 As Dictionary
    rM700 As Dictionary
    rM800 As Dictionary
    rM810 As Dictionary
    rM990 As Dictionary
    
    'Dados do Bloco P
    rP001 As Dictionary
    rP010 As Dictionary
    rP100 As Dictionary
    rP110 As Dictionary
    rP199 As Dictionary
    rP200 As Dictionary
    rP210 As Dictionary
    rP990 As Dictionary
    
    'Dados do Bloco 1
    r1001 As Dictionary
    r1010 As Dictionary
    r1010_Contr As Dictionary
    r1011 As Dictionary
    r1020 As Dictionary
    r1050 As Dictionary
    r1100 As Dictionary
    r1100_Contr As Dictionary
    r1101 As Dictionary
    r1102 As Dictionary
    r1105 As Dictionary
    r1110 As Dictionary
    r1200 As Dictionary
    r1210 As Dictionary
    r1220 As Dictionary
    r1250 As Dictionary
    r1255 As Dictionary
    r1300 As Dictionary
    r1300_Contr As Dictionary
    r1310 As Dictionary
    r1320 As Dictionary
    r1350 As Dictionary
    r1360 As Dictionary
    r1370 As Dictionary
    r1390 As Dictionary
    r1391 As Dictionary
    r1400 As Dictionary
    r1500 As Dictionary
    r1500_Contr As Dictionary
    r1501 As Dictionary
    r1502 As Dictionary
    r1510 As Dictionary
    r1600 As Dictionary
    r1600_Contr As Dictionary
    r1601 As Dictionary
    r1610 As Dictionary
    r1620 As Dictionary
    r1700 As Dictionary
    r1700_Contr As Dictionary
    r1710 As Dictionary
    r1800 As Dictionary
    r1800_Contr As Dictionary
    r1809 As Dictionary
    r1900 As Dictionary
    r1900_Contr As Dictionary
    r1910 As Dictionary
    r1920 As Dictionary
    r1921 As Dictionary
    r1922 As Dictionary
    r1923 As Dictionary
    r1925 As Dictionary
    r1926 As Dictionary
    r1960 As Dictionary
    r1970 As Dictionary
    r1975 As Dictionary
    r1980 As Dictionary
    r1990 As Dictionary
    
    'Dados do Bloco 9
    r9001 As Dictionary
    r9900 As Dictionary
    r9990 As Dictionary
    r9999 As Dictionary
    
End Type

Public Type ListaTitulosSPED
    
    'Títulos do Bloco 0
    t0000 As Dictionary
    t0000_Contr As Dictionary
    t0001 As Dictionary
    t0002 As Dictionary
    t0005 As Dictionary
    t0015 As Dictionary
    t0035 As Dictionary
    t0100 As Dictionary
    t0110 As Dictionary
    t0111 As Dictionary
    t0120 As Dictionary
    t0140 As Dictionary
    t0145 As Dictionary
    t0150 As Dictionary
    t0175 As Dictionary
    t0190 As Dictionary
    t0200 As Dictionary
    t0205 As Dictionary
    t0206 As Dictionary
    t0208 As Dictionary
    t0210 As Dictionary
    t0220 As Dictionary
    t0221 As Dictionary
    t0300 As Dictionary
    t0305 As Dictionary
    t0400 As Dictionary
    t0450 As Dictionary
    t0460 As Dictionary
    t0500 As Dictionary
    t0600 As Dictionary
    t0900 As Dictionary
    t0990 As Dictionary

    'Títulos do Bloco A
    tA001 As Dictionary
    tA010 As Dictionary
    tA100 As Dictionary
    tA110 As Dictionary
    tA111 As Dictionary
    tA120 As Dictionary
    tA170 As Dictionary
    tA990 As Dictionary

    'Títulos do Bloco B
    tB001 As Dictionary
    tB020 As Dictionary
    tB025 As Dictionary
    tB030 As Dictionary
    tB035 As Dictionary
    tB350 As Dictionary
    tB420 As Dictionary
    tB440 As Dictionary
    tB460 As Dictionary
    tB470 As Dictionary
    tB500 As Dictionary
    tB510 As Dictionary
    tB990 As Dictionary
    
    'Títulos do Bloco C
    tC001 As Dictionary
    tC010 As Dictionary
    tC100 As Dictionary
    tC101 As Dictionary
    tC105 As Dictionary
    tC110 As Dictionary
    tC111 As Dictionary
    tC112 As Dictionary
    tC113 As Dictionary
    tC114 As Dictionary
    tC115 As Dictionary
    tC116 As Dictionary
    tC120 As Dictionary
    tC130 As Dictionary
    tC140 As Dictionary
    tC141 As Dictionary
    tC160 As Dictionary
    tC165 As Dictionary
    tC170 As Dictionary
    tC171 As Dictionary
    tC172 As Dictionary
    tC173 As Dictionary
    tC174 As Dictionary
    tC175 As Dictionary
    tC175_Contr As Dictionary
    tC176 As Dictionary
    tC177 As Dictionary
    tC178 As Dictionary
    tC179 As Dictionary
    tC180 As Dictionary
    tC180_Contr As Dictionary
    tC181 As Dictionary
    tC181_Contr As Dictionary
    tC185 As Dictionary
    tC185_Contr As Dictionary
    tC186 As Dictionary
    tC188 As Dictionary
    tC190 As Dictionary
    tC190_Contr As Dictionary
    tC191 As Dictionary
    tC191_Contr As Dictionary
    tC195 As Dictionary
    tC195_Contr As Dictionary
    tC197 As Dictionary
    tC198 As Dictionary
    tC199 As Dictionary
    tC300 As Dictionary
    tC310 As Dictionary
    tC320 As Dictionary
    tC321 As Dictionary
    tC330 As Dictionary
    tC350 As Dictionary
    tC370 As Dictionary
    tC380 As Dictionary
    tC380_Contr As Dictionary
    tC381 As Dictionary
    tC385 As Dictionary
    tC390 As Dictionary
    tC395 As Dictionary
    tC396 As Dictionary
    tC400 As Dictionary
    tC405 As Dictionary
    tC405_Contr As Dictionary
    tC410 As Dictionary
    tC420 As Dictionary
    tC425 As Dictionary
    tC430 As Dictionary
    tC460 As Dictionary
    tC465 As Dictionary
    tC470 As Dictionary
    tC480 As Dictionary
    tC481 As Dictionary
    tC485 As Dictionary
    tC489 As Dictionary
    tC490 As Dictionary
    tC490_Contr As Dictionary
    tC491 As Dictionary
    tC495 As Dictionary
    tC495_Contr As Dictionary
    tC499 As Dictionary
    tC500 As Dictionary
    tC500_Contr As Dictionary
    tC501 As Dictionary
    tC505 As Dictionary
    tC509 As Dictionary
    tC510 As Dictionary
    tC590 As Dictionary
    tC591 As Dictionary
    tC595 As Dictionary
    tC597 As Dictionary
    tC600 As Dictionary
    tC601 As Dictionary
    tC601_Contr As Dictionary
    tC605 As Dictionary
    tC609 As Dictionary
    tC610 As Dictionary
    tC690 As Dictionary
    tC700 As Dictionary
    tC790 As Dictionary
    tC791 As Dictionary
    tC800 As Dictionary
    tC810 As Dictionary
    tC815 As Dictionary
    tC820 As Dictionary
    tC830 As Dictionary
    tC850 As Dictionary
    tC855 As Dictionary
    tC857 As Dictionary
    tC860 As Dictionary
    tC870 As Dictionary
    tC870_Contr As Dictionary
    tC880 As Dictionary
    tC880_Contr As Dictionary
    tC890 As Dictionary
    tC890_Contr As Dictionary
    tC895 As Dictionary
    tC897 As Dictionary
    tC990 As Dictionary
    
    'Títulos do Bloco D
    tD001 As Dictionary
    tD010 As Dictionary
    tD100 As Dictionary
    tD101 As Dictionary
    tD101_Contr As Dictionary
    tD105 As Dictionary
    tD110 As Dictionary
    tD111 As Dictionary
    tD120 As Dictionary
    tD130 As Dictionary
    tD140 As Dictionary
    tD150 As Dictionary
    tD160 As Dictionary
    tD161 As Dictionary
    tD162 As Dictionary
    tD170 As Dictionary
    tD180 As Dictionary
    tD190 As Dictionary
    tD195 As Dictionary
    tD197 As Dictionary
    tD200 As Dictionary
    tD201 As Dictionary
    tD205 As Dictionary
    tD209 As Dictionary
    tD300 As Dictionary
    tD300_Contr As Dictionary
    tD301 As Dictionary
    tD309 As Dictionary
    tD310 As Dictionary
    tD350 As Dictionary
    tD350_Contr As Dictionary
    tD355 As Dictionary
    tD359 As Dictionary
    tD360 As Dictionary
    tD365 As Dictionary
    tD370 As Dictionary
    tD390 As Dictionary
    tD400 As Dictionary
    tD410 As Dictionary
    tD411 As Dictionary
    tD420 As Dictionary
    tD500 As Dictionary
    tD501 As Dictionary
    tD505 As Dictionary
    tD509 As Dictionary
    tD510 As Dictionary
    tD530 As Dictionary
    tD590 As Dictionary
    tD600 As Dictionary
    tD600_Contr As Dictionary
    tD601 As Dictionary
    tD605 As Dictionary
    tD609 As Dictionary
    tD610 As Dictionary
    tD690 As Dictionary
    tD695 As Dictionary
    tD696 As Dictionary
    tD697 As Dictionary
    tD700 As Dictionary
    tD730 As Dictionary
    tD731 As Dictionary
    tD735 As Dictionary
    tD737 As Dictionary
    tD750 As Dictionary
    tD760 As Dictionary
    tD761 As Dictionary
    tD990 As Dictionary
    
    'Títulos do Bloco E
    tE001 As Dictionary
    tE100 As Dictionary
    tE110 As Dictionary
    tE111 As Dictionary
    tE112 As Dictionary
    tE113 As Dictionary
    tE115 As Dictionary
    tE116 As Dictionary
    tE200 As Dictionary
    tE210 As Dictionary
    tE220 As Dictionary
    tE230 As Dictionary
    tE240 As Dictionary
    tE250 As Dictionary
    tE300 As Dictionary
    tE310 As Dictionary
    tE311 As Dictionary
    tE312 As Dictionary
    tE313 As Dictionary
    tE316 As Dictionary
    tE500 As Dictionary
    tE510 As Dictionary
    tE520 As Dictionary
    tE530 As Dictionary
    tE531 As Dictionary
    tE990 As Dictionary
    
    'Títulos do Bloco F
    tF001 As Dictionary
    tF010 As Dictionary
    tF100 As Dictionary
    tF111 As Dictionary
    tF120 As Dictionary
    tF129 As Dictionary
    tF130 As Dictionary
    tF139 As Dictionary
    tF150 As Dictionary
    tF200 As Dictionary
    tF205 As Dictionary
    tF210 As Dictionary
    tF211 As Dictionary
    tF500 As Dictionary
    tF509 As Dictionary
    tF510 As Dictionary
    tF519 As Dictionary
    tF525 As Dictionary
    tF550 As Dictionary
    tF559 As Dictionary
    tF560 As Dictionary
    tF569 As Dictionary
    tF600 As Dictionary
    tF700 As Dictionary
    tF800 As Dictionary
    tF990 As Dictionary
    
    'Títulos do Bloco G
    tG001 As Dictionary
    tG110 As Dictionary
    tG125 As Dictionary
    tG126 As Dictionary
    tG130 As Dictionary
    tG140 As Dictionary
    tG990 As Dictionary
    
    'Títulos do Bloco H
    tH001 As Dictionary
    tH005 As Dictionary
    tH010 As Dictionary
    tH020 As Dictionary
    tH030 As Dictionary
    tH990 As Dictionary
    
    'Títulos do Bloco I
    tI001 As Dictionary
    tI010 As Dictionary
    tI100 As Dictionary
    tI199 As Dictionary
    tI200 As Dictionary
    tI299 As Dictionary
    tI300 As Dictionary
    tI399 As Dictionary
    tI990 As Dictionary
    
    'Títulos do Bloco K
    tK001 As Dictionary
    tK010 As Dictionary
    tK100 As Dictionary
    tK200 As Dictionary
    tK210 As Dictionary
    tK215 As Dictionary
    tK220 As Dictionary
    tK230 As Dictionary
    tK235 As Dictionary
    tK250 As Dictionary
    tK255 As Dictionary
    tK260 As Dictionary
    tK265 As Dictionary
    tK270 As Dictionary
    tK275 As Dictionary
    tK280 As Dictionary
    tK290 As Dictionary
    tK291 As Dictionary
    tK292 As Dictionary
    tK300 As Dictionary
    tK301 As Dictionary
    tK302 As Dictionary
    tK990 As Dictionary
    
    'Títulos do Bloco M
    tM001 As Dictionary
    tM100 As Dictionary
    tM105 As Dictionary
    tM110 As Dictionary
    tM115 As Dictionary
    tM200 As Dictionary
    tM205 As Dictionary
    tM210 As Dictionary
    tM210_INI As Dictionary
    tM211 As Dictionary
    tM215 As Dictionary
    tM220 As Dictionary
    tM225 As Dictionary
    tM230 As Dictionary
    tM300 As Dictionary
    tM350 As Dictionary
    tM400 As Dictionary
    tM410 As Dictionary
    tM500 As Dictionary
    tM505 As Dictionary
    tM510 As Dictionary
    tM515 As Dictionary
    tM600 As Dictionary
    tM605 As Dictionary
    tM610 As Dictionary
    tM610_INI As Dictionary
    tM611 As Dictionary
    tM615 As Dictionary
    tM620 As Dictionary
    tM625 As Dictionary
    tM630 As Dictionary
    tM700 As Dictionary
    tM800 As Dictionary
    tM810 As Dictionary
    tM990 As Dictionary
    
    'Títulos do Bloco P
    tP001 As Dictionary
    tP010 As Dictionary
    tP100 As Dictionary
    tP110 As Dictionary
    tP199 As Dictionary
    tP200 As Dictionary
    tP210 As Dictionary
    tP990 As Dictionary
    
    'Títulos do Bloco 1
    t1001 As Dictionary
    t1010 As Dictionary
    t1010_Contr As Dictionary
    t1011 As Dictionary
    t1020 As Dictionary
    t1050 As Dictionary
    t1100 As Dictionary
    t1100_Contr As Dictionary
    t1101 As Dictionary
    t1102 As Dictionary
    t1105 As Dictionary
    t1110 As Dictionary
    t1200 As Dictionary
    t1210 As Dictionary
    t1220 As Dictionary
    t1250 As Dictionary
    t1255 As Dictionary
    t1300 As Dictionary
    t1300_Contr As Dictionary
    t1310 As Dictionary
    t1320 As Dictionary
    t1350 As Dictionary
    t1360 As Dictionary
    t1370 As Dictionary
    t1390 As Dictionary
    t1391 As Dictionary
    t1400 As Dictionary
    t1500 As Dictionary
    t1500_Contr As Dictionary
    t1501 As Dictionary
    t1502 As Dictionary
    t1510 As Dictionary
    t1600 As Dictionary
    t1600_Contr As Dictionary
    t1601 As Dictionary
    t1610 As Dictionary
    t1620 As Dictionary
    t1700 As Dictionary
    t1700_Contr As Dictionary
    t1710 As Dictionary
    t1800 As Dictionary
    t1800_Contr As Dictionary
    t1809 As Dictionary
    t1900 As Dictionary
    t1900_Contr As Dictionary
    t1910 As Dictionary
    t1920 As Dictionary
    t1921 As Dictionary
    t1922 As Dictionary
    t1923 As Dictionary
    t1925 As Dictionary
    t1926 As Dictionary
    t1960 As Dictionary
    t1970 As Dictionary
    t1975 As Dictionary
    t1980 As Dictionary
    t1990 As Dictionary
    
    'Títulos do Bloco 9
    t9001 As Dictionary
    t9900 As Dictionary
    t9990 As Dictionary
    t9999 As Dictionary
    
End Type

'Títulos do Relatório
Public dicTitulos As Dictionary
Public dicTitulosLivro As Dictionary
Public dicTitulosApuracao As Dictionary

Public Function ResetarRegistrosSPED()

Dim RegistrosVazios As ListaRegistrosSPED
Dim TitulosVazios As ListaTitulosSPED
    
    LSet dtoRegSPED = RegistrosVazios
    LSet dtoTitSPED = TitulosVazios
    
End Function
