Attribute VB_Name = "ExportadorRegistros_ListaBlocos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public colPlanilhas As New Collection
Public colRegistros As New Collection

Public Sub CarregarRegistrosExportacao(ByVal Registros As Variant)

Dim Registro As Variant
    
    For Each Registro In Registros
        
        Call SelecionarBlocoRegistro(Registro)
        
    Next Registro
    
End Sub

Private Sub SelecionarBlocoRegistro(ByVal Registro As String)
    
    Select Case VBA.Left(Util.RemoverAspaSimples(Registro), 1)
        
        Case "0"
            Call SelecionarRegistros_Bloco0(Registro)

        Case "A"
            Call SelecionarRegistros_BlocoA(Registro)

        Case "B"
            Call SelecionarRegistros_BlocoB(Registro)

        Case "C"
            Call SelecionarRegistros_BlocoC(Registro)

        Case "D"
            Call SelecionarRegistros_BlocoD(Registro)

        Case "E"
            Call SelecionarRegistros_BlocoE(Registro)

        Case "F"
            Call SelecionarRegistros_BlocoF(Registro)

        Case "G"
            Call SelecionarRegistros_BlocoG(Registro)

        Case "H"
            Call SelecionarRegistros_BlocoH(Registro)

        Case "I"
            Call SelecionarRegistros_BlocoI(Registro)

        Case "K"
            Call SelecionarRegistros_BlocoK(Registro)

        Case "M"
            Call SelecionarRegistros_BlocoM(Registro)

        Case "P"
            Call SelecionarRegistros_BlocoP(Registro)

        Case "1"
            Call SelecionarRegistros_Bloco1(Registro)

        Case "9"
            Call SelecionarRegistros_Bloco9(Registro)

    End Select

End Sub

Private Sub SelecionarRegistros_Bloco0(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "0000"
                colPlanilhas.Add reg0000
                colRegistros.Add .r0000

            Case "0000_Contr"
                colPlanilhas.Add reg0000_Contr
                colRegistros.Add .r0000_Contr

            Case "0001"
                colPlanilhas.Add reg0001
                colRegistros.Add .r0001

            Case "0002"
                colPlanilhas.Add reg0002
                colRegistros.Add .r0002

            Case "0005"
                colPlanilhas.Add reg0005
                colRegistros.Add .r0005

            Case "0015"
                colPlanilhas.Add reg0015
                colRegistros.Add .r0015

            Case "0035"
                colPlanilhas.Add reg0035
                colRegistros.Add .r0035

            Case "0100"
                colPlanilhas.Add reg0100
                colRegistros.Add .r0100

            Case "0110"
                colPlanilhas.Add reg0110
                colRegistros.Add .r0110

            Case "0111"
                colPlanilhas.Add reg0111
                colRegistros.Add .r0111

            Case "0120"
                colPlanilhas.Add reg0120
                colRegistros.Add .r0120

            Case "0140"
                colPlanilhas.Add reg0140
                colRegistros.Add .r0140

            Case "0145"
                colPlanilhas.Add reg0145
                colRegistros.Add .r0145

            Case "0150"
                colPlanilhas.Add reg0150
                colRegistros.Add .r0150

            Case "0175"
                colPlanilhas.Add reg0175
                colRegistros.Add .r0175

            Case "0190"
                colPlanilhas.Add reg0190
                colRegistros.Add .r0190

            Case "0200"
                colPlanilhas.Add reg0200
                colRegistros.Add .r0200

            Case "0205"
                colPlanilhas.Add reg0205
                colRegistros.Add .r0205

            Case "0206"
                colPlanilhas.Add reg0206
                colRegistros.Add .r0206

            Case "0208"
                colPlanilhas.Add reg0208
                colRegistros.Add .r0208

            Case "0210"
                colPlanilhas.Add reg0210
                colRegistros.Add .r0210

            Case "0220"
                colPlanilhas.Add reg0220
                colRegistros.Add .r0220

            Case "0221"
                colPlanilhas.Add reg0221
                colRegistros.Add .r0221

            Case "0300"
                colPlanilhas.Add reg0300
                colRegistros.Add .r0300

            Case "0305"
                colPlanilhas.Add reg0305
                colRegistros.Add .r0305

            Case "0400"
                colPlanilhas.Add reg0400
                colRegistros.Add .r0400

            Case "0450"
                colPlanilhas.Add reg0450
                colRegistros.Add .r0450

            Case "0460"
                colPlanilhas.Add reg0460
                colRegistros.Add .r0460

            Case "0500"
                colPlanilhas.Add reg0500
                colRegistros.Add .r0500

            Case "0600"
                colPlanilhas.Add reg0600
                colRegistros.Add .r0600

            Case "0900"
                colPlanilhas.Add reg0900
                colRegistros.Add .r0900

            Case "0990"
                colPlanilhas.Add reg0990
                colRegistros.Add .r0990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoA(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "A001"
                colPlanilhas.Add regA001
                colRegistros.Add .rA001

            Case "A010"
                colPlanilhas.Add regA010
                colRegistros.Add .rA010

            Case "A100"
                colPlanilhas.Add regA100
                colRegistros.Add .rA100

            Case "A110"
                colPlanilhas.Add regA110
                colRegistros.Add .rA110

            Case "A111"
                colPlanilhas.Add regA111
                colRegistros.Add .rA111

            Case "A120"
                colPlanilhas.Add regA120
                colRegistros.Add .rA120

            Case "A170"
                colPlanilhas.Add regA170
                colRegistros.Add .rA170

            Case "A990"
                colPlanilhas.Add regA990
                colRegistros.Add .rA990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoB(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "B001"
                colPlanilhas.Add regB001
                colRegistros.Add .rB001

            Case "B020"
                colPlanilhas.Add regB020
                colRegistros.Add .rB020

            Case "B025"
                colPlanilhas.Add regB025
                colRegistros.Add .rB025

            Case "B030"
                colPlanilhas.Add regB030
                colRegistros.Add .rB030

            Case "B035"
                colPlanilhas.Add regB035
                colRegistros.Add .rB035

            Case "B350"
                colPlanilhas.Add regB350
                colRegistros.Add .rB350

            Case "B420"
                colPlanilhas.Add regB420
                colRegistros.Add .rB420

            Case "B440"
                colPlanilhas.Add regB440
                colRegistros.Add .rB440

            Case "B460"
                colPlanilhas.Add regB460
                colRegistros.Add .rB460

            Case "B470"
                colPlanilhas.Add regB470
                colRegistros.Add .rB470

            Case "B500"
                colPlanilhas.Add regB500
                colRegistros.Add .rB500

            Case "B510"
                colPlanilhas.Add regB510
                colRegistros.Add .rB510

            Case "B990"
                colPlanilhas.Add regB990
                colRegistros.Add .rB990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoC(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "C001"
                colPlanilhas.Add regC001
                colRegistros.Add .rC001

            Case "C010"
                colPlanilhas.Add regC010
                colRegistros.Add .rC010

            Case "C100"
                colPlanilhas.Add regC100
                colRegistros.Add .rC100

            Case "C101"
                colPlanilhas.Add regC101
                colRegistros.Add .rC101

            Case "C105"
                colPlanilhas.Add regC105
                colRegistros.Add .rC105

            Case "C110"
                colPlanilhas.Add regC110
                colRegistros.Add .rC110

            Case "C111"
                colPlanilhas.Add regC111
                colRegistros.Add .rC111

            Case "C112"
                colPlanilhas.Add regC112
                colRegistros.Add .rC112

            Case "C113"
                colPlanilhas.Add regC113
                colRegistros.Add .rC113

            Case "C114"
                colPlanilhas.Add regC114
                colRegistros.Add .rC114

            Case "C115"
                colPlanilhas.Add regC115
                colRegistros.Add .rC115

            Case "C116"
                colPlanilhas.Add regC116
                colRegistros.Add .rC116

            Case "C120"
                colPlanilhas.Add regC120
                colRegistros.Add .rC120

            Case "C130"
                colPlanilhas.Add regC130
                colRegistros.Add .rC130

            Case "C140"
                colPlanilhas.Add regC140
                colRegistros.Add .rC140

            Case "C141"
                colPlanilhas.Add regC141
                colRegistros.Add .rC141

            Case "C160"
                colPlanilhas.Add regC160
                colRegistros.Add .rC160

            Case "C165"
                colPlanilhas.Add regC165
                colRegistros.Add .rC165

            Case "C170"
                colPlanilhas.Add regC170
                colRegistros.Add .rC170

            Case "C171"
                colPlanilhas.Add regC171
                colRegistros.Add .rC171

            Case "C172"
                colPlanilhas.Add regC172
                colRegistros.Add .rC172

            Case "C173"
                colPlanilhas.Add regC173
                colRegistros.Add .rC173

            Case "C174"
                colPlanilhas.Add regC174
                colRegistros.Add .rC174

            Case "C175"
                colPlanilhas.Add regC175
                colRegistros.Add .rC175

            Case "C175_Contr"
                colPlanilhas.Add regC175_Contr
                colRegistros.Add .rC175_Contr

            Case "C176"
                colPlanilhas.Add regC176
                colRegistros.Add .rC176

            Case "C177"
                colPlanilhas.Add regC177
                colRegistros.Add .rC177

            Case "C178"
                colPlanilhas.Add regC178
                colRegistros.Add .rC178

            Case "C179"
                colPlanilhas.Add regC179
                colRegistros.Add .rC179

            Case "C180"
                colPlanilhas.Add regC180
                colRegistros.Add .rC180

            Case "C180_Contr"
                colPlanilhas.Add regC180_Contr
                colRegistros.Add .rC180_Contr

            Case "C181"
                colPlanilhas.Add regC181
                colRegistros.Add .rC181

            Case "C181_Contr"
                colPlanilhas.Add regC181_Contr
                colRegistros.Add .rC181_Contr

            Case "C185"
                colPlanilhas.Add regC185
                colRegistros.Add .rC185

            Case "C185_Contr"
                colPlanilhas.Add regC185_Contr
                colRegistros.Add .rC185_Contr

            Case "C186"
                colPlanilhas.Add regC186
                colRegistros.Add .rC186

            Case "C188"
                colPlanilhas.Add regC188
                colRegistros.Add .rC188

            Case "C190"
                colPlanilhas.Add regC190
                colRegistros.Add .rC190

            Case "C190_Contr"
                colPlanilhas.Add regC190_Contr
                colRegistros.Add .rC190_Contr

            Case "C191"
                colPlanilhas.Add regC191
                colRegistros.Add .rC191

            Case "C191_Contr"
                colPlanilhas.Add regC191_Contr
                colRegistros.Add .rC191_Contr

            Case "C195"
                colPlanilhas.Add regC195
                colRegistros.Add .rC195

            Case "C195_Contr"
                colPlanilhas.Add regC195_Contr
                colRegistros.Add .rC195_Contr

            Case "C197"
                colPlanilhas.Add regC197
                colRegistros.Add .rC197

            Case "C198"
                colPlanilhas.Add regC198
                colRegistros.Add .rC198

            Case "C199"
                colPlanilhas.Add regC199
                colRegistros.Add .rC199

            Case "C300"
                colPlanilhas.Add regC300
                colRegistros.Add .rC300

            Case "C310"
                colPlanilhas.Add regC310
                colRegistros.Add .rC310

            Case "C320"
                colPlanilhas.Add regC320
                colRegistros.Add .rC320

            Case "C321"
                colPlanilhas.Add regC321
                colRegistros.Add .rC321

            Case "C330"
                colPlanilhas.Add regC330
                colRegistros.Add .rC330

            Case "C350"
                colPlanilhas.Add regC350
                colRegistros.Add .rC350

            Case "C370"
                colPlanilhas.Add regC370
                colRegistros.Add .rC370

            Case "C380"
                colPlanilhas.Add regC380
                colRegistros.Add .rC380

            Case "C380_Contr"
                colPlanilhas.Add regC380_Contr
                colRegistros.Add .rC380_Contr

            Case "C381"
                colPlanilhas.Add regC381
                colRegistros.Add .rC381

            Case "C385"
                colPlanilhas.Add regC385
                colRegistros.Add .rC385

            Case "C390"
                colPlanilhas.Add regC390
                colRegistros.Add .rC390

            Case "C395"
                colPlanilhas.Add regC395
                colRegistros.Add .rC395

            Case "C396"
                colPlanilhas.Add regC396
                colRegistros.Add .rC396

            Case "C400"
                colPlanilhas.Add regC400
                colRegistros.Add .rC400

            Case "C405"
                colPlanilhas.Add regC405
                colRegistros.Add .rC405

            Case "C405_Contr"
                colPlanilhas.Add regC405_Contr
                colRegistros.Add .rC405_Contr

            Case "C410"
                colPlanilhas.Add regC410
                colRegistros.Add .rC410

            Case "C420"
                colPlanilhas.Add regC420
                colRegistros.Add .rC420

            Case "C425"
                colPlanilhas.Add regC425
                colRegistros.Add .rC425

            Case "C430"
                colPlanilhas.Add regC430
                colRegistros.Add .rC430

            Case "C460"
                colPlanilhas.Add regC460
                colRegistros.Add .rC460

            Case "C465"
                colPlanilhas.Add regC465
                colRegistros.Add .rC465

            Case "C470"
                colPlanilhas.Add regC470
                colRegistros.Add .rC470

            Case "C480"
                colPlanilhas.Add regC480
                colRegistros.Add .rC480

            Case "C481"
                colPlanilhas.Add regC481
                colRegistros.Add .rC481

            Case "C485"
                colPlanilhas.Add regC485
                colRegistros.Add .rC485

            Case "C489"
                colPlanilhas.Add regC489
                colRegistros.Add .rC489

            Case "C490"
                colPlanilhas.Add regC490
                colRegistros.Add .rC490

            Case "C490_Contr"
                colPlanilhas.Add regC490_Contr
                colRegistros.Add .rC490_Contr

            Case "C491"
                colPlanilhas.Add regC491
                colRegistros.Add .rC491

            Case "C495"
                colPlanilhas.Add regC495
                colRegistros.Add .rC495

            Case "C495_Contr"
                colPlanilhas.Add regC495_Contr
                colRegistros.Add .rC495_Contr

            Case "C499"
                colPlanilhas.Add regC499
                colRegistros.Add .rC499

            Case "C500"
                colPlanilhas.Add regC500
                colRegistros.Add .rC500

            Case "C500_Contr"
                colPlanilhas.Add regC500_Contr
                colRegistros.Add .rC500_Contr

            Case "C501"
                colPlanilhas.Add regC501
                colRegistros.Add .rC501

            Case "C505"
                colPlanilhas.Add regC505
                colRegistros.Add .rC505

            Case "C509"
                colPlanilhas.Add regC509
                colRegistros.Add .rC509

            Case "C510"
                colPlanilhas.Add regC510
                colRegistros.Add .rC510

            Case "C590"
                colPlanilhas.Add regC590
                colRegistros.Add .rC590

            Case "C591"
                colPlanilhas.Add regC591
                colRegistros.Add .rC591

            Case "C595"
                colPlanilhas.Add regC595
                colRegistros.Add .rC595

            Case "C597"
                colPlanilhas.Add regC597
                colRegistros.Add .rC597

            Case "C600"
                colPlanilhas.Add regC600
                colRegistros.Add .rC600

            Case "C601"
                colPlanilhas.Add regC601
                colRegistros.Add .rC601

            Case "C601_Contr"
                colPlanilhas.Add regC601_Contr
                colRegistros.Add .rC601_Contr

            Case "C605"
                colPlanilhas.Add regC605
                colRegistros.Add .rC605

            Case "C609"
                colPlanilhas.Add regC609
                colRegistros.Add .rC609

            Case "C610"
                colPlanilhas.Add regC610
                colRegistros.Add .rC610

            Case "C690"
                colPlanilhas.Add regC690
                colRegistros.Add .rC690

            Case "C700"
                colPlanilhas.Add regC700
                colRegistros.Add .rC700

            Case "C790"
                colPlanilhas.Add regC790
                colRegistros.Add .rC790

            Case "C791"
                colPlanilhas.Add regC791
                colRegistros.Add .rC791

            Case "C800"
                colPlanilhas.Add regC800
                colRegistros.Add .rC800

            Case "C810"
                colPlanilhas.Add regC810
                colRegistros.Add .rC810

            Case "C815"
                colPlanilhas.Add regC815
                colRegistros.Add .rC815

            'Case "C820"
                'colPlanilhas.Add regC820
                'colRegistros.Add .rC820

            'Case "C830"
                'colPlanilhas.Add regC830
                'colRegistros.Add .rC830

            Case "C850"
                colPlanilhas.Add regC850
                colRegistros.Add .rC850

            Case "C855"
                colPlanilhas.Add regC855
                colRegistros.Add .rC855

            Case "C857"
                colPlanilhas.Add regC857
                colRegistros.Add .rC857

            Case "C860"
                colPlanilhas.Add regC860
                colRegistros.Add .rC860

            Case "C870"
                colPlanilhas.Add regC870
                colRegistros.Add .rC870

            Case "C870_Contr"
                colPlanilhas.Add regC870_Contr
                colRegistros.Add .rC870_Contr

            Case "C880"
                colPlanilhas.Add regC880
                colRegistros.Add .rC880

            Case "C880_Contr"
                colPlanilhas.Add regC880_Contr
                colRegistros.Add .rC880_Contr

            Case "C890"
                colPlanilhas.Add regC890
                colRegistros.Add .rC890

            Case "C890_Contr"
                colPlanilhas.Add regC890_Contr
                colRegistros.Add .rC890_Contr

            Case "C895"
                colPlanilhas.Add regC895
                colRegistros.Add .rC895

            Case "C897"
                colPlanilhas.Add regC897
                colRegistros.Add .rC897

            Case "C990"
                colPlanilhas.Add regC990
                colRegistros.Add .rC990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoD(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "D001"
                colPlanilhas.Add regD001
                colRegistros.Add .rD001

            Case "D010"
                colPlanilhas.Add regD010
                colRegistros.Add .rD010

            Case "D100"
                colPlanilhas.Add regD100
                colRegistros.Add .rD100

            Case "D101"
                colPlanilhas.Add regD101
                colRegistros.Add .rD101

            Case "D101_Contr"
                colPlanilhas.Add regD101_Contr
                colRegistros.Add .rD101_Contr

            Case "D105"
                colPlanilhas.Add regD105
                colRegistros.Add .rD105

            Case "D110"
                colPlanilhas.Add regD110
                colRegistros.Add .rD110

            Case "D111"
                colPlanilhas.Add regD111
                colRegistros.Add .rD111

            Case "D120"
                colPlanilhas.Add regD120
                colRegistros.Add .rD120

            Case "D130"
                colPlanilhas.Add regD130
                colRegistros.Add .rD130

            Case "D140"
                colPlanilhas.Add regD140
                colRegistros.Add .rD140

            Case "D150"
                colPlanilhas.Add regD150
                colRegistros.Add .rD150

            Case "D160"
                colPlanilhas.Add regD160
                colRegistros.Add .rD160

            Case "D161"
                colPlanilhas.Add regD161
                colRegistros.Add .rD161

            Case "D162"
                colPlanilhas.Add regD162
                colRegistros.Add .rD162

            Case "D170"
                colPlanilhas.Add regD170
                colRegistros.Add .rD170

            Case "D180"
                colPlanilhas.Add regD180
                colRegistros.Add .rD180

            Case "D190"
                colPlanilhas.Add regD190
                colRegistros.Add .rD190

            Case "D195"
                colPlanilhas.Add regD195
                colRegistros.Add .rD195

            Case "D197"
                colPlanilhas.Add regD197
                colRegistros.Add .rD197

            Case "D200"
                colPlanilhas.Add regD200
                colRegistros.Add .rD200

            Case "D201"
                colPlanilhas.Add regD201
                colRegistros.Add .rD201

            Case "D205"
                colPlanilhas.Add regD205
                colRegistros.Add .rD205

            Case "D209"
                colPlanilhas.Add regD209
                colRegistros.Add .rD209

            Case "D300"
                colPlanilhas.Add regD300
                colRegistros.Add .rD300

            Case "D300_Contr"
                colPlanilhas.Add regD300_Contr
                colRegistros.Add .rD300_Contr

            Case "D301"
                colPlanilhas.Add regD301
                colRegistros.Add .rD301

            Case "D309"
                colPlanilhas.Add regD309
                colRegistros.Add .rD309

            Case "D310"
                colPlanilhas.Add regD310
                colRegistros.Add .rD310

            Case "D350"
                colPlanilhas.Add regD350
                colRegistros.Add .rD350

            Case "D350_Contr"
                colPlanilhas.Add regD350_Contr
                colRegistros.Add .rD350_Contr

            Case "D355"
                colPlanilhas.Add regD355
                colRegistros.Add .rD355

            Case "D359"
                colPlanilhas.Add regD359
                colRegistros.Add .rD359

            Case "D360"
                colPlanilhas.Add regD360
                colRegistros.Add .rD360

            Case "D365"
                colPlanilhas.Add regD365
                colRegistros.Add .rD365

            Case "D370"
                colPlanilhas.Add regD370
                colRegistros.Add .rD370

            Case "D390"
                colPlanilhas.Add regD390
                colRegistros.Add .rD390

            Case "D400"
                colPlanilhas.Add regD400
                colRegistros.Add .rD400

            Case "D410"
                colPlanilhas.Add regD410
                colRegistros.Add .rD410

            Case "D411"
                colPlanilhas.Add regD411
                colRegistros.Add .rD411

            Case "D420"
                colPlanilhas.Add regD420
                colRegistros.Add .rD420

            Case "D500"
                colPlanilhas.Add regD500
                colRegistros.Add .rD500

            Case "D501"
                colPlanilhas.Add regD501
                colRegistros.Add .rD501

            Case "D505"
                colPlanilhas.Add regD505
                colRegistros.Add .rD505

            Case "D509"
                colPlanilhas.Add regD509
                colRegistros.Add .rD509

            Case "D510"
                colPlanilhas.Add regD510
                colRegistros.Add .rD510

            Case "D530"
                colPlanilhas.Add regD530
                colRegistros.Add .rD530

            Case "D590"
                colPlanilhas.Add regD590
                colRegistros.Add .rD590

            Case "D600"
                colPlanilhas.Add regD600
                colRegistros.Add .rD600

            Case "D600_Contr"
                colPlanilhas.Add regD600_Contr
                colRegistros.Add .rD600_Contr

            Case "D601"
                colPlanilhas.Add regD601
                colRegistros.Add .rD601

            Case "D605"
                colPlanilhas.Add regD605
                colRegistros.Add .rD605

            Case "D609"
                colPlanilhas.Add regD609
                colRegistros.Add .rD609

            Case "D610"
                colPlanilhas.Add regD610
                colRegistros.Add .rD610

            Case "D690"
                colPlanilhas.Add regD690
                colRegistros.Add .rD690

            Case "D695"
                colPlanilhas.Add regD695
                colRegistros.Add .rD695

            Case "D696"
                colPlanilhas.Add regD696
                colRegistros.Add .rD696

            Case "D697"
                colPlanilhas.Add regD697
                colRegistros.Add .rD697

            Case "D700"
                colPlanilhas.Add regD700
                colRegistros.Add .rD700

            Case "D730"
                colPlanilhas.Add regD730
                colRegistros.Add .rD730

            Case "D731"
                colPlanilhas.Add regD731
                colRegistros.Add .rD731

            Case "D735"
                colPlanilhas.Add regD735
                colRegistros.Add .rD735

            Case "D737"
                colPlanilhas.Add regD737
                colRegistros.Add .rD737

            Case "D750"
                colPlanilhas.Add regD750
                colRegistros.Add .rD750

            Case "D760"
                colPlanilhas.Add regD760
                colRegistros.Add .rD760

            Case "D761"
                colPlanilhas.Add regD761
                colRegistros.Add .rD761

            Case "D990"
                colPlanilhas.Add regD990
                colRegistros.Add .rD990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoE(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "E001"
                colPlanilhas.Add regE001
                colRegistros.Add .rE001

            Case "E100"
                colPlanilhas.Add regE100
                colRegistros.Add .rE100

            Case "E110"
                colPlanilhas.Add regE110
                colRegistros.Add .rE110

            Case "E111"
                colPlanilhas.Add regE111
                colRegistros.Add .rE111

            Case "E112"
                colPlanilhas.Add regE112
                colRegistros.Add .rE112

            Case "E113"
                colPlanilhas.Add regE113
                colRegistros.Add .rE113

            Case "E115"
                colPlanilhas.Add regE115
                colRegistros.Add .rE115

            Case "E116"
                colPlanilhas.Add regE116
                colRegistros.Add .rE116

            Case "E200"
                colPlanilhas.Add regE200
                colRegistros.Add .rE200

            Case "E210"
                colPlanilhas.Add regE210
                colRegistros.Add .rE210

            Case "E220"
                colPlanilhas.Add regE220
                colRegistros.Add .rE220

            Case "E230"
                colPlanilhas.Add regE230
                colRegistros.Add .rE230

            Case "E240"
                colPlanilhas.Add regE240
                colRegistros.Add .rE240

            Case "E250"
                colPlanilhas.Add regE250
                colRegistros.Add .rE250

            Case "E300"
                colPlanilhas.Add regE300
                colRegistros.Add .rE300

            Case "E310"
                colPlanilhas.Add regE310
                colRegistros.Add .rE310

            Case "E311"
                colPlanilhas.Add regE311
                colRegistros.Add .rE311

            Case "E312"
                colPlanilhas.Add regE312
                colRegistros.Add .rE312

            Case "E313"
                colPlanilhas.Add regE313
                colRegistros.Add .rE313

            Case "E316"
                colPlanilhas.Add regE316
                colRegistros.Add .rE316

            Case "E500"
                colPlanilhas.Add regE500
                colRegistros.Add .rE500

            Case "E510"
                colPlanilhas.Add regE510
                colRegistros.Add .rE510

            Case "E520"
                colPlanilhas.Add regE520
                colRegistros.Add .rE520

            Case "E530"
                colPlanilhas.Add regE530
                colRegistros.Add .rE530

            Case "E531"
                colPlanilhas.Add regE531
                colRegistros.Add .rE531

            Case "E990"
                colPlanilhas.Add regE990
                colRegistros.Add .rE990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoF(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "F001"
                colPlanilhas.Add regF001
                colRegistros.Add .rF001

            Case "F010"
                colPlanilhas.Add regF010
                colRegistros.Add .rF010

            Case "F100"
                colPlanilhas.Add regF100
                colRegistros.Add .rF100

            Case "F111"
                colPlanilhas.Add regF111
                colRegistros.Add .rF111

            Case "F120"
                colPlanilhas.Add regF120
                colRegistros.Add .rF120

            Case "F129"
                colPlanilhas.Add regF129
                colRegistros.Add .rF129

            Case "F130"
                colPlanilhas.Add regF130
                colRegistros.Add .rF130

            Case "F139"
                colPlanilhas.Add regF139
                colRegistros.Add .rF139

            Case "F150"
                colPlanilhas.Add regF150
                colRegistros.Add .rF150

            Case "F200"
                colPlanilhas.Add regF200
                colRegistros.Add .rF200

            Case "F205"
                colPlanilhas.Add regF205
                colRegistros.Add .rF205

            Case "F210"
                colPlanilhas.Add regF210
                colRegistros.Add .rF210

            Case "F211"
                colPlanilhas.Add regF211
                colRegistros.Add .rF211

            Case "F500"
                colPlanilhas.Add regF500
                colRegistros.Add .rF500

            Case "F509"
                colPlanilhas.Add regF509
                colRegistros.Add .rF509

            Case "F510"
                colPlanilhas.Add regF510
                colRegistros.Add .rF510

            Case "F519"
                colPlanilhas.Add regF519
                colRegistros.Add .rF519

            Case "F525"
                colPlanilhas.Add regF525
                colRegistros.Add .rF525

            Case "F550"
                colPlanilhas.Add regF550
                colRegistros.Add .rF550

            Case "F559"
                colPlanilhas.Add regF559
                colRegistros.Add .rF559

            Case "F560"
                colPlanilhas.Add regF560
                colRegistros.Add .rF560

            Case "F569"
                colPlanilhas.Add regF569
                colRegistros.Add .rF569

            Case "F600"
                colPlanilhas.Add regF600
                colRegistros.Add .rF600

            Case "F700"
                colPlanilhas.Add regF700
                colRegistros.Add .rF700

            Case "F800"
                colPlanilhas.Add regF800
                colRegistros.Add .rF800

            Case "F990"
                colPlanilhas.Add regF990
                colRegistros.Add .rF990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoG(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "G001"
                colPlanilhas.Add regG001
                colRegistros.Add .rG001

            Case "G110"
                colPlanilhas.Add regG110
                colRegistros.Add .rG110

            Case "G125"
                colPlanilhas.Add regG125
                colRegistros.Add .rG125

            Case "G126"
                colPlanilhas.Add regG126
                colRegistros.Add .rG126

            Case "G130"
                colPlanilhas.Add regG130
                colRegistros.Add .rG130

            Case "G140"
                colPlanilhas.Add regG140
                colRegistros.Add .rG140

            Case "G990"
                colPlanilhas.Add regG990
                colRegistros.Add .rG990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoH(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "H001"
                colPlanilhas.Add regH001
                colRegistros.Add .rH001

            Case "H005"
                colPlanilhas.Add regH005
                colRegistros.Add .rH005

            Case "H010"
                colPlanilhas.Add regH010
                colRegistros.Add .rH010

            Case "H020"
                colPlanilhas.Add regH020
                colRegistros.Add .rH020

            Case "H030"
                colPlanilhas.Add regH030
                colRegistros.Add .rH030

            Case "H990"
                colPlanilhas.Add regH990
                colRegistros.Add .rH990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoI(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "I001"
                colPlanilhas.Add regI001
                colRegistros.Add .rI001

            Case "I010"
                colPlanilhas.Add regI010
                colRegistros.Add .rI010

            Case "I100"
                colPlanilhas.Add regI100
                colRegistros.Add .rI100

            Case "I199"
                colPlanilhas.Add regI199
                colRegistros.Add .rI199

            Case "I200"
                colPlanilhas.Add regI200
                colRegistros.Add .rI200

            Case "I299"
                colPlanilhas.Add regI299
                colRegistros.Add .rI299

            Case "I300"
                colPlanilhas.Add regI300
                colRegistros.Add .rI300

            Case "I399"
                colPlanilhas.Add regI399
                colRegistros.Add .rI399

            Case "I990"
                colPlanilhas.Add regI990
                colRegistros.Add .rI990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoK(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "K001"
                colPlanilhas.Add regK001
                colRegistros.Add .rK001

            Case "K010"
                colPlanilhas.Add regK010
                colRegistros.Add .rK010

            Case "K100"
                colPlanilhas.Add regK100
                colRegistros.Add .rK100

            Case "K200"
                colPlanilhas.Add regK200
                colRegistros.Add .rK200

            Case "K210"
                colPlanilhas.Add regK210
                colRegistros.Add .rK210

            Case "K215"
                colPlanilhas.Add regK215
                colRegistros.Add .rK215

            Case "K220"
                colPlanilhas.Add regK220
                colRegistros.Add .rK220

            Case "K230"
                colPlanilhas.Add regK230
                colRegistros.Add .rK230

            Case "K235"
                colPlanilhas.Add regK235
                colRegistros.Add .rK235

            Case "K250"
                colPlanilhas.Add regK250
                colRegistros.Add .rK250

            Case "K255"
                colPlanilhas.Add regK255
                colRegistros.Add .rK255

            Case "K260"
                colPlanilhas.Add regK260
                colRegistros.Add .rK260

            Case "K265"
                colPlanilhas.Add regK265
                colRegistros.Add .rK265

            Case "K270"
                colPlanilhas.Add regK270
                colRegistros.Add .rK270

            Case "K275"
                colPlanilhas.Add regK275
                colRegistros.Add .rK275

            Case "K280"
                colPlanilhas.Add regK280
                colRegistros.Add .rK280

            Case "K290"
                colPlanilhas.Add regK290
                colRegistros.Add .rK290

            Case "K291"
                colPlanilhas.Add regK291
                colRegistros.Add .rK291

            Case "K292"
                colPlanilhas.Add regK292
                colRegistros.Add .rK292

            Case "K300"
                colPlanilhas.Add regK300
                colRegistros.Add .rK300

            Case "K301"
                colPlanilhas.Add regK301
                colRegistros.Add .rK301

            Case "K302"
                colPlanilhas.Add regK302
                colRegistros.Add .rK302

            Case "K990"
                colPlanilhas.Add regK990
                colRegistros.Add .rK990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoM(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "M001"
                colPlanilhas.Add regM001
                colRegistros.Add .rM001

            Case "M100"
                colPlanilhas.Add regM100
                colRegistros.Add .rM100

            Case "M105"
                colPlanilhas.Add regM105
                colRegistros.Add .rM105

            Case "M110"
                colPlanilhas.Add regM110
                colRegistros.Add .rM110

            Case "M115"
                colPlanilhas.Add regM115
                colRegistros.Add .rM115

            Case "M200"
                colPlanilhas.Add regM200
                colRegistros.Add .rM200

            Case "M205"
                colPlanilhas.Add regM205
                colRegistros.Add .rM205

            Case "M210"
                colPlanilhas.Add regM210
                colRegistros.Add .rM210

            Case "M210_INI"
                colPlanilhas.Add regM210_INI
                colRegistros.Add .rM210_INI

            Case "M211"
                colPlanilhas.Add regM211
                colRegistros.Add .rM211

            Case "M215"
                colPlanilhas.Add regM215
                colRegistros.Add .rM215

            Case "M220"
                colPlanilhas.Add regM220
                colRegistros.Add .rM220

            Case "M225"
                colPlanilhas.Add regM225
                colRegistros.Add .rM225

            Case "M230"
                colPlanilhas.Add regM230
                colRegistros.Add .rM230

            Case "M300"
                colPlanilhas.Add regM300
                colRegistros.Add .rM300

            Case "M350"
                colPlanilhas.Add regM350
                colRegistros.Add .rM350

            Case "M400"
                colPlanilhas.Add regM400
                colRegistros.Add .rM400

            Case "M410"
                colPlanilhas.Add regM410
                colRegistros.Add .rM410

            Case "M500"
                colPlanilhas.Add regM500
                colRegistros.Add .rM500

            Case "M505"
                colPlanilhas.Add regM505
                colRegistros.Add .rM505

            Case "M510"
                colPlanilhas.Add regM510
                colRegistros.Add .rM510

            Case "M515"
                colPlanilhas.Add regM515
                colRegistros.Add .rM515

            Case "M600"
                colPlanilhas.Add regM600
                colRegistros.Add .rM600

            Case "M605"
                colPlanilhas.Add regM605
                colRegistros.Add .rM605

            Case "M610"
                colPlanilhas.Add regM610
                colRegistros.Add .rM610

            Case "M610_INI"
                colPlanilhas.Add regM610_INI
                colRegistros.Add .rM610_INI

            Case "M611"
                colPlanilhas.Add regM611
                colRegistros.Add .rM611

            Case "M615"
                colPlanilhas.Add regM615
                colRegistros.Add .rM615

            Case "M620"
                colPlanilhas.Add regM620
                colRegistros.Add .rM620

            Case "M625"
                colPlanilhas.Add regM625
                colRegistros.Add .rM625

            Case "M630"
                colPlanilhas.Add regM630
                colRegistros.Add .rM630

            Case "M700"
                colPlanilhas.Add regM700
                colRegistros.Add .rM700

            Case "M800"
                colPlanilhas.Add regM800
                colRegistros.Add .rM800

            Case "M810"
                colPlanilhas.Add regM810
                colRegistros.Add .rM810

            Case "M990"
                colPlanilhas.Add regM990
                colRegistros.Add .rM990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_BlocoP(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "P001"
                colPlanilhas.Add regP001
                colRegistros.Add .rP001

            Case "P010"
                colPlanilhas.Add regP010
                colRegistros.Add .rP010

            Case "P100"
                colPlanilhas.Add regP100
                colRegistros.Add .rP100

            Case "P110"
                colPlanilhas.Add regP110
                colRegistros.Add .rP110

            Case "P199"
                colPlanilhas.Add regP199
                colRegistros.Add .rP199

            Case "P200"
                colPlanilhas.Add regP200
                colRegistros.Add .rP200

            Case "P210"
                colPlanilhas.Add regP210
                colRegistros.Add .rP210

            Case "P990"
                colPlanilhas.Add regP990
                colRegistros.Add .rP990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_Bloco1(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "1001"
                colPlanilhas.Add reg1001
                colRegistros.Add .r1001

            Case "1010"
                colPlanilhas.Add reg1010
                colRegistros.Add .r1010

            Case "1010_Contr"
                colPlanilhas.Add reg1010_Contr
                colRegistros.Add .r1010_Contr

            Case "1011"
                colPlanilhas.Add reg1011
                colRegistros.Add .r1011

            Case "1020"
                colPlanilhas.Add reg1020
                colRegistros.Add .r1020

            Case "1050"
                colPlanilhas.Add reg1050
                colRegistros.Add .r1050

            Case "1100"
                colPlanilhas.Add reg1100
                colRegistros.Add .r1100

            Case "1100_Contr"
                colPlanilhas.Add reg1100_Contr
                colRegistros.Add .r1100_Contr

            Case "1101"
                colPlanilhas.Add reg1101
                colRegistros.Add .r1101

            Case "1102"
                colPlanilhas.Add reg1102
                colRegistros.Add .r1102

            Case "1105"
                colPlanilhas.Add reg1105
                colRegistros.Add .r1105

            Case "1110"
                colPlanilhas.Add reg1110
                colRegistros.Add .r1110

            Case "1200"
                colPlanilhas.Add reg1200
                colRegistros.Add .r1200

            Case "1210"
                colPlanilhas.Add reg1210
                colRegistros.Add .r1210

            Case "1220"
                colPlanilhas.Add reg1220
                colRegistros.Add .r1220

            Case "1250"
                colPlanilhas.Add reg1250
                colRegistros.Add .r1250

            Case "1255"
                colPlanilhas.Add reg1255
                colRegistros.Add .r1255

            Case "1300"
                colPlanilhas.Add reg1300
                colRegistros.Add .r1300

            Case "1300_Contr"
                colPlanilhas.Add reg1300_Contr
                colRegistros.Add .r1300_Contr

            Case "1310"
                colPlanilhas.Add reg1310
                colRegistros.Add .r1310

            Case "1320"
                colPlanilhas.Add reg1320
                colRegistros.Add .r1320

            Case "1350"
                colPlanilhas.Add reg1350
                colRegistros.Add .r1350

            Case "1360"
                colPlanilhas.Add reg1360
                colRegistros.Add .r1360

            Case "1370"
                colPlanilhas.Add reg1370
                colRegistros.Add .r1370

            Case "1390"
                colPlanilhas.Add reg1390
                colRegistros.Add .r1390

            Case "1391"
                colPlanilhas.Add reg1391
                colRegistros.Add .r1391

            Case "1400"
                colPlanilhas.Add reg1400
                colRegistros.Add .r1400

            Case "1500"
                colPlanilhas.Add reg1500
                colRegistros.Add .r1500

            Case "1500_Contr"
                colPlanilhas.Add reg1500_Contr
                colRegistros.Add .r1500_Contr

            Case "1501"
                colPlanilhas.Add reg1501
                colRegistros.Add .r1501

            Case "1502"
                colPlanilhas.Add reg1502
                colRegistros.Add .r1502

            Case "1510"
                colPlanilhas.Add reg1510
                colRegistros.Add .r1510

            Case "1600"
                colPlanilhas.Add reg1600
                colRegistros.Add .r1600

            Case "1600_Contr"
                colPlanilhas.Add reg1600_Contr
                colRegistros.Add .r1600_Contr

            Case "1601"
                colPlanilhas.Add reg1601
                colRegistros.Add .r1601

            Case "1610"
                colPlanilhas.Add reg1610
                colRegistros.Add .r1610

            Case "1620"
                colPlanilhas.Add reg1620
                colRegistros.Add .r1620

            Case "1700"
                colPlanilhas.Add reg1700
                colRegistros.Add .r1700

            Case "1700_Contr"
                colPlanilhas.Add reg1700_Contr
                colRegistros.Add .r1700_Contr

            Case "1710"
                colPlanilhas.Add reg1710
                colRegistros.Add .r1710

            Case "1800"
                colPlanilhas.Add reg1800
                colRegistros.Add .r1800

            Case "1800_Contr"
                colPlanilhas.Add reg1800_Contr
                colRegistros.Add .r1800_Contr

            Case "1809"
                colPlanilhas.Add reg1809
                colRegistros.Add .r1809

            Case "1900"
                colPlanilhas.Add reg1900
                colRegistros.Add .r1900

            Case "1900_Contr"
                colPlanilhas.Add reg1900_Contr
                colRegistros.Add .r1900_Contr

            Case "1910"
                colPlanilhas.Add reg1910
                colRegistros.Add .r1910

            Case "1920"
                colPlanilhas.Add reg1920
                colRegistros.Add .r1920

            Case "1921"
                colPlanilhas.Add reg1921
                colRegistros.Add .r1921

            Case "1922"
                colPlanilhas.Add reg1922
                colRegistros.Add .r1922

            Case "1923"
                colPlanilhas.Add reg1923
                colRegistros.Add .r1923

            Case "1925"
                colPlanilhas.Add reg1925
                colRegistros.Add .r1925

            Case "1926"
                colPlanilhas.Add reg1926
                colRegistros.Add .r1926

            Case "1960"
                colPlanilhas.Add reg1960
                colRegistros.Add .r1960

            Case "1970"
                colPlanilhas.Add reg1970
                colRegistros.Add .r1970

            Case "1975"
                colPlanilhas.Add reg1975
                colRegistros.Add .r1975

            Case "1980"
                colPlanilhas.Add reg1980
                colRegistros.Add .r1980

            Case "1990"
                colPlanilhas.Add reg1990
                colRegistros.Add .r1990

        End Select

    End With

End Sub

Private Sub SelecionarRegistros_Bloco9(ByVal Registro As String)

    With dtoRegSPED

        Select Case Registro

            Case "9001"
                colPlanilhas.Add reg9001
                colRegistros.Add .r9001

            Case "9900"
                colPlanilhas.Add reg9900
                colRegistros.Add .r9900

            Case "9990"
                colPlanilhas.Add reg9990
                colRegistros.Add .r9990

            Case "9999"
                colPlanilhas.Add reg9999
                colRegistros.Add .r9999

        End Select

    End With

End Sub
