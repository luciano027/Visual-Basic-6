Attribute VB_Name = "Funcao"


Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function Arc Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Declare Function RoundRect Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public sResultadoDaConsulta As String

Public contapagina As Integer

'-------------------- variaveis do grafico ----------------------
Public Titulo1 As String
Public Titulo2 As String
Public NBarras As Integer
Public TituloB(10) As String
Public ValorB(10) As Double
Public LegendaB(10) As String
'----------------------------------------------------------------



Public Enum VBFLugar
    Vbfantes = 1
    VbFDepois = 2
End Enum

Public Function strzero(s As String, Q As Integer, Optional lugar As VBFLugar = Vbfantes) As String
'S = String a Adicionar os Zeros
'Q = quantidade de Zeros
'Lugar = Adicionar os Zeros Antes ou Depois
'wcodigo_indicadores = strzero(Str(wcodigo), 2, VbFAntes)
    Dim t As Integer
    s = Trim(s)
    ' Adiciona Zero na String
    For t = Len(s) To Q - 1
        If lugar = Vbfantes Then
            s = "0" + s
        Else
            s = s + "0"
        End If
    Next t
    ' Retorna a String
    strzero = Mid(s, 1, Q)
End Function

Public Sub Centela(Parent As Form, Child As Form)
    Dim iTop As Integer
    Dim iLeft As Integer
    iTop = 0
    iLeft = 0
    If Parent.WindowState <> 0 Then
        iTop = ((Parent.Height - Child.Height) \ 2)
        iLeft = ((Parent.Width - Child.Width) \ 2)
        Child.Move iLeft, iTop
    End If
End Sub

Public Sub Centerform(F As Form)
    F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
    F.Left = (Screen.Width) / 2 - F.Width / 2
End Sub
Public Sub travar(Janela As Form)
    Dim i As Integer
    For i = 0 To Janela.Controls.Count - 1
        If TypeOf Janela.Controls(i) Is TextBox Then
            Janela.Controls(i).Locked = True
        End If
    Next i
End Sub
Public Sub Destravar(Janela As Form)
    Dim i As Integer
    For i = 0 To Janela.Controls.Count - 1
        If TypeOf Janela.Controls(i) Is TextBox Then
            Janela.Controls(i).Locked = False
        End If
    Next i
End Sub
Public Sub Lipar(Janela As Form)
    Dim i As Integer
    For i = 0 To Janela.Controls.Count - 1
        If TypeOf Janela.Controls(i) Is TextBox Then
            Janela.Controls(i).text = ""
        End If
        If TypeOf Janela.Controls(i) Is Label Then
            If Janela.Controls(i).BackColor = &H80000005 Then Janela.Controls(i).Caption = ""
        End If
    Next i
End Sub
Public Sub Ver(Janela As Form)
    Dim i As Integer
    For i = 0 To Janela.Controls.Count - 1
        If TypeOf Janela.Controls(i) Is TextBox Then
            Janela.Controls(i).Visible = True
        End If
        If TypeOf Janela.Controls(i) Is Label Then
            If Janela.Controls(i).BackColor = &H80000005 Then Janela.Controls(i).Visible = False
        End If

    Next i
End Sub

Public Sub Omitir(Janela As Form)
    Dim i As Integer
    For i = 0 To Janela.Controls.Count - 1
        If TypeOf Janela.Controls(i) Is TextBox Then
            Janela.Controls(i).Visible = False
        End If
        If TypeOf Janela.Controls(i) Is Label Then
            If Janela.Controls(i).BackColor = &H80000005 Then Janela.Controls(i).Visible = True
        End If
    Next i
End Sub

Public Sub Exibe_Erros(strMensagem_Erro As String)
    On Error Resume Next
    MsgBox strMensagem_Erro, vbInformation, "Ocorreu um erro !"
End Sub


Function SoNumero(campo As String) As String
'Desenvolvido por : Carlos Montoya
'Criado em        : 03/01/2002
'Ultima Manuteção : 03/01/2002
'Função/Sintaxe   : SoNumero(Campo) AS String so com numeros
'Descriçao        : Função para eleminição de caracteres que não sejam numeros de uma string
'                   A função varre a cadeia de string, eliminando letras, simbolos, e caracteres
'                   especiais.

    Dim iX As Integer, Aux As String
    Aux = ""
    For iX = 1 To Len(campo)
        If IsNumeric(Mid(campo, iX, 1)) Then
            Aux = Aux + Mid(campo, iX, 1)
        End If
    Next
    SoNumero = Aux
End Function
Function Extenso_Valor(pdbl_Valor As Double) As String
'Rotina Criada para ler um número e transformá-lo em extenso
'Limite máximo de 9 Bilhões (9.999.999.999,99)
'Não aceita números negativos
'Criada em : 17/12/2004 - Carlos Alberto Cotta Seilhe

    Dim strValorExtenso As String    'Variável que irá armazenar o valor por extenso do número informado
    Dim strNumero As String    'Irá armazenar o número para exibir por extenso
    Dim strCentena, strDezena, strUnidade As String
    Dim dblCentavos, dblValorInteiro As Double
    Dim intContador As Integer
    Dim bln_Bilhao, bln_Milhao, bln_Mil, bln_Real, bln_Unidade As Boolean

    'Verificar se foi informado um dado indevido
    If Not IsNumeric(pdbl_Valor) Or IsEmpty(pdbl_Valor) Then
        strValorExtenso = "Função só suporta números"
    ElseIf pdbl_Valor <= 0 Then    'Verificar se há valor negativo ou nada foi informado
        strValorExtenso = ""
        'Verificar se foi informado um valor não suportado pela função
    ElseIf pdbl_Valor > 9999999999.99 Then
        strValorExtenso = "Valor não Suportado pela Função"
    Else
        'Gerar Extenso Centavos
        dblCentavos = pdbl_Valor - Int(pdbl_Valor)
        'Gerar Extenso parte Inteira
        dblValorInteiro = Int(pdbl_Valor)

        If dblValorInteiro > 0 Then
            For intContador = Len(Trim(STR(dblValorInteiro))) To 1 Step -1
                strNumero = Mid(Trim(STR(dblValorInteiro)), (Len(Trim(STR(dblValorInteiro))) - intContador) + 1, 1)
                Select Case intContador
                Case Is = 10    'Bilhão
                    strValorExtenso = fcn_Numero_Unidade(strNumero) + IIf(strNumero > "1", "Bilhões ", " Bilhão ")
                    bln_Bilhao = True
                Case Is = 9, 6, 3   'Centena
                    If strNumero > "0" Then
                        strCentena = Mid(Trim(STR(dblValorInteiro)), (Len(Trim(STR(dblValorInteiro))) - intContador) + 1, 3)
                        If strCentena > "100" And strCentena < "200" Then
                            strValorExtenso = strValorExtenso + " Cento e "
                        Else
                            strValorExtenso = strValorExtenso + " " + fcn_Numero_Centena(strNumero)
                        End If
                        If intContador = 9 Then
                            bln_Milhao = True
                        ElseIf intContador = 6 Then
                            bln_Mil = True
                        End If
                    End If
                Case Is = 8, 5, 2   'Dezena de Milhão
                    If strNumero > "0" Then
                        strDezena = Mid(Trim(STR(dblValorInteiro)), (Len(Trim(STR(dblValorInteiro))) - intContador) + 1, 2)
                        If strDezena > 10 And strDezena < 20 Then
                            strValorExtenso = strValorExtenso + IIf(Trim(Right(strValorExtenso, 5)) = "entos", " e ", " ") + fcn_Numero_Dezena0(Right(strDezena, 1))
                            bln_Unidade = True
                        Else
                            strValorExtenso = strValorExtenso + IIf(Trim(Right(strValorExtenso, 5)) = "entos", " e ", " ") + fcn_Numero_Dezena1(strNumero)
                            bln_Unidade = False
                        End If
                        If intContador = 8 Then
                            bln_Milhao = True
                        ElseIf intContador = 5 Then
                            bln_Mil = True
                        End If
                    End If
                Case Is = 7, 4, 1   'Unidade de Milhão
                    If strNumero > "0" And Not bln_Unidade Then
                        If Trim(Right(strValorExtenso, 5)) = "entos" Or Trim(Right(strValorExtenso, 3)) = "nte" Or Trim(Right(strValorExtenso, 3)) = "nta" Then
                            strValorExtenso = strValorExtenso + " e "
                        Else
                            strValorExtenso = strValorExtenso + " "
                        End If
                        strValorExtenso = strValorExtenso + fcn_Numero_Unidade(strNumero)
                    End If
                    If intContador = 7 Then
                        If bln_Milhao Or strNumero > "0" Then
                            strValorExtenso = strValorExtenso + IIf(strNumero = "1" And Not bln_Unidade, " Milhão ", " Milhões ")
                            bln_Milhao = True
                        End If
                    End If
                    If intContador = 4 Then
                        If bln_Mil Or strNumero > "0" Then
                            strValorExtenso = strValorExtenso + " Mil "
                            bln_Mil = True
                        End If
                    End If
                    If intContador = 1 Then
                        If (bln_Bilhao And Not bln_Milhao And Not bln_Mil And Right(Trim(STR(dblValorInteiro)), 3) = 0) Or _
                           (Not bln_Bilhao And bln_Milhao And Not bln_Mil And Right(Trim(STR(dblValorInteiro)), 3) = 0) Then
                            strValorExtenso = strValorExtenso + " de "
                        End If
                        strValorExtenso = strValorExtenso + IIf(dblValorInteiro > 1, " Reais ", " Real ")
                    End If
                    bln_Unidade = False
                End Select
            Next intContador
        End If

        If dblCentavos > 0# And dblCentavos < 0.1 Then
            strNumero = Right(Trim(STR(Round(dblCentavos, 2))), 1)
            strValorExtenso = strValorExtenso + IIf(dblValorInteiro > 0, " e ", " ") + fcn_Numero_Unidade(strNumero) + IIf(strNumero > "1", " Centavos ", " Centavo ")
        ElseIf dblCentavos > 0.1 And dblCentavos < 0.2 Then
            strNumero = Right(Trim(STR(Round(dblCentavos, 2) - 0.1)), 1)
            strValorExtenso = strValorExtenso + IIf(dblValorInteiro > 0, " e ", " ") + fcn_Numero_Dezena0(strNumero) + " Centavos "
        Else
            If dblCentavos > 0# Then
                strNumero = Mid(Trim(STR(dblCentavos)), 2, 1)
                strValorExtenso = strValorExtenso + IIf(dblValorInteiro > 0, " e ", " ") + fcn_Numero_Dezena1(strNumero)
                If Len(Trim(STR(dblCentavos))) > 2 Then
                    strNumero = Right(Trim(STR(Round(dblCentavos, 2))), 1)
                    strValorExtenso = strValorExtenso + " e " + fcn_Numero_Unidade(strNumero)
                End If
                strValorExtenso = strValorExtenso + " Centavos "
            End If
        End If

    End If
    Extenso_Valor = Trim(strValorExtenso)
End Function

Function fcn_Numero_Unidade(pstrUnidade As String) As String
'Vetor que irá conter o número por extenso
    Dim array_Unidade(9) As String

    array_Unidade(0) = "Um"
    array_Unidade(1) = "Dois"
    array_Unidade(2) = "Três"
    array_Unidade(3) = "Quatro"
    array_Unidade(4) = "Cinco"
    array_Unidade(5) = "Seis"
    array_Unidade(6) = "Sete"
    array_Unidade(7) = "Oito"
    array_Unidade(8) = "Nove"

    fcn_Numero_Unidade = array_Unidade(Val(pstrUnidade) - 1)
End Function

Function fcn_Numero_Dezena0(pstrDezena0 As String) As String
'Vetor que irá conter o número por extenso
    Dim array_Dezena0(9) As String

    array_Dezena0(0) = "Onze"
    array_Dezena0(1) = "Doze"
    array_Dezena0(2) = "Treze"
    array_Dezena0(3) = "Quatorze"
    array_Dezena0(4) = "Quinze"
    array_Dezena0(5) = "Dezesseis"
    array_Dezena0(6) = "Dezessete"
    array_Dezena0(7) = "Dezoito"
    array_Dezena0(8) = "Dezenove"

    fcn_Numero_Dezena0 = array_Dezena0(Val(pstrDezena0) - 1)
End Function

Function fcn_Numero_Dezena1(pstrDezena1 As String) As String
'Vetor que irá conter o número por extenso
    Dim array_Dezena1(9) As String

    array_Dezena1(0) = "Dez"
    array_Dezena1(1) = "Vinte"
    array_Dezena1(2) = "Trinta"
    array_Dezena1(3) = "Quarenta"
    array_Dezena1(4) = "Cinquenta"
    array_Dezena1(5) = "Sessenta"
    array_Dezena1(6) = "Setenta"
    array_Dezena1(7) = "Oitenta"
    array_Dezena1(8) = "Noventa"

    fcn_Numero_Dezena1 = array_Dezena1(Val(pstrDezena1) - 1)
End Function

Function fcn_Numero_Centena(pstrCentena As String) As String
'Vetor que irá conter o número por extenso
    Dim array_Centena(9) As String

    array_Centena(0) = "Cem"
    array_Centena(1) = "Duzentos"
    array_Centena(2) = "Trezentos"
    array_Centena(3) = "Quatrocentos"
    array_Centena(4) = "Quinhentos"
    array_Centena(5) = "Seiscentos"
    array_Centena(6) = "Setecentos"
    array_Centena(7) = "Oitocentos"
    array_Centena(8) = "Novecentos"

    fcn_Numero_Centena = array_Centena(Val(pstrCentena) - 1)
End Function

Public Function ValidaCPF(CPF As String) As Boolean
    Dim soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    ''Valida argumento
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If
    soma = 0
    For i = 1 To 9
        soma = soma + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
    soma = 0
    For i = 1 To 10
        soma = soma + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
    ValidaCPF = True
End Function


Function FormataTelefone(ByVal text As String) As String
    Dim i As Long
    ' ignora vazio
    If Len(text) = 0 Then Exit Function
    'verifica valores invalidos
    For i = Len(text) To 1 Step -1
        If InStr("0123456789", Mid$(text, i, 1)) = 0 Then
            text = Left$(text, i - 1) & Mid$(text, i + 1)
        End If
    Next
    ' ajusta a posicao correta
    If Len(text) <= 8 Then
        FormataTelefone = Format$(text, "@@@@-@@@@")
    ElseIf Len(text) > 8 And Len(text) <= 9 Then
        FormataTelefone = Format$(text, "(@@@) @@@@-@@@@")
    ElseIf Len(text) > 9 Then
        FormataTelefone = Format$(text, "(@@@) @@@@-@@@@")
    End If
End Function

' preenche o combo para consulta

Public Sub fillCombo(cmbName As ComboBox, TblName As String, fldName As String, Optional criteria As String)
    Dim Tabela As ADODB.Recordset
    Dim Sql As String

    Set Tabela = CreateObject("ADODB.Recordset")

    If criteria = "" Then
        Sql = "Select  * from " & TblName & criteria & " order by " & fldName
    Else
        Sql = criteria
    End If

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        Tabela.MoveFirst
        cmbName.Clear
        While Tabela.EOF = False
            cmbName.AddItem Tabela.Fields(fldName)
            cmbName.ItemData(cmbName.NewIndex) = Tabela.Fields(0)
            Tabela.MoveNext
        Wend
    End If
    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

End Sub

' vericar tabela
Public Function IsExistingTable(ByVal Database As String, ByVal TableName As String) As Boolean
    Dim ConnectString As String
    Dim ADOXConnection As Object
    Dim ADODBConnection As Object
    Dim table As Variant
    ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & MicroBD
    Set ADOXConnection = CreateObject("ADOX.Catalog")
    Set ADODBConnection = CreateObject("ADODB.Connection")
    ADODBConnection.Open ConnectString
    ADOXConnection.ActiveConnection = ADODBConnection
    For Each table In ADOXConnection.Tables
        If LCase(table.Name) = LCase(TableName) Then
            IsExistingTable = True
            Exit For
        End If
    Next
    ADODBConnection.Close
End Function

' exportar arquivos
Public Sub ExportReport(ExportType As String, rptExport As Object, FileNm As String)
    On Error Resume Next
    Dim oPDF As ActiveReportsPDFExport.ARExportPDF
    Dim oEXL As ActiveReportsExcelExport.ARExportExcel
    Dim oTXT As ActiveReportsRTFExport.ARExportRTF
    'Dim oTXT As ActiveReportsTextExport.ARExportText

    rptExport.Run


    Select Case ExportType

    Case "PDF"
        Set oPDF = New ActiveReportsPDFExport.ARExportPDF
        oPDF.FileName = App.Path & "\" & FileNm & ".PDF"
        oPDF.Export rptExport.Pages


    Case "Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = App.Path & "\" & FileNm & ".xls"
        oEXL.Export rptExport.Pages

    End Select


End Sub

Function CalcularIdade(DataInicial As Date, DataFinal As Date) As String
    Dim Anos, Meses, Dias
    Dim Ianos, Imeses, Idias As Double
    Dim Diferenca As Double

    If IsNull(DataInicial) Or DataInicial > Now Or DataInicial > DataFinal Then
        CalcularIdade = "Você ainda não nasceu."
        Exit Function
    End If

    Diferenca = DataFinal - DataInicial

    Ianos = Diferenca / 365.25
    Anos = Int(Ianos)
    Imeses = (Ianos - Anos) * 12
    Meses = Int(Imeses)
    Dias = DateDiff("D", DateSerial(DatePart("yyyy", DataInicial) + Anos, DatePart("m", DataInicial) + Meses, Day(DataInicial)), DataFinal)

    If Dias = 30 Then Dias = 0

    If Meses = 12 Then
        Meses = 0
        Anos = Anos + 1
    End If
    If Anos > 1 Then
        Anos = Anos & " anos "
    Else
        Anos = Anos & " ano "
    End If
    If Meses > 1 Then
        Meses = Meses & " meses "
    Else
        Meses = Meses & " mês "
    End If
    If Dias > 1 Then
        Dias = Dias & " dias "
    Else
        Dias = Dias & " dias "
    End If
    CalcularIdade = Anos & Meses & Dias
End Function

Public Sub GravaLog(ByVal texto As String)
'caso contrario o arquivo eh criado
    Open ArqLog & "\Acessos.txt" For Append As #1
    'Escreve no arquivo
    Print #1, Date & "  " & Time & "  " & texto
    'Fecha o arquivo
    Close #1

End Sub

'Saber o dia da semana de uma determinada data
Function DiaDaSemana(Data As String) As String
    If IsDate(Data) Then
        Select Case Format(Data, "w")
        Case 1
            DiaDaSemana = "Domingo"
        Case 2
            DiaDaSemana = "Segunda-feira"
        Case 3
            DiaDaSemana = "Terça-feira"
        Case 4
            DiaDaSemana = "Quarta-feira"
        Case 5
            DiaDaSemana = "Quinta-feira"
        Case 6
            DiaDaSemana = "Sexta-feira"
        Case 7
            DiaDaSemana = "Sábado"
        End Select
    Else
        DiaDaSemana = "Data Inválida!"
    End If
End Function


Public Function FormatValor(ByVal VVALOR As String, ByVal VTIPO As Integer) As String
    FormatValor = VVALOR
    Select Case VTIPO
    Case 1    'MOEDA
        FormatValor = Replace(VVALOR, ".", "")
        FormatValor = Replace(FormatValor, ",", ".")
    Case 2    'CNPJ
        FormatValor = Replace(VVALOR, ".", "")
        FormatValor = Replace(FormatValor, "/", "")
        FormatValor = Replace(FormatValor, "-", "")
    Case 3    'CEP
        FormatValor = Mid(VVALOR, 1, 5) & "-" & Mid(VVALOR, 6, 3)
    Case 4    'FONE
        FormatValor = Replace(VVALOR, "(", "")
        FormatValor = Replace(VVALOR, ")", "")
        FormatValor = Replace(FormatValor, "-", "")
    End Select

End Function

Public Function Ajusta_Form(frm As Form)
    frm.Move 0, 0, frm.ScaleWidth, frm.ScaleHeight
End Function


Public Function FindInList(Cbo As ComboBox, StrFind As String) As Integer
    Dim X As Integer
    Dim idx As Integer
    'Locate an items index in a combobox
    For X = 0 To Cbo.ListCount
        If LCase(StrFind) = LCase(Cbo.List(X)) Then
            idx = X
            Exit For
        End If
    Next X

    FindInList = idx
End Function


Public Function Aguarde_Process(tela As Form, OP As Boolean, Optional Server As Integer)
    On Error Resume Next
    Dim FraAguarda As Frame
    Dim lblDesc As Label
    Dim shpBorda As Shape
    Dim BarraPSQ As String

    If OP = True Then
        tela.Enabled = False: Screen.MousePointer = 11
        'CRIO O FRAME
        Set FraAguarda = tela.Controls.Add("VB.Frame", "Aguarda", tela)
        FraAguarda.Width = 3055
        FraAguarda.Height = 1035
        FraAguarda.Top = (tela.ScaleHeight - FraAguarda.Height) / 2
        FraAguarda.Left = (tela.ScaleWidth - FraAguarda.Width) / 2
        FraAguarda.BackColor = &H8000000F
        FraAguarda.Visible = True
        FraAguarda.BorderStyle = 0
        FraAguarda.ZOrder (vbBringToFront)
        'CRIO A BORDA PARA O FRAME
        Set shpBorda = tela.Controls.Add("VB.Shape", "Borda", FraAguarda)
        shpBorda.Width = 3055
        shpBorda.Height = 1035
        shpBorda.Top = 0
        shpBorda.Shape = 0
        shpBorda.BorderWidth = 3
        shpBorda.Visible = True
        'CRIO LABEL COM MSG
        Set lblDesc = tela.Controls.Add("VB.Label", "lblDescri", FraAguarda)
        lblDesc.Caption = "  Aguarde, processando ..."
        If Server = 1 Then lblDesc.Caption = "     Conexão Servidor..."
        lblDesc.Font = "Arial"
        lblDesc.Top = 250
        lblDesc.Height = 210
        lblDesc.FontSize = 12
        lblDesc.FontBold = True
        lblDesc.AutoSize = True
        lblDesc.ForeColor = &HFF0000
        lblDesc.ZOrder (vbBringToFront)
        lblDesc.BackStyle = 0
        lblDesc.Visible = True

        tela.Refresh
        tela.Enabled = False: Screen.MousePointer = 11
    Else
        tela.Controls("Aguarda").Visible = False
        tela.Controls.Remove tela.Controls("Aguarda")
        tela.Enabled = True: Screen.MousePointer = 0
    End If
End Function


'----------------------   zebra listview

Function LVZebra(LV As ListView, Pic As PictureBox, Cor1 As Long, Cor2 As Long, tela As Form) As Boolean
    Dim lHght As Long
    Dim lWdth As Long

    LVZebra = False

    If LV.View <> lvwReport Then Exit Function
    If LV.ListItems.Count = 0 Then Exit Function

    With LV
        .Picture = Nothing
        .Refresh
        .Visible = True
        .PictureAlignment = lvwTile
        lWdth = .Width
    End With

    With Pic
        .AutoRedraw = False
        .Picture = Nothing
        .BackColor = vbWhite
        .Height = 1
        .AutoRedraw = True
        .BorderStyle = vbBSNone
        .ScaleMode = vbTwips
        .Top = tela.Top - 10000
        .Width = Screen.Width
        .Visible = False
        .Font = LV.Font

        With .Font
            .Bold = LV.Font.Bold
            .Charset = LV.Font.Charset
            .Italic = LV.Font.Italic
            .Name = LV.Font.Name
            .Strikethrough = LV.Font.Strikethrough
            .Underline = LV.Font.Underline
            .Weight = LV.Font.Weight
            .Size = LV.Font.Size
        End With

        lHght = LV.ListItems(1).Height

        .Height = lHght * 2
        .Width = lWdth

        Pic.Line (0, 0)-(lWdth, lHght), Cor1, BF
        Pic.Line (0, lHght)-(lWdth, (lHght * 2)), Cor2, BF
        .AutoSize = True
        .Refresh

    End With

    LV.Refresh
    LV.Picture = Pic.Image
    LVZebra = True

End Function

'This is a REALLY useful function
Public Function min(var1, var2)
    If var1 < var2 Then min = var1 Else min = var2
End Function


' Leitos NF Eletronica
Public Function cnull(xx As Variant, yy As Variant)

End Function


