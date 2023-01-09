VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptExtratoFornecedor 
   Caption         =   "Extrato Fornecedor"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19288
   SectionData     =   "rptExtratoFornecedor.dsx":0000
End
Attribute VB_Name = "rptExtratoFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Debitog As Double
Dim Creditog As Double
Dim TotalG As Double

Private Sub ActiveReport_Initialize()
    Call ConfiguraIcones(rptExtratoFornecedor)
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload rptExtratoFornecedor
        Set rptExtratoFornecedor = Nothing
    End If
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Select Case Tool.ID
    Case Is = 1000
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoFornecedor)
    Case Is = 23
        Unload rptExtratoFornecedor
        Set rptExtratoFornecedor = Nothing
    Case Is = 4015
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoFornecedor)
    Case Is = 22
        Call EnviarPeloEmail(MenuPrincipal.cdgRelatorio, rptExtratoFornecedor)
    End Select
End Sub


Private Sub PageFooter_Format()
    lblEmissao.Caption = Format$(Now, "dddd, d mmm yyyy hh:nn:ss")
    lblPagina.Caption = "Página " & STR(Me.pageNumber)
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo trata_erro
    Dim Server As String
    Dim username As String
    Dim password As String
    Dim Database As String

    Server = ReadINI("Servidor", "Server", App.Path & "\vendas.ini")
    username = ReadINI("Servidor", "username", App.Path & "\vendas.ini")
    password = "2562"
    Database = "Papelaria"

    ' string de conexao
    strConnect = "Driver={MySQL ODBC 5.1 Driver};" _
               & " Server= " & Server & ";" _
               & " Database=" & Database & ";" _
               & " User=" & username & ";" _
               & " Password=" & password & ";" _
               & " OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384    'Option=3"

    Sql = " SELECT entrada.*, entradaitens.*,"
    Sql = Sql & " fornecedores.id_fornecedor, fornecedores.fornecedor,"
    Sql = Sql & " Estoques.id_estoque , Estoques.descricao, Estoques.unidade,"
    Sql = Sql & " entradaitens.quantidade * entradaitens.precocusto as total"
    Sql = Sql & " From"
    Sql = Sql & " Entrada"
    Sql = Sql & " LEFT JOIN entradaitens ON entrada.id_entrada = entradaitens.id_entrada"
    Sql = Sql & " LEFT JOIN fornecedores ON entrada.id_fornecedor = fornecedores.id_fornecedor"
    Sql = Sql & " LEFT JOIN estoques ON entradaitens.id_estoque = estoques.id_estoque"
    Sql = Sql & " where "
    Sql = Sql & lblConsulta.Caption
    Sql = Sql & " order by entrada.nfdata"

    ' abre um Recrodset da Tabela Terapia
    Me.dados.ConnectionString = strConnect
    Me.dados.Source = Sql

    Exit Sub

trata_erro:
    Exibe_Erros (Err.Description)

End Sub

















