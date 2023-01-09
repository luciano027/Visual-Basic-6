VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptExtratoOrcamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orçamento"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20370
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "rptExtratoOrcamento.dsx":0000
End
Attribute VB_Name = "rptExtratoOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Debitog As Double
Dim Creditog As Double
Dim TotalG As Double

Private Sub ActiveReport_Initialize()
    Call ConfiguraIcones(rptExtratoOrcamento)
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload rptExtratoOrcamento
        Set rptExtratoOrcamento = Nothing
    End If
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Select Case Tool.ID
    Case Is = 1000
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoOrcamento)
    Case Is = 23
        Unload rptExtratoOrcamento
        Set rptExtratoOrcamento = Nothing
    Case Is = 4015
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoOrcamento)
    Case Is = 22
        Call EnviarPeloEmail(MenuPrincipal.cdgRelatorio, rptExtratoOrcamento)
    End Select
End Sub


Private Sub PageFooter_Format()
    lblEmissao.Caption = Format$(Now, "dddd, d mmm yyyy hh:nn:ss")
    lblPagina.Caption = "Página " & STR(Me.pageNumber)
End Sub

Private Sub Detail_BeforePrint()
    On Error GoTo trata_erro

    Debitog = Debitog + txtTotalPagar.text

    Creditog = lbltotalCredito.Caption


    TotalG = Debitog - Creditog

    lbltotalDebito.Caption = Format(Debitog, "###,##0.00")
    lblTotalPagar.Caption = Format(TotalG, "###,##0.00")

    Exit Sub

trata_erro:
    Exibe_Erros (Err.Description)

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

    Sql = " SELECT saida.*, estoques.id_estoque, estoques.unidade, estoques.descricao,"
    Sql = Sql & " (saida.quantidade * saida.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " where "
    Sql = Sql & lblConsulta.Caption
    Sql = Sql & " order by saida.dataSaida"

    ' abre um Recrodset da Tabela Terapia
    Me.dados.ConnectionString = strConnect
    Me.dados.Source = Sql

    Exit Sub

trata_erro:
    Exibe_Erros (Err.Description)

End Sub















