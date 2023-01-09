VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptExtratoConta 
   Caption         =   "Extrato"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19288
   SectionData     =   "rptExtratoConta.dsx":0000
End
Attribute VB_Name = "rptExtratoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Debitog As Double
Dim Creditog As Double
Dim TotalG As Double
Dim mItems As Integer




Private Sub ActiveReport_Initialize()
    Call ConfiguraIcones(rptExtratoConta)
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload rptExtratoConta
        Set rptExtratoConta = Nothing
    End If
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Select Case Tool.ID
    Case Is = 1000
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoConta)
    Case Is = 23
        Unload rptExtratoConta
        Set rptExtratoConta = Nothing
    Case Is = 4015
        Call ExportaRelatorio(MenuPrincipal.cdgRelatorio, rptExtratoConta)
    Case Is = 22
        Call EnviarPeloEmail(MenuPrincipal.cdgRelatorio, rptExtratoConta)
    End Select
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
    Sql = Sql & " estoques.codigo_est, (saida.quantidade * saida.preco_venda) as total"
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

Private Sub Detail_Format()
    mItems = mItems + 1
    If mItems > 15 Then
        ' Me.Detail.NewPage = ddNPBeforeAfter
        PageBreak.Enabled = True
        mItems = 0
    End If
End Sub
