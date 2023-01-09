VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientesSaidaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Saída do estoque"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Consulta"
      Height          =   1095
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   1815
      Begin VB.OptionButton optPagamentos 
         Caption         =   "Pagamentos"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optCompras 
         Caption         =   "Compras"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   360
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid_entrada 
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView ListaSaida 
      Height          =   6375
      Left            =   5760
      TabIndex        =   2
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7980
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   32527
            MinWidth        =   32527
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   13800
      TabIndex        =   4
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "  Sair"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmClientesSaidaConsulta.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdRelatorios 
      Height          =   615
      Left            =   12600
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "  Imprimir"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmClientesSaidaConsulta.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdConsultar 
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Consultar"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmClientesSaidaConsulta.frx":021C
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComCtl2.MonthView txtDataF 
      Height          =   2370
      Left            =   3000
      TabIndex        =   9
      Top             =   960
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108658689
      CurrentDate     =   41801
   End
   Begin MSComCtl2.MonthView txtDataI 
      Height          =   2370
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108658689
      CurrentDate     =   41801
   End
   Begin VB.TextBox txtVendedor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   4815
   End
   Begin Vendas.VistaButton cmdLimparConsulta 
      Height          =   615
      Left            =   2520
      TabIndex        =   23
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Limpar Consulta"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmClientesSaidaConsulta.frx":0238
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Image cmdConsultaVendedor 
      Height          =   255
      Left            =   5160
      Picture         =   "frmClientesSaidaConsulta.frx":0254
      Stretch         =   -1  'True
      Top             =   4305
      Width           =   360
   End
   Begin VB.Label lblDinheiro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   19
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Total saida Periodo (R$)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Image cmdConsultaFornecedor 
      Height          =   315
      Left            =   5160
      Picture         =   "frmClientesSaidaConsulta.frx":055E
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3900
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblConsulta 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label lblCadastro 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11760
      TabIndex        =   5
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7350
      Left            =   120
      Picture         =   "frmClientesSaidaConsulta.frx":0868
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5520
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   -120
      Picture         =   "frmClientesSaidaConsulta.frx":2568
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   16680
   End
End
Attribute VB_Name = "frmClientesSaidaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Scampo As String
Dim campo As String
Dim ChaveM As String
Dim Sql As String
Dim SQsort As String
Dim sqlwhere As String
Dim Sqlconsulta As String

Private Sub cmdConsultaFornecedor_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaClientes
        .Show 1
    End With

    If IDCliente <> "" Then
        txtid_cliente.text = IDCliente
        IDCliente = ""
    End If

    Sql = "SELECT id_Cliente, Cliente FROM Clientes WHERE id_Cliente = '" & txtid_cliente.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Cliente")) <> vbNull Then txtid_cliente.text = Tabela("id_Cliente") Else txtid_cliente.text = ""
        If VarType(Tabela("Cliente")) <> vbNull Then txtCliente.text = Tabela("Cliente") Else txtCliente.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdConsultaVendedor_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaVendedor
        .Show 1
    End With

    If IDVendedor <> "" Then
        txtid_vendedor.text = IDVendedor
        IDVendedor = ""
    End If

    Sql = "SELECT id_vendedor, vendedor FROM vendedores WHERE id_vendedor = '" & txtid_vendedor.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Vendedor")) <> vbNull Then txtid_vendedor.text = Tabela("id_Vendedor") Else txtid_vendedor.text = ""
        If VarType(Tabela("Vendedor")) <> vbNull Then txtVendedor.text = Tabela("Vendedor") Else txtVendedor.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdLimparConsulta_Click()
    txtid_cliente.text = ""
    txtCliente.text = ""
    txtid_vendedor.text = ""
    txtVendedor.text = ""

    cmdConsultar_Click
End Sub

Private Sub cmdRelatorios_Click()
    If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then

        If optCompras.Value = True Then

            With rptClienteSaidaConsulta
                .lblConsulta.Caption = Sqlconsulta
                .lblCliente.Caption = "Cliente: " & txtCliente.text
                .lblCliente1.Caption = "Vendedor: " & txtVendedor.text
                .lblPeriodo.Caption = "Período:" & Format(txtDataI.Value, "dd/mm/yyyy") & " a " & Format(txtDataF.Value, "dd/mm/yyyy")
                .Show 1
            End With

        End If

        If optPagamentos.Value = True Then
            If txtid_cliente.text <> "" Then

                Sqlconsulta = " caixa.datacaixa Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
                Sqlconsulta = Sqlconsulta & " and "
                Sqlconsulta = Sqlconsulta & " historico LIKE '%Pagamento conta:" & txtCliente.text & "%'"
            End If
            With rptCaixa
                .lblConsulta.Caption = Sqlconsulta
                .lblCliente.Caption = "Cliente: " & txtCliente.text
                .lbltipo.Caption = "2"

                ' .lblCliente1.Caption = "Vendedor: " & txtVendedor.text
                '.lblPeriodo.Caption = "Período:" & Format(txtDataI.Value, "dd/mm/yyyy") & " a " & Format(txtDataF.Value, "dd/mm/yyyy")
                .Show 1
            End With

        End If


    Else
        MsgBox ("Favor selecionar um periodo..."), vbInformation
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 15285
    Me.Height = 8730
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    txtDataI.Value = Now
    txtDataF.Value = Now

    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmClientesSaidaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmClientesSaidaConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub



'--------------------------- define dados da lista grid Consulta
Private Sub ListaPrazo(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Entradas As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mDinheiro As Double

    ' conecta ao banco de dados
    Set Entradas = CreateObject("ADODB.Recordset")

    Sql = "select prazo.*, clientes.id_cliente, clientes.cliente,"
    Sql = Sql & " prazopagto.*"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " left join clientes on prazo.id_cliente = clientes.id_cliente"
    Sql = Sql & " left join prazopagto on prazo.id_prazo = prazopagto.id_prazo"
    Sql = Sql & " Where"
    Sql = Sql & SQconsulta

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Entradas
    If Entradas.State = 1 Then Entradas.Close
    Entradas.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaSaida.ColumnHeaders.Clear
    ListaSaida.ListItems.Clear

    If Entradas.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Entradas encontrado(s): " & Entradas.RecordCount

    ListaSaida.ColumnHeaders.Add , , "Cliente", 5000
    ListaSaida.ColumnHeaders.Add , , "Data pagto", 2000, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Valor Pago ", 2000, lvwColumnRight


    mDinheiro = 0

    If Entradas.BOF = True And Entradas.EOF = True Then Exit Sub
    While Not Entradas.EOF
        If VarType(Entradas("cliente")) <> vbNull Then Set itemx = ListaSaida.ListItems.Add(, , Entradas("cliente"))
        If VarType(Entradas("datapagto")) <> vbNull Then itemx.SubItems(1) = Format(Entradas("datapagto"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Entradas("valorpagto")) <> vbNull Then itemx.SubItems(2) = Format(Entradas("valorpagto"), "###,##0.00") Else itemx.SubItems(2) = ""

        mDinheiro = mDinheiro + Entradas("valorpagto")

        Entradas.MoveNext
    Wend

    lblDinheiro.Caption = Format(mDinheiro, "###,###,##0.00")

    'Zebra o listview
    If LVZebra(ListaSaida, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Entradas.State = 1 Then Entradas.Close
    Set Entradas = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
    ' Exibe_Erros (Sql)
End Sub






Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Entradas As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mDinheiro As Double

    ' conecta ao banco de dados
    Set Entradas = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = " Estoques.descricao"

    Sql = " SELECT saida.*, estoques.id_estoque, estoques.unidade, estoques.descricao,"
    Sql = Sql & " sum(saida.quantidade) as saidaq,"
    Sql = Sql & " sum(saida.quantidade)*saida.preco_venda AS totalItem"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where"
    Sql = Sql & SQconsulta
    Sql = Sql & " group by estoques.id_estoque"
    Sql = Sql & " order by " & SQsort

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Entradas
    If Entradas.State = 1 Then Entradas.Close
    Entradas.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaSaida.ColumnHeaders.Clear
    ListaSaida.ListItems.Clear

    If Entradas.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Entradas encontrado(s): " & Entradas.RecordCount

    ListaSaida.ColumnHeaders.Add , , "Descrição", 3300
    ListaSaida.ColumnHeaders.Add , , "Quant.", 1300, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Preço ", 1400, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Total ", 1400, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Data", 1700, lvwColumnCenter

    mDinheiro = 0

    If Entradas.BOF = True And Entradas.EOF = True Then Exit Sub
    While Not Entradas.EOF
        If VarType(Entradas("descricao")) <> vbNull Then Set itemx = ListaSaida.ListItems.Add(, , Entradas("descricao"))
        If VarType(Entradas("saidaq")) <> vbNull Then itemx.SubItems(1) = Format(Entradas("saidaq"), "###,##0.00") Else itemx.SubItems(1) = ""
        If VarType(Entradas("preco_venda")) <> vbNull Then itemx.SubItems(2) = Format(Entradas("preco_venda"), "###,##0.00") Else itemx.SubItems(2) = ""
        If VarType(Entradas("totalItem")) <> vbNull Then
            itemx.SubItems(3) = Format(Entradas("totalItem"), "###,##0.00")
            mDinheiro = mDinheiro + Entradas("totalItem")
        Else
            itemx.SubItems(3) = ""
        End If
        If VarType(Entradas("datasaida")) <> vbNull Then itemx.SubItems(4) = Format(Entradas("datasaida"), "DD/MM/YYYY") Else itemx.SubItems(4) = ""
        If VarType(Entradas("id_saida")) <> vbNull Then itemx.Tag = Entradas("id_saida")
        Entradas.MoveNext
    Wend

    lblDinheiro.Caption = Format(mDinheiro, "###,###,##0.00")

    'Zebra o listview
    If LVZebra(ListaSaida, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Entradas.State = 1 Then Entradas.Close
    Set Entradas = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
Private Sub ListaSaida_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Select Case ColumnHeader
    Case Is = "Descriçao"
        SQsort = "estoques.descricao"
    Case Is = "Quant."
        SQsort = "saidaq"
    Case Is = "Preço "
        SQsort = "preco_venda"
    Case Is = "Total "
        SQsort = "totalItem"
    Case Is = "Data"
        SQsort = "datasaida"
    End Select

    cmdConsultar_Click

End Sub


Private Sub cmdConsultar_Click()
    If optCompras.Value = True Then

        Sqlconsulta = " 1=1 "

        If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
            Sqlconsulta = Sqlconsulta & " and saida.datasaida Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
        End If

        If txtid_cliente.text <> "" Then Sqlconsulta = Sqlconsulta & " and saida.id_cliente = '" & txtid_cliente.text & "'"
        If txtid_vendedor.text <> "" Then Sqlconsulta = Sqlconsulta & " and saida.id_vendedor = '" & txtid_vendedor.text & "'"

        Lista (Sqlconsulta)

    End If

    If optPagamentos.Value = True Then
        Sqlconsulta = " 1=1 "

        '     If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
        '         Sqlconsulta = Sqlconsulta & " and prazopagto.datapagto Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
        '     End If

        If txtid_cliente.text <> "" Then

            If txtid_cliente.text <> "" Then Sqlconsulta = Sqlconsulta & " and prazo.id_cliente = '" & txtid_cliente.text & "'"

            ListaPrazo (Sqlconsulta)

        End If

    End If

End Sub


Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub

