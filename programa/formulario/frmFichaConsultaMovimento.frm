VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaConsultaMovimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Ficha Cliente"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtvalorItem 
      Height          =   285
      Left            =   12240
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   11880
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtConsulta 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   11160
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid_prazoitem 
      Height          =   285
      Left            =   11520
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView ListaMov 
      Height          =   5895
      Left            =   5640
      TabIndex        =   2
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10398
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
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   13680
      TabIndex        =   3
      Top             =   7080
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
      Picture         =   "frmFichaConsultaMovimento.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   7785
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
   Begin MSComctlLib.ListView ListProduto 
      Height          =   4095
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7223
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Vendas.VistaButton cmdAlterar 
      Height          =   615
      Left            =   12480
      TabIndex        =   16
      Top             =   7080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Alterar"
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
      Picture         =   "frmFichaConsultaMovimento.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Total Geral (R$)"
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
      Left            =   5640
      TabIndex        =   22
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label lblTotalGeral 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Produto"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Label lblProduto 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   6600
      Width           =   5775
   End
   Begin VB.Label lblQuantidade 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   14
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   13
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Quantidade"
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
      Left            =   11520
      TabIndex        =   12
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Total (R$)"
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
      Left            =   13200
      TabIndex        =   11
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Produto do Estoque"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   4815
   End
   Begin VB.Image cmdConsultar 
      Height          =   315
      Left            =   5040
      Picture         =   "frmFichaConsultaMovimento.frx":045C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   360
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
      TabIndex        =   6
      Top             =   120
      Width           =   5415
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
      Left            =   5640
      TabIndex        =   5
      Top             =   120
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
      Left            =   11640
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7335
      Left            =   120
      Picture         =   "frmFichaConsultaMovimento.frx":0766
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5400
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   -360
      Picture         =   "frmFichaConsultaMovimento.frx":2D21
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   15480
   End
End
Attribute VB_Name = "frmFichaConsultaMovimento"
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



Private Sub cmdAlterar_Click()
    If txtid_prazoitem.text <> "" Then
        With frmFichaAlterarItem
            .lblProduto.Caption = lblProduto.Caption
            .txtValorCompra.text = Format(txtvalorItem.text, "###,##0.00")
            .lblCliente.Caption = txtCliente.text
            .txtid_prazoitem.text = txtid_prazoitem.text
            .Show 1
        End With
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 15030
    Me.Height = 8565
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    ListaProduto ("")

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFichaConsultaMovimento = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaConsultaMovimento = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Prazo As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mQuantidade As Double
    Dim mTotal As Double

    ' conecta ao banco de dados
    Set Prazo = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = " prazoitem.dataCompra"

    Sql = " SELECT clientes.cliente, prazoitem.dataCompra, prazoitem.quantidade, prazoitem.id_estoque,"
    Sql = Sql & " prazoitem.preco_venda, (prazoitem.quantidade*prazoitem.preco_venda) AS totalVenda,"
    Sql = Sql & " prazoitem.id_estoque, prazoitem.id_prazoitem,"
    Sql = Sql & " prazo.*"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN clientes ON prazo.id_cliente = clientes.id_cliente"
    Sql = Sql & " LEFT JOIN prazoitem ON prazo.id_prazo = prazoitem.id_prazo"
    Sql = Sql & " Where "
    Sql = Sql & " prazoitem.id_estoque = '" & txtid_estoque.text & "'"
    Sql = Sql & " order by  prazoitem.dataCompra"

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Prazo
    If Prazo.State = 1 Then Prazo.Close
    Prazo.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaMov.ColumnHeaders.Clear
    ListaMov.ListItems.Clear

    If Prazo.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Prazo encontrado(s): " & Prazo.RecordCount

    ListaMov.ColumnHeaders.Add , , "Cliente", 5800
    ListaMov.ColumnHeaders.Add , , "Data", 1700, lvwColumnCenter
    ListaMov.ColumnHeaders.Add , , "Quantidade", 1500, lvwColumnRight
    ListaMov.ColumnHeaders.Add , , "Preço ", 1500, lvwColumnRight
    ListaMov.ColumnHeaders.Add , , "Total", 1500, lvwColumnRight

    mQuantidade = 0
    mTotal = 0

    If Prazo.BOF = True And Prazo.EOF = True Then Exit Sub
    While Not Prazo.EOF
        If VarType(Prazo("cliente")) <> vbNull Then Set itemx = ListaMov.ListItems.Add(, , Prazo("cliente"))
        If VarType(Prazo("dataCompra")) <> vbNull Then itemx.SubItems(1) = Format(Prazo("dataCompra"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Prazo("quantidade")) <> vbNull Then itemx.SubItems(2) = Format(Prazo("quantidade"), "###,##0.00")
        If VarType(Prazo("preco_venda")) <> vbNull Then itemx.SubItems(3) = Format(Prazo("preco_venda"), "###,##0.00")
        If VarType(Prazo("totalVenda")) <> vbNull Then itemx.SubItems(4) = Format(Prazo("totalVenda"), "###,##0.00")
        If VarType(Prazo("id_prazoitem")) <> vbNull Then itemx.Tag = Prazo("id_prazoitem")

        If VarType(Prazo("quantidade")) <> vbNull Then mQuantidade = mQuantidade + Prazo("quantidade")
        If VarType(Prazo("totalvenda")) <> vbNull Then mTotal = mTotal + Prazo("totalvenda")


        Prazo.MoveNext
    Wend

    lblTotal.Caption = Format(mTotal, "###,##0.00")
    lblQuantidade.Caption = Format(mQuantidade, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaMov, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Prazo.State = 1 Then Prazo.Close
    Set Prazo = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaMov_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    Set Estoque = CreateObject("ADODB.Recordset")

    txtid_prazoitem.text = ListaMov.SelectedItem.Tag

    If txtid_prazoitem.text <> "" Then

        Sql = "SELECT clientes.cliente, prazo.*, prazoitem.preco_venda"
        Sql = Sql & " From"
        Sql = Sql & " Prazo"
        Sql = Sql & " LEFT JOIN clientes ON prazo.id_cliente = clientes.id_cliente"
        Sql = Sql & " LEFT JOIN prazoitem ON prazo.id_prazo = prazoitem.id_prazo"
        Sql = Sql & " where"
        Sql = Sql & " prazoitem.id_prazoitem = '" & txtid_prazoitem.text & "'"

        If Estoque.State = 1 Then Estoque.Close
        Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Estoque.RecordCount > 0 Then
            If VarType(Estoque("preco_venda")) <> vbNull Then txtvalorItem.text = Estoque("preco_venda") Else txtvalorItem.text = ""
            If VarType(Estoque("cliente")) <> vbNull Then txtCliente.text = Estoque("cliente") Else txtCliente.text = ""

        End If
    End If

    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub cmdConsultar_Click()

    If txtConsulta.text <> "" Then
        Sqlconsulta = " estoques.descricao like '%" & txtConsulta.text & "%'"
    Else
        Sqlconsulta = " 1=1"
    End If

    ListaProduto (Sqlconsulta)

End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub ListaProduto(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Prazo As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String

    ' conecta ao banco de dados
    Set Prazo = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = " estoques.descricao"

    Sql = " SELECT prazoitem.*, estoques.id_estoque,estoques.descricao,"
    Sql = Sql & " SUM(prazoitem.quantidade) As quant"
    Sql = Sql & " From"
    Sql = Sql & " prazoitem"
    Sql = Sql & " LEFT JOIN estoques ON prazoitem.id_estoque = estoques.id_estoque"
    If SQconsulta <> "" Then Sql = Sql & " where " & SQconsulta
    Sql = Sql & " GROUP BY prazoitem.id_estoque"
    Sql = Sql & " order by " & SQsort

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Prazo
    If Prazo.State = 1 Then Prazo.Close
    Prazo.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListProduto.ColumnHeaders.Clear
    ListProduto.ListItems.Clear

    If Prazo.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If

    ListProduto.ColumnHeaders.Add , , "Descrição", 3400
    ListProduto.ColumnHeaders.Add , , "Quantidade", 1100, lvwColumnRight

    If Prazo.BOF = True And Prazo.EOF = True Then Exit Sub
    While Not Prazo.EOF
        If VarType(Prazo("descricao")) <> vbNull Then Set itemx = ListProduto.ListItems.Add(, , Prazo("descricao"))
        If VarType(Prazo("quant")) <> vbNull Then itemx.SubItems(1) = Format(Prazo("quant"), "###,##0.00")
        If VarType(Prazo("id_estoque")) <> vbNull Then itemx.Tag = Prazo("id_estoque")
        Prazo.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListProduto, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Prazo.State = 1 Then Prazo.Close
    Set Prazo = Nothing


    Call Total_Geral


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub Total_Geral()
    On Error GoTo trata_erro
    Dim aPrazo As ADODB.Recordset
    Dim Geral As Double

    ' conecta ao banco de dados
    Set aPrazo = CreateObject("ADODB.Recordset")

    Sql = " SELECT prazoitem.*, estoques.id_estoque,estoques.descricao,"
    Sql = Sql & " (prazoitem.quantidade * prazoitem.preco_venda) As quant"
    Sql = Sql & " From"
    Sql = Sql & " prazoitem"
    Sql = Sql & " LEFT JOIN estoques ON prazoitem.id_estoque = estoques.id_estoque"


    If aPrazo.State = 1 Then aPrazo.Close
    aPrazo.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Geral = 0
    While Not aPrazo.EOF
        If VarType(aPrazo("quant")) <> vbNull Then Geral = Geral + aPrazo("quant")
        aPrazo.MoveNext
    Wend

    lblTotalGeral.Caption = Format(Geral, "###,###,##0.00")

    If aPrazo.State = 1 Then aPrazo.Close
    Set aPrazo = Nothing


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub ListProduto_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    Set Estoque = CreateObject("ADODB.Recordset")

    txtid_estoque.text = ListProduto.SelectedItem.Tag

    If txtid_estoque.text <> "" Then
        Sql = " SELECT id_Estoque, descricao FROM Estoques where id_Estoque = '" & txtid_estoque.text & "'"

        If Estoque.State = 1 Then Estoque.Close
        Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Estoque.RecordCount > 0 Then
            If VarType(Estoque("descricao")) <> vbNull Then lblProduto.Caption = Estoque("descricao") Else lblProduto.Caption = ""
            Lista ("")
        End If
    End If

    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
