VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntradaCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtid_entradaitens 
      Height          =   285
      Left            =   4680
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame2"
      Height          =   4575
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   11295
      Begin Vendas.VistaButton cmdIncluirItem 
         Height          =   375
         Left            =   10800
         TabIndex        =   19
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
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
         Picture         =   "frmEntradaCadastro.frx":0000
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin MSComctlLib.ListView ListaNFs 
         Height          =   3975
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7011
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
      Begin Vendas.VistaButton cmdExcluirItem 
         Height          =   375
         Left            =   10800
         TabIndex        =   20
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
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
         Picture         =   "frmEntradaCadastro.frx":0352
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Itens da NF"
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
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11295
      End
   End
   Begin VB.TextBox txtid_Entrada 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4920
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      Begin VB.TextBox txtfornecedor 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   10575
      End
      Begin VB.TextBox txtid_fornecedor 
         Height          =   285
         Left            =   5760
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNFNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8760
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   6975
      End
      Begin MSComCtl2.DTPicker txtNFdata 
         Height          =   315
         Left            =   7320
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108658689
         CurrentDate     =   41879
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dados NF"
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
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11655
      End
      Begin VB.Label Label13 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4935
      End
      Begin VB.Image cmdConsultaFornecedor 
         Height          =   315
         Left            =   10800
         Picture         =   "frmEntradaCadastro.frx":06A4
         Stretch         =   -1  'True
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Data NF"
         Height          =   255
         Left            =   7320
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Nº da NF"
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Historico"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7920
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   30763
            MinWidth        =   30763
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   9720
      TabIndex        =   21
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
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
      Picture         =   "frmEntradaCadastro.frx":09AE
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdExcluir 
      Height          =   615
      Left            =   7920
      TabIndex        =   22
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "         Excluir"
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
      Picture         =   "frmEntradaCadastro.frx":0AB8
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdEntradaPagar 
      Height          =   615
      Left            =   6000
      TabIndex        =   25
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Contas a Pagar"
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
      Picture         =   "frmEntradaCadastro.frx":100A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   24
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label lbltotalReceita 
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
      Left            =   120
      TabIndex        =   23
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8280
      Left            =   0
      Picture         =   "frmEntradaCadastro.frx":155C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "frmEntradaCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' variaveis do modulo
Dim Sqlconsulta As String
Dim confirma As String
Dim Scampo As String
Dim campo As String
Dim ChaveM As String
Dim Sql As String
Dim SQsort As String
Dim sqlwhere As String
Dim SaldoAn As Double
Dim Mitem As Integer


Private Sub cmdEntradaPagar_Click()
    With frmPagarCadastro
        .txtid_fornecedor.text = txtid_fornecedor.text
        .txtFornecedor.text = txtFornecedor.text
        .txtDocumento.text = txtNFNro.text
        .txtHistorico.text = txtHistorico.text
        .Show 1
    End With
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro
    Dim Excluir As Boolean

    If Mitem = 0 Then

        Sqlconsulta = "id_entrada = '" & txtid_entrada.text & "'"
        confirma = MsgBox("Confirma Exclusão da NF", vbQuestion + vbYesNo, "Excluir")
        If confirma = vbYes Then
            sqlDeletar "entrada", Sqlconsulta, Me, "N"
            sqlDeletar "entradaitens", Sqlconsulta, Me, "S"
            MsgBox ("NF excluida com sucesso..."), vbInformation
            Unload Me
        End If
    Else
        MsgBox ("Favor excluir os itens primeiro..."), vbInformation
        Exit Sub
    End If
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Form_Activate()
    AtualizaNF
    '
End Sub

Private Sub Form_Load()
    Me.Width = 11820
    Me.Height = 8910
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu

    txtNFdata.Value = Now

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmEstoqueCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEstoqueCadastro = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------- campos do formulario--------------------------------------------------

Private Sub AtualizaNF()
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    ' conecta ao banco de dados

    Set Estoque = CreateObject("ADODB.Recordset")    '''

    ' abre um Recrodset da Tabela estoque
    Sql = " SELECT entrada.*, fornecedores.id_fornecedor, fornecedores.fornecedor"
    Sql = Sql & " From"
    Sql = Sql & " Entrada"
    Sql = Sql & " LEFT JOIN fornecedores ON entrada.id_fornecedor = fornecedores.id_fornecedor"
    Sql = Sql & " where "
    Sql = Sql & " entrada.id_entrada = '" & txtid_entrada.text & "'"

    If Estoque.State = 1 Then Estoque.Close
    Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Estoque.RecordCount > 0 Then
        If VarType(Estoque("id_fornecedor")) <> vbNull Then txtid_fornecedor.text = Estoque("id_fornecedor")
        If VarType(Estoque("fornecedor")) <> vbNull Then txtFornecedor.text = Estoque("fornecedor")
        If VarType(Estoque("historico")) <> vbNull Then txtHistorico.text = Estoque("historico")
        If VarType(Estoque("nfdata")) <> vbNull Then txtNFdata.Value = Format(Estoque("nfdata"), "DD/MM/YYYY")
        If VarType(Estoque("nfnro")) <> vbNull Then txtNFNro.text = Estoque("nfnro")
    End If
    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim NFs As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mTotalG As Double
    Dim mtotalI As Double
    ' conecta ao banco de dados
    Set NFs = CreateObject("ADODB.Recordset")

    Sql = "SELECT entradaitens.*, estoques.id_estoque,"
    Sql = Sql & " Estoques.DESCRICAO , Estoques.UNIDADE"
    Sql = Sql & " From"
    Sql = Sql & " entradaitens"
    Sql = Sql & " LEFT JOIN estoques ON entradaitens.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where "
    Sql = Sql & " entradaitens.id_entrada = '" & txtid_entrada.text & "'"

    ' abre um Recrodset da Tabela NFs
    If NFs.State = 1 Then NFs.Close
    NFs.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaNFs.ColumnHeaders.Clear
    ListaNFs.ListItems.Clear

    Mitem = 0

    If NFs.RecordCount = 0 Then
        Exit Sub
    End If

    ListaNFs.ColumnHeaders.Add , , "Referência", 1300
    ListaNFs.ColumnHeaders.Add , , "Descrição", 5000
    ListaNFs.ColumnHeaders.Add , , "Unidade", 900, lvwColumnCenter
    ListaNFs.ColumnHeaders.Add , , "Quantidade", 1100, lvwColumnRight
    ListaNFs.ColumnHeaders.Add , , "Preço (R$)", 1100, lvwColumnRight
    ListaNFs.ColumnHeaders.Add , , "Total (R$)", 1100, lvwColumnRight

    If NFs.BOF = True And NFs.EOF = True Then Exit Sub

    mtotalI = 0
    mTotalG = 0


    While Not NFs.EOF

        If VarType(NFs("precocusto")) <> vbNull And VarType(NFs("quantidade")) <> vbNull Then
            mtotalI = NFs("precocusto") * NFs("quantidade")
        Else
            mtotalI = 0
        End If

        If VarType(NFs("id_estoque")) <> vbNull Then Set itemx = ListaNFs.ListItems.Add(, , NFs("id_estoque"))
        If VarType(NFs("descricao")) <> vbNull Then itemx.SubItems(1) = NFs("descricao") Else itemx.SubItems(1) = ""
        If VarType(NFs("unidade")) <> vbNull Then itemx.SubItems(2) = NFs("unidade") Else itemx.SubItems(2) = ""
        If VarType(NFs("quantidade")) <> vbNull Then itemx.SubItems(3) = Format(NFs("quantidade"), "###,##0.0000") Else itemx.SubItems(3) = ""
        If VarType(NFs("precocusto")) <> vbNull Then itemx.SubItems(4) = Format(NFs("precocusto"), "###,##0.00") Else itemx.SubItems(4) = ""
        If mtotalI > 0 Then itemx.SubItems(5) = Format(mtotalI, "###,##0.00") Else itemx.SubItems(5) = ""
        If VarType(NFs("id_entradaitens")) <> vbNull Then itemx.Tag = NFs("id_entradaitens")
        NFs.MoveNext

        mTotalG = mTotalG + mtotalI

        Mitem = Mitem + 1
    Wend

    lbltotalReceita.Caption = Format(mTotalG, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaNFs, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If NFs.State = 1 Then NFs.Close
    Set NFs = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub ListaNFs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_entradaitens.text = ListaNFs.SelectedItem.Tag

    If txtid_entradaitens.text <> "" Then cmdExcluirItem.Enabled = True

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub cmdIncluirItem_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double

    If txtid_entrada.text = "" Then
        campo = "nfdata"
        Scampo = "'" & Format(txtNFdata.Value, "YYYYMMDD") & "'"

        If txtid_fornecedor.text <> "" Then
            campo = campo & ", id_fornecedor"
            Scampo = Scampo & ", '" & txtid_fornecedor.text & "'"
        Else
            MsgBox ("Favor selecionar um fonrecedor..")
            txtid_fornecedor.SetFocus
            Exit Sub
        End If

        If txtNFNro.text <> "" Then
            campo = campo & ", nfnro"
            Scampo = Scampo & ", '" & txtNFNro.text & "'"
        End If

        If txtHistorico.text <> "" Then
            campo = campo & ", historico"
            Scampo = Scampo & ", '" & txtHistorico.text & "'"
        End If

        sqlIncluir "Entrada", campo, Scampo, Me, "N"

        Buscar_id

    End If

    If txtid_entrada.text <> "" Then
        With frmEntradaCadastroItem
            .txtid_entrada.text = txtid_entrada.text
            .txtdataEntrada.text = txtNFdata.Value
            .txtid_fornecedor.text = txtid_fornecedor.text
            .Show 1
        End With
    End If

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(id_entrada) as MaxID "
    Sql = Sql & " FROM entrada"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then txtid_entrada.text = Tabela("maxid") Else txtid_entrada.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


Private Sub cmdExcluirItem_Click()
    On Error GoTo trata_erro
    Dim Excluir As Boolean

    confirma = MsgBox("Confirma Exclusão do item da NF", vbQuestion + vbYesNo, "Excluir")
    If confirma = vbYes Then

        ExcluirSaldo

        Sqlconsulta = "id_entradaitens = '" & txtid_entradaitens.text & "'"

        sqlDeletar "EntradaItens", Sqlconsulta, Me, "S"

    End If

    cmdExcluirItem.Enabled = False

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub ExcluirSaldo()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim mSaldoE As Double
    Dim mIdestoque As String
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sqlconsulta = "id_entradaitens = '" & txtid_entradaitens.text & "'"

    Sql = "select entradaitens.quantidade, entradaitens.id_estoque"
    Sql = Sql & " from"
    Sql = Sql & " entradaitens"
    Sql = Sql & " where "
    Sql = Sql & Sqlconsulta

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("quantidade")) <> vbNull Then SaldoAn = Tabela("quantidade")
        If VarType(Tabela("id_estoque")) <> vbNull Then mIdestoque = Tabela("id_estoque")
    Else
        SaldoAn = 0
    End If

    ' Altera saldo na tabela Grupo
    Sqlconsulta = " id_estoque = '" & mIdestoque & "'"

    Sql = "SELECT estoquesaldo.* "
    Sql = Sql & " FROM estoquesaldo"
    Sql = Sql & " where"
    Sql = Sql & Sqlconsulta

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("saldo")) <> vbNull Then mSaldo = Tabela("saldo") Else mSaldo = "0"

        mSaldo = mSaldo - SaldoAn
        If mSaldo < 0 Then mSaldo = 0

        campo = " saldo = '" & FormatValor(mSaldo, 1) & "'"

        sqlAlterar "Estoquesaldo", campo, Sqlconsulta, Me, "N"

    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultaFornecedor_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaFornecedor
        .Show 1
    End With

    If IDFornecedor <> "" Then
        txtid_fornecedor.text = IDFornecedor
        IDFornecedor = ""
    End If

    Sql = "SELECT id_fornecedor, fornecedor FROM fornecedores WHERE id_fornecedor = '" & txtid_fornecedor.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_fornecedor")) <> vbNull Then txtid_fornecedor.text = Tabela("id_fornecedor") Else txtid_fornecedor.text = ""
        If VarType(Tabela("fornecedor")) <> vbNull Then txtFornecedor.text = Tabela("fornecedor") Else txtFornecedor.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



