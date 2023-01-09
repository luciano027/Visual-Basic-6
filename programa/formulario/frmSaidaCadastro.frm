VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaidaCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saída Cadastro"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtid_saida 
      Height          =   285
      Left            =   7800
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   11295
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   9495
      End
      Begin MSComCtl2.DTPicker txtDataaCerto 
         Height          =   315
         Left            =   9840
         TabIndex        =   10
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108658689
         CurrentDate     =   41879
      End
      Begin VB.Label Label11 
         Caption         =   "Historico"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Data Saida"
         Height          =   255
         Left            =   9840
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dados acerto Estoque"
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
         TabIndex        =   11
         Top             =   0
         Width           =   11655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7800
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_saidaacerto 
      Height          =   285
      Left            =   7800
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame2"
      Height          =   5620
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11295
      Begin Vendas.VistaButton cmdIncluirItem 
         Height          =   375
         Left            =   10800
         TabIndex        =   1
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
         Picture         =   "frmSaidaCadastro.frx":0000
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin MSComctlLib.ListView ListaSaida 
         Height          =   5010
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8837
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
         TabIndex        =   3
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
         Picture         =   "frmSaidaCadastro.frx":0352
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
         Caption         =   "Itens "
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
         TabIndex        =   4
         Top             =   0
         Width           =   11295
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   7950
      Width           =   11565
      _ExtentX        =   20399
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
      Left            =   10320
      TabIndex        =   15
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
      Picture         =   "frmSaidaCadastro.frx":06A4
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
      Left            =   9120
      TabIndex        =   16
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Excluir"
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
      Picture         =   "frmSaidaCadastro.frx":07AE
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
      TabIndex        =   19
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label lbltotalReceita 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   18
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8280
      Left            =   0
      Picture         =   "frmSaidaCadastro.frx":0D00
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "frmSaidaCadastro"
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


Private Sub Form_Activate()
    AtualizaNF
    '
End Sub

Private Sub Form_Load()
    Me.Width = 11655
    Me.Height = 8700
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu

    txtdataAcerto.Value = Now

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmSaidaCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSaidaCadastro = Nothing
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
    Sql = " SELECT saidaacerto.*, saida.id_saidaAcerto, saida.id_estoque,"
    Sql = Sql & " saida.preco_venda, saida.quantidade,"
    Sql = Sql & " Estoques.id_estoque , Estoques.unidade, Estoques.descricao"
    Sql = Sql & " From"
    Sql = Sql & " saidaacerto"
    Sql = Sql & " LEFT JOIN saida ON saidaacerto.id_saidaAcerto = saida.id_saidaAcerto"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " where "
    Sql = Sql & " saida.id_saidaacerto = '" & txtid_saidaAcerto.text & "'"

    If Estoque.State = 1 Then Estoque.Close
    Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Estoque.RecordCount > 0 Then
        If VarType(Estoque("historico")) <> vbNull Then txtHistorico.text = Estoque("historico")
        If VarType(Estoque("dataacerto")) <> vbNull Then txtdataAcerto.Value = Format(Estoque("dataacerto"), "DD/MM/YYYY")
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

    ' abre um Recrodset da Tabela estoque
    Sql = " SELECT saidaacerto.*, saida.id_saidaAcerto, saida.id_estoque,"
    Sql = Sql & " saida.preco_custo, saida.quantidade, saida.id_saida,"
    Sql = Sql & " Estoques.id_estoque , Estoques.unidade, Estoques.descricao"
    Sql = Sql & " From"
    Sql = Sql & " saidaacerto"
    Sql = Sql & " LEFT JOIN saida ON saidaacerto.id_saidaAcerto = saida.id_saidaAcerto"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " where "
    Sql = Sql & " saida.id_saidaacerto = '" & txtid_saidaAcerto.text & "'"

    ' abre um Recrodset da Tabela NFs
    If NFs.State = 1 Then NFs.Close
    NFs.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaSaida.ColumnHeaders.Clear
    ListaSaida.ListItems.Clear

    Mitem = 0

    If NFs.RecordCount = 0 Then
        Exit Sub
    End If

    ListaSaida.ColumnHeaders.Add , , "Referência", 1300
    ListaSaida.ColumnHeaders.Add , , "Descrição", 5000
    ListaSaida.ColumnHeaders.Add , , "Unidade", 900, lvwColumnCenter
    ListaSaida.ColumnHeaders.Add , , "Quantidade", 1100, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Preço Custo(R$)", 1100, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Total (R$)", 1100, lvwColumnRight

    If NFs.BOF = True And NFs.EOF = True Then Exit Sub

    mtotalI = 0
    mTotalG = 0


    While Not NFs.EOF

        If VarType(NFs("preco_custo")) <> vbNull And VarType(NFs("quantidade")) <> vbNull Then
            mtotalI = NFs("preco_custo") * NFs("quantidade")
        Else
            mtotalI = 0
        End If

        If VarType(NFs("id_estoque")) <> vbNull Then Set itemx = ListaSaida.ListItems.Add(, , NFs("id_estoque"))
        If VarType(NFs("descricao")) <> vbNull Then itemx.SubItems(1) = NFs("descricao") Else itemx.SubItems(1) = ""
        If VarType(NFs("unidade")) <> vbNull Then itemx.SubItems(2) = NFs("unidade") Else itemx.SubItems(2) = ""
        If VarType(NFs("quantidade")) <> vbNull Then itemx.SubItems(3) = Format(NFs("quantidade"), "###,##0.0000") Else itemx.SubItems(3) = ""
        If VarType(NFs("preco_custo")) <> vbNull Then itemx.SubItems(4) = Format(NFs("preco_custo"), "###,##0.00") Else itemx.SubItems(4) = ""
        If mtotalI > 0 Then itemx.SubItems(5) = Format(mtotalI, "###,##0.00") Else itemx.SubItems(5) = ""
        If VarType(NFs("id_Saida")) <> vbNull Then itemx.Tag = NFs("id_Saida")
        NFs.MoveNext

        mTotalG = mTotalG + mtotalI

        Mitem = Mitem + 1
    Wend

    lbltotalReceita.Caption = Format(mTotalG, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaSaida, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If NFs.State = 1 Then NFs.Close
    Set NFs = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaSaida_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_saida.text = ListaSaida.SelectedItem.Tag

    If txtid_saida.text <> "" Then cmdExcluirItem.Enabled = True

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro
    Dim Excluir As Boolean

    If Mitem = 0 Then

        Sqlconsulta = "id_saidaacerto = '" & txtid_saidaAcerto.text & "'"
        confirma = MsgBox("Confirma Exclusão do documento saida", vbQuestion + vbYesNo, "Excluir")
        If confirma = vbYes Then
            sqlDeletar "saida", Sqlconsulta, Me, "N"
            sqlDeletar "saidaacerto", Sqlconsulta, Me, "S"
            MsgBox ("documento excluido com sucesso..."), vbInformation
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



Private Sub cmdIncluirItem_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double

    If txtid_saidaAcerto.text = "" Then
        campo = "dataacerto"
        Scampo = "'" & Format(txtdataAcerto.Value, "YYYYMMDD") & "'"

        If txtHistorico.text <> "" Then
            campo = campo & ", historico"
            Scampo = Scampo & ", '" & Mid(txtHistorico.text, 1, 100) & "'"
        Else
            campo = campo & ", historico"
            Scampo = Scampo & ", '" & "Acerto no estoque dia: " & Format(txtdataAcerto.Value, "DD/MM/YYYY") & "'"
        End If

        sqlIncluir "saidaacerto", campo, Scampo, Me, "N"

        Buscar_id

    End If

    If txtid_saidaAcerto.text <> "" Then
        With frmsaidaCadastroItem
            .txtid_saidaAcerto.text = txtid_saidaAcerto.text
            .txtdataAcerto.text = txtdataAcerto.Value
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

    Sql = "SELECT max(id_saidaacerto) as MaxID "
    Sql = Sql & " FROM saidaacerto"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then txtid_saidaAcerto.text = Tabela("maxid") Else txtid_saidaAcerto.text = ""
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

    confirma = MsgBox("Confirma Exclusão do item", vbQuestion + vbYesNo, "Excluir")
    If confirma = vbYes Then

        ExcluirSaldo

        Sqlconsulta = "id_Saida = '" & txtid_saida.text & "'"

        sqlDeletar "Saida", Sqlconsulta, Me, "S"

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

    Sqlconsulta = "id_Saida = '" & txtid_saida.text & "'"

    Sql = "select Saida.quantidade, Saida.id_estoque"
    Sql = Sql & " from"
    Sql = Sql & " Saida"
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





