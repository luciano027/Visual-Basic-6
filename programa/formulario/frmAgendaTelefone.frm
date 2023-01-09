VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgendaTelefone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda Telefone"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAgendaTelefone.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   11475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   11175
      Begin VB.ComboBox cmbConsulta 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAgendaTelefone.frx":0376
         Left            =   120
         List            =   "frmAgendaTelefone.frx":0386
         TabIndex        =   10
         Text            =   "Nome"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   480
         Width           =   6975
      End
      Begin VB.Image cmdConsulta 
         Height          =   360
         Left            =   10680
         Picture         =   "frmAgendaTelefone.frx":03AE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListPed 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin VB.TextBox txtid_telefone 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComctlLib.TabStrip AgLetra 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   26
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "A"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "B"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "C"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "D"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "F"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "G"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "H"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "I"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "J"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "K"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "L"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "M"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "N"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "O"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Q"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "R"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "S"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "T"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab21 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "U"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab22 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "V"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab23 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab24 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Y"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab25 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Z"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab26 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Geral"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8475
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21944
            MinWidth        =   21944
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   10200
      TabIndex        =   13
      Top             =   7800
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
      Picture         =   "frmAgendaTelefone.frx":06B8
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
      Left            =   9000
      TabIndex        =   14
      Top             =   7800
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
      Picture         =   "frmAgendaTelefone.frx":07C2
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdAlterar 
      Height          =   615
      Left            =   7800
      TabIndex        =   15
      Top             =   7800
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
      Picture         =   "frmAgendaTelefone.frx":0D14
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdIncluir 
      Height          =   615
      Left            =   6600
      TabIndex        =   16
      Top             =   7800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Incluir"
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
      Picture         =   "frmAgendaTelefone.frx":1066
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label lbltotal 
      BackStyle       =   0  'Transparent
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
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   8985
      Left            =   0
      Picture         =   "frmAgendaTelefone.frx":13B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "frmAgendaTelefone"
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
Dim chave As String
Dim Sql As String


Private Sub cmbConsulta_Click()
    Dim sTipo As String

    Select Case cmbConsulta.ListIndex
    Case 0: sTipo = "Nome"
    Case 1: sTipo = "Telefone"
    Case 2: sTipo = "Celular"
    Case 3: sTipo = "Atividade"
    End Select

End Sub


Private Sub cmdConsulta_Click()

    If cmbConsulta.ListIndex = 0 Or cmbConsulta.text = "Nome" Then Sqlconsulta = " Nome Like '%" & txtConsulta.text & "%'"
    If cmbConsulta.ListIndex = 1 Then Sqlconsulta = " telefone Like '%" & txtConsulta.text & "%'"
    If cmbConsulta.ListIndex = 2 Then Sqlconsulta = " celular Like '%" & txtConsulta.text & "%'"
    If cmbConsulta.ListIndex = 3 Then Sqlconsulta = " atividade Like '%" & txtConsulta.text & "%'"

    AgLetra.Tabs.Item(26).Selected = True
    ListaGeral (Sqlconsulta)
End Sub

Private Sub Form_Activate()

'

End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadPicture(ICONBD)
    Me.Width = 11565
    Me.Height = 9330
    Centerform Me
    MenuPrincipal.DesabilitaMenu
    ListaGeral ("mid(nome,1,1) = 'A'")

    '  If mAgendaTIn = 1 Then cmdIncluir.Enabled = True Else cmdIncluir.Enabled = False
    '  If mAgendaTCo = 1 Then cmdConsulta.Enabled = True Else cmdConsulta.Enabled = False
    '  If mAgendaTAl = 1 Then cmdAlterar.Enabled = True Else cmdAlterar.Enabled = False
    '  If mAgendaTEx = 1 Then cmdExcluir.Enabled = True Else cmdExcluir.Enabled = False

End Sub


Private Sub cmdAlterar_Click()
    With frmAgendaTelefoneCadastro
        .txtTipo.text = "A"
        .txtid_telefone.text = txtid_telefone.text
        .Show 1
    End With
    ListaGeral ("mid(nome,1,1) = 'A'")
End Sub

Private Sub AgLetra_Click()
    If AgLetra.Tabs.Item(1).Selected = True Then ListaGeral ("mid(nome,1,1) = 'A'")
    If AgLetra.Tabs.Item(2).Selected = True Then ListaGeral ("mid(nome,1,1) = 'B'")
    If AgLetra.Tabs.Item(3).Selected = True Then ListaGeral ("mid(nome,1,1) = 'C'")
    If AgLetra.Tabs.Item(4).Selected = True Then ListaGeral ("mid(nome,1,1) = 'D'")
    If AgLetra.Tabs.Item(5).Selected = True Then ListaGeral ("mid(nome,1,1) = 'E'")
    If AgLetra.Tabs.Item(6).Selected = True Then ListaGeral ("mid(nome,1,1) = 'F'")
    If AgLetra.Tabs.Item(7).Selected = True Then ListaGeral ("mid(nome,1,1) = 'G'")
    If AgLetra.Tabs.Item(8).Selected = True Then ListaGeral ("mid(nome,1,1) = 'H'")
    If AgLetra.Tabs.Item(9).Selected = True Then ListaGeral ("mid(nome,1,1) = 'I'")
    If AgLetra.Tabs.Item(10).Selected = True Then ListaGeral ("mid(nome,1,1) = 'J'")
    If AgLetra.Tabs.Item(11).Selected = True Then ListaGeral ("mid(nome,1,1) = 'K'")
    If AgLetra.Tabs.Item(12).Selected = True Then ListaGeral ("mid(nome,1,1) = 'L'")
    If AgLetra.Tabs.Item(13).Selected = True Then ListaGeral ("mid(nome,1,1) = 'M'")
    If AgLetra.Tabs.Item(14).Selected = True Then ListaGeral ("mid(nome,1,1) = 'N'")
    If AgLetra.Tabs.Item(15).Selected = True Then ListaGeral ("mid(nome,1,1) = 'O'")
    If AgLetra.Tabs.Item(16).Selected = True Then ListaGeral ("mid(nome,1,1) = 'P'")
    If AgLetra.Tabs.Item(17).Selected = True Then ListaGeral ("mid(nome,1,1) = 'Q'")
    If AgLetra.Tabs.Item(18).Selected = True Then ListaGeral ("mid(nome,1,1) = 'R'")
    If AgLetra.Tabs.Item(19).Selected = True Then ListaGeral ("mid(nome,1,1) = 'S'")
    If AgLetra.Tabs.Item(20).Selected = True Then ListaGeral ("mid(nome,1,1) = 'T'")
    If AgLetra.Tabs.Item(21).Selected = True Then ListaGeral ("mid(nome,1,1) = 'U'")
    If AgLetra.Tabs.Item(22).Selected = True Then ListaGeral ("mid(nome,1,1) = 'V'")
    If AgLetra.Tabs.Item(23).Selected = True Then ListaGeral ("mid(nome,1,1) = 'X'")
    If AgLetra.Tabs.Item(24).Selected = True Then ListaGeral ("mid(nome,1,1) = 'Y'")
    If AgLetra.Tabs.Item(25).Selected = True Then ListaGeral ("mid(nome,1,1) = 'Z'")
    If AgLetra.Tabs.Item(26).Selected = True Then
        ListaGeral ("")
    End If

End Sub

Private Sub cmdExcluir_Click()
    With frmAgendaTelefoneCadastro
        .txtTipo.text = "E"
        .txtid_telefone.text = txtid_telefone.text
        .Show 1
    End With
    ListaGeral ("mid(nome,1,1) = 'A'")

End Sub


Private Sub cmdIncluir_Click()
    With frmAgendaTelefoneCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    ListaGeral ("mid(nome,1,1) = 'A'")
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmAgendaTelefone = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub ListaGeral(SQconsulta As String)
    Dim telefone As ADODB.Recordset
    Dim itemx As ListItem
    Dim cond As String
    ' conecta ao banco de dados
    Set telefone = CreateObject("ADODB.Recordset")
    If SQconsulta <> "" Then
        '    sqlConsulta = "mid(nome,1,1) = '" & SQconsulta & "'"
        Sql = "Select * FROM Telefone where " & SQconsulta & " ORDER BY nome"
    Else
        Sql = "Select * FROM Telefone ORDER BY nome"
    End If

    ' abre um Recrodset da Tabela Telefone
    If telefone.State = 1 Then telefone.Close
    telefone.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListPed.ColumnHeaders.Clear
    ListPed.ListItems.Clear

    If telefone.RecordCount = 0 Then
        ListPed.ListItems.Clear
        Exit Sub
    End If

    ListPed.ColumnHeaders.Add , , "Nome", 5200
    ListPed.ColumnHeaders.Add , , "Telefone", 1880
    ListPed.ColumnHeaders.Add , , "Telefone", 1880
    ListPed.ColumnHeaders.Add , , "Celular", 1880
    If telefone.BOF = True And telefone.EOF = True Then Exit Sub
    While Not telefone.EOF
        Set itemx = ListPed.ListItems.Add(, , telefone("nome"))
        If VarType(telefone("telefone")) <> vbNull Then itemx.SubItems(1) = telefone("telefone") Else itemx.SubItems(1) = "   "
        If VarType(telefone("telefone2")) <> vbNull Then itemx.SubItems(2) = telefone("telefone2") Else itemx.SubItems(2) = "   "
        If VarType(telefone("celular")) <> vbNull Then itemx.SubItems(3) = telefone("celular") Else itemx.SubItems(3) = "  "
        itemx.Tag = telefone("id_telefone")
        telefone.MoveNext
    Wend
    lblTotal.Caption = " Existem " & ListPed.ListItems.Count & " Cadastrado(s)"

    'Zebra o listview
    If LVZebra(ListPed, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If telefone.State = 1 Then telefone.Close
    Set telefone = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAgendaTelefone = Nothing
End Sub


Private Sub Listped_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim telefone As ADODB.Recordset
    ' conecta ao banco de dados
    Sql = "SELECT * FROM Telefone WHERE id_telefone = " & ListPed.SelectedItem.Tag
    Set telefone = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela Telefone
    If telefone.State = 1 Then telefone.Close
    telefone.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If telefone.RecordCount > 0 Then
        If VarType(telefone("id_telefone")) <> vbNull Then txtid_telefone.text = telefone("id_telefone") Else txtid_telefone.text = ""
        If VarType(telefone("email")) <> vbNull Then txtemail.text = telefone("email") Else txtemail.text = ""
        '  If mAgendaTAl = 1 Then cmdAlterar.Enabled = True Else cmdAlterar.Enabled = False
        '  If mAgendaTEx = 1 Then cmdExcluir.Enabled = True Else cmdExcluir.Enabled = False
    Else
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
    End If
    If telefone.State = 1 Then telefone.Close
    Set telefone = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
