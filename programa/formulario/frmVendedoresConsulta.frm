VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendedoresConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendedores"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTipo_acesso 
      Height          =   285
      Left            =   8400
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
      Begin VB.OptionButton optADM 
         Caption         =   "Administrador"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativo"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
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
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   9480
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   9120
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   6480
      Width           =   9015
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8775
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9015
      End
   End
   Begin MSComctlLib.ListView ListaVendedores 
      Height          =   6015
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10610
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
      TabIndex        =   12
      Top             =   8205
      Width           =   11130
      _ExtentX        =   19632
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
      Left            =   9720
      TabIndex        =   13
      Top             =   7560
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
      Picture         =   "frmVendedoresConsulta.frx":0000
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
      Left            =   8520
      TabIndex        =   14
      Top             =   7560
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
      Picture         =   "frmVendedoresConsulta.frx":010A
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
      Left            =   7320
      TabIndex        =   15
      Top             =   7560
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
      Picture         =   "frmVendedoresConsulta.frx":065C
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
      Left            =   6120
      TabIndex        =   16
      Top             =   7560
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
      Picture         =   "frmVendedoresConsulta.frx":09AE
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
      Left            =   4680
      TabIndex        =   20
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
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
      Picture         =   "frmVendedoresConsulta.frx":0D00
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
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
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   7575
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
      Left            =   7800
      TabIndex        =   17
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmVendedoresConsulta.frx":1052
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   11160
   End
End
Attribute VB_Name = "frmVendedoresConsulta"
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



Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Me.Width = 11220
    Me.Height = 9015
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmVendedoresConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendedoresConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAlterar_Click()
    With frmVendedoresCadastro
        .txtid_vendedor.text = txtid_vendedor.text
        .txtTipo.text = "A"
        .txtTipo_acesso.text = txtTipo_acesso.text
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdExcluir_Click()
    With frmVendedoresCadastro
        .txtid_vendedor.text = txtid_vendedor.text
        .txtTipo.text = "E"
        .txtTipo_acesso.text = txtTipo_acesso.text
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdIncluir_Click()
    With frmVendedoresCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim vendedores As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set vendedores = CreateObject("ADODB.Recordset")

    If SQconsulta = "" Then
        Sql = "SELECT Vendedores.* "
        Sql = Sql & " FROM Vendedores "
        Sql = Sql & " order by Vendedor"
    Else
        Sql = "SELECT Vendedores.*"
        Sql = Sql & " FROM Vendedores "
        Sql = Sql & " Where " & SQconsulta
        Sql = Sql & " order by Vendedor"
    End If

    ' abre um Recrodset da Tabela Vendedores
    If vendedores.State = 1 Then vendedores.Close
    vendedores.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaVendedores.ColumnHeaders.Clear
    ListaVendedores.ListItems.Clear

    If vendedores.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Vendedores(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "Vendedor(s) encontrado(s): " & vendedores.RecordCount

    ListaVendedores.ColumnHeaders.Add , , "Nome", 7500
    ListaVendedores.ColumnHeaders.Add , , "Tipo", 3000

    If vendedores.BOF = True And vendedores.EOF = True Then Exit Sub
    While Not vendedores.EOF

        If VarType(vendedores("Vendedor")) <> vbNull Then Set itemx = ListaVendedores.ListItems.Add(, , vendedores("Vendedor"))
        If VarType(vendedores("tipo_Acesso")) <> vbNull Then
            If vendedores("tipo_Acesso") = "" Then itemx.SubItems(1) = "Vendedor"
            If vendedores("tipo_Acesso") = "A" Then itemx.SubItems(1) = "Administrador"
            If vendedores("tipo_Acesso") = "P" Then itemx.SubItems(1) = "Caixa"
        End If
        If VarType(vendedores("tipo_Acesso")) = vbNull Then itemx.SubItems(1) = "Vendedor"
        If VarType(vendedores("id_vendedor")) <> vbNull Then itemx.Tag = vendedores("id_vendedor")
        vendedores.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaVendedores, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If vendedores.State = 1 Then vendedores.Close
    Set vendedores = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaVendedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim vendedores As ADODB.Recordset
    ' conecta ao banco de dados

    Set vendedores = CreateObject("ADODB.Recordset")    '''

    txtid_vendedor.text = ListaVendedores.SelectedItem.Tag

    ' abre um Recrodset da Tabela Vendedores
    Sql = " select "
    Sql = Sql & " Vendedores.*"
    Sql = Sql & " from  "
    Sql = Sql & " Vendedores "
    Sql = Sql & " where "
    Sql = Sql & " id_vendedor = '" & txtid_vendedor.text & "'"

    If vendedores.State = 1 Then vendedores.Close
    vendedores.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If vendedores.RecordCount > 0 Then

        If VarType(vendedores("tipo_acesso")) <> vbNull Then txtTipo_acesso.text = vendedores("tipo_acesso") Else txtTipo_acesso.text = ""

    End If
    If vendedores.State = 1 Then vendedores.Close
    Set vendedores = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarNome_Vendedor
    Aguarde_Process Me, False
End Sub

Private Sub consultarNome_Vendedor()

    If optAtivo.Value = True Then Sqlconsulta = " status = 'A'"
    If optInativo.Value = True Then Sqlconsulta = " status = 'I'"
    If optTodos.Value = True Then Sqlconsulta = " 1=1 "
    If optADM.Value = True Then Sqlconsulta = " tipo_acesso = 'A'"

    If txtConsulta.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and Vendedores.Vendedor like '%" & txtConsulta.text & "%'"
    End If

    Lista (Sqlconsulta)

End Sub

Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub
