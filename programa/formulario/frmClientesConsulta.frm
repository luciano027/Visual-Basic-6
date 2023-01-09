VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientesConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1680
      TabIndex        =   8
      Top             =   6960
      Width           =   5655
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.Image cmdConsultar 
         Height          =   315
         Left            =   5280
         Picture         =   "frmClientesConsulta.frx":0000
         Stretch         =   -1  'True
         Top             =   480
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
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   285
      Left            =   11280
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   11640
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12000
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativo"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
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
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListaClientes 
      Height          =   6495
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11456
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
      Top             =   8175
      Width           =   13530
      _ExtentX        =   23865
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
      Left            =   12240
      TabIndex        =   15
      Top             =   7320
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
      Picture         =   "frmClientesConsulta.frx":030A
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
      Left            =   9840
      TabIndex        =   16
      Top             =   7320
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
      Picture         =   "frmClientesConsulta.frx":0414
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
      Left            =   8640
      TabIndex        =   17
      Top             =   7320
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
      Picture         =   "frmClientesConsulta.frx":0966
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
      Left            =   7440
      TabIndex        =   18
      Top             =   7320
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
      Picture         =   "frmClientesConsulta.frx":0CB8
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdLimite 
      Height          =   615
      Left            =   11040
      TabIndex        =   19
      Top             =   7320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Limite"
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
      Picture         =   "frmClientesConsulta.frx":100A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
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
      Left            =   10200
      TabIndex        =   14
      Top             =   120
      Width           =   3135
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
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   10095
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   -120
      Picture         =   "frmClientesConsulta.frx":155C
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   13680
   End
End
Attribute VB_Name = "frmClientesConsulta"
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



Private Sub cmdLimite_Click()
    With frmVendedorCodigo
        .txtTipo.text = "L"
        .txtACesso.text = ""
        .txtid_cliente.text = txtid_cliente.text
        .Show 1
    End With

    cmdConsultar_Click
End Sub

Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Me.Width = 13620
    Me.Height = 8985
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmClientesConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmClientesConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAlterar_Click()
    With frmClientesCadastro
        .txtid_cliente.text = txtid_cliente.text
        .txtTipo.text = "A"
        .txtLimite.Enabled = False
        .Show 1
    End With
    cmdConsultar_Click
    'Lista ("")
End Sub

Private Sub cmdExcluir_Click()
    With frmClientesCadastro
        .txtid_cliente.text = txtid_cliente.text
        .txtTipo.text = "E"
        .Show 1
    End With
    cmdConsultar_Click
    'Lista ("")
End Sub

Private Sub cmdIncluir_Click()
    With frmClientesCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    cmdConsultar_Click
    'Lista ("")
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim clientes As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set clientes = CreateObject("ADODB.Recordset")

    If SQconsulta = "" Then
        Sql = "SELECT clientes.* "
        Sql = Sql & " FROM clientes "
        Sql = Sql & " order by cliente"
    Else
        Sql = "SELECT clientes.*"
        Sql = Sql & " FROM clientes "
        Sql = Sql & " Where " & SQconsulta
        Sql = Sql & " order by cliente"
    End If

    ' abre um Recrodset da Tabela clientes
    If clientes.State = 1 Then clientes.Close
    clientes.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaClientes.ColumnHeaders.Clear
    ListaClientes.ListItems.Clear

    If clientes.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "clientes(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "cliente(s) encontrado(s): " & clientes.RecordCount

    ListaClientes.ColumnHeaders.Add , , "Descrição", 12900

    If clientes.BOF = True And clientes.EOF = True Then Exit Sub
    While Not clientes.EOF

        If VarType(clientes("cliente")) <> vbNull Then Set itemx = ListaClientes.ListItems.Add(, , clientes("cliente"))
        If VarType(clientes("id_cliente")) <> vbNull Then itemx.Tag = clientes("id_cliente")
        clientes.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaClientes, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If clientes.State = 1 Then clientes.Close
    Set clientes = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaClientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_cliente.text = ListaClientes.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarNome_Cliente
    Aguarde_Process Me, False
End Sub

Private Sub consultarNome_Cliente()

    If optAtivo.Value = True Then Sqlconsulta = " status = 'A'"
    If optInativo.Value = True Then Sqlconsulta = " status = 'I'"
    If optTodos.Value = True Then Sqlconsulta = " 1=1 "

    If txtConsulta.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and clientes.cliente like '%" & txtConsulta.text & "%'"
        Sqlconsulta = Sqlconsulta & " or clientes.cnpj like '%" & txtConsulta.text & "%'"
    End If

    Lista (Sqlconsulta)

End Sub

Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub
