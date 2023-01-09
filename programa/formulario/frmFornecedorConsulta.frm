VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFornecedorConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fornecedor"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
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
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10920
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   10560
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_fornecedor 
      Height          =   285
      Left            =   10200
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   6960
      Width           =   5535
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
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
         Width           =   6855
      End
      Begin VB.Image cmdConsultar 
         Height          =   315
         Left            =   5040
         Picture         =   "frmFornecedorConsulta.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView ListaFornecedores 
      Height          =   6495
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
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
      Top             =   8265
      Width           =   12345
      _ExtentX        =   21775
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
      Left            =   11040
      TabIndex        =   13
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
      Picture         =   "frmFornecedorConsulta.frx":030A
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
      TabIndex        =   14
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
      Picture         =   "frmFornecedorConsulta.frx":0414
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
      TabIndex        =   15
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
      Picture         =   "frmFornecedorConsulta.frx":0966
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
      TabIndex        =   16
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
      Picture         =   "frmFornecedorConsulta.frx":0CB8
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
      Width           =   9255
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
      Left            =   9480
      TabIndex        =   17
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmFornecedorConsulta.frx":100A
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   12480
   End
End
Attribute VB_Name = "frmFornecedorConsulta"
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
    Me.Width = 12435
    Me.Height = 9015
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFornecedorConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFornecedorConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAlterar_Click()
    With frmFornecedorCadastro
        .txtid_fornecedor.text = txtid_fornecedor.text
        .txtTipo.text = "A"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdExcluir_Click()
    With frmFornecedorCadastro
        .txtid_fornecedor.text = txtid_fornecedor.text
        .txtTipo.text = "E"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdIncluir_Click()
    With frmFornecedorCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Fornecedores As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Fornecedores = CreateObject("ADODB.Recordset")

    If SQconsulta = "" Then
        Sql = "SELECT Fornecedores.* "
        Sql = Sql & " FROM Fornecedores "
        Sql = Sql & " order by Fornecedor"
    Else
        Sql = "SELECT Fornecedores.*"
        Sql = Sql & " FROM Fornecedores "
        Sql = Sql & " Where " & SQconsulta
        Sql = Sql & " order by Fornecedor"
    End If

    ' abre um Recrodset da Tabela Fornecedores
    If Fornecedores.State = 1 Then Fornecedores.Close
    Fornecedores.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaFornecedores.ColumnHeaders.Clear
    ListaFornecedores.ListItems.Clear

    If Fornecedores.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Fornecedores(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "Fornecedor(s) encontrado(s): " & Fornecedores.RecordCount

    ListaFornecedores.ColumnHeaders.Add , , "Descrição", 11500

    If Fornecedores.BOF = True And Fornecedores.EOF = True Then Exit Sub
    While Not Fornecedores.EOF

        If VarType(Fornecedores("Fornecedor")) <> vbNull Then Set itemx = ListaFornecedores.ListItems.Add(, , Fornecedores("Fornecedor"))
        If VarType(Fornecedores("id_fornecedor")) <> vbNull Then itemx.Tag = Fornecedores("id_fornecedor")
        Fornecedores.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaFornecedores, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Fornecedores.State = 1 Then Fornecedores.Close
    Set Fornecedores = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaFornecedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_fornecedor.text = ListaFornecedores.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarNome_Fornecedor
    Aguarde_Process Me, False
End Sub

Private Sub consultarNome_Fornecedor()

    If optAtivo.Value = True Then Sqlconsulta = " status = 'A'"
    If optInativo.Value = True Then Sqlconsulta = " status = 'I'"
    If optTodos.Value = True Then Sqlconsulta = " 1=1 "

    If txtConsulta.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and Fornecedores.Fornecedor like '%" & txtConsulta.text & "%'"
        Sqlconsulta = Sqlconsulta & " or Fornecedores.cnpj like '%" & txtConsulta.text & "%'"
    End If

    Lista (Sqlconsulta)

End Sub

Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub
