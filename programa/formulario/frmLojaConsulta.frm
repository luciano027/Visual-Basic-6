VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLojaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loja"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativo"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   11
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
         TabIndex        =   14
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7680
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   7320
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_loja 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   7080
      Width           =   5895
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.Image cmdConsultar 
         Height          =   315
         Left            =   5400
         Picture         =   "frmLojaConsulta.frx":0000
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
         TabIndex        =   2
         Top             =   0
         Width           =   6135
      End
   End
   Begin MSComctlLib.ListView ListaLoja 
      Height          =   6615
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11668
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8235
      Width           =   12585
      _ExtentX        =   22199
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
      Left            =   11280
      TabIndex        =   15
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
      Picture         =   "frmLojaConsulta.frx":0376
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
      Left            =   10080
      TabIndex        =   16
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
      Picture         =   "frmLojaConsulta.frx":0480
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
      Left            =   8880
      TabIndex        =   17
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
      Picture         =   "frmLojaConsulta.frx":09D2
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
      Left            =   7680
      TabIndex        =   18
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
      Picture         =   "frmLojaConsulta.frx":0D24
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
      TabIndex        =   9
      Top             =   120
      Width           =   9015
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
      Left            =   9240
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmLojaConsulta.frx":1076
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "frmLojaConsulta"
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
    frmLojaConsulta.ZOrder (0)
End Sub

Private Sub Form_Load()
    Me.Width = 12675
    Me.Height = 9090
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu

    '    If mLojaconsulta = 1 Then cmdConsultar.Enabled = True Else cmdConsultar.Enabled = False
    '    If mLojaincluir = 1 Then cmdIncluir.Enabled = True Else cmdIncluir.Enabled = False

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmLojaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLojaConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAlterar_Click()
    With frmLojaCadastro
        .txtid_loja.text = txtid_loja.text
        .txtTipo.text = "A"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdExcluir_Click()
    With frmLojaCadastro
        .txtid_loja.text = txtid_loja.text
        .txtTipo.text = "E"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub cmdIncluir_Click()
    With frmLojaCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim loja As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set loja = CreateObject("ADODB.Recordset")

    If SQconsulta = "" Then
        Sql = "SELECT Loja.* "
        Sql = Sql & " FROM Loja "
        Sql = Sql & " order by descricao"
    Else
        Sql = "SELECT Loja.*"
        Sql = Sql & " FROM Loja "
        Sql = Sql & " Where " & SQconsulta
        Sql = Sql & " order by descricao"
    End If

    ' abre um Recrodset da Tabela Loja
    If loja.State = 1 Then loja.Close
    loja.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaLoja.ColumnHeaders.Clear
    ListaLoja.ListItems.Clear

    If loja.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Loja(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "Loja(s) encontrado(s): " & loja.RecordCount

    ListaLoja.ColumnHeaders.Add , , "Loja", 11700

    If loja.BOF = True And loja.EOF = True Then Exit Sub
    While Not loja.EOF

        If VarType(loja("descricao")) <> vbNull Then Set itemx = ListaLoja.ListItems.Add(, , loja("descricao"))
        If VarType(loja("id_Loja")) <> vbNull Then itemx.Tag = loja("id_Loja")
        loja.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaLoja, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If loja.State = 1 Then loja.Close
    Set loja = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaLoja_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_loja.text = ListaLoja.SelectedItem.Tag

    '   If mLojaalterar = 1 Then cmdAlterar.Enabled = True Else cmdAlterar.Enabled = False
    '    If mLojaexcluir = 1 Then cmdExcluir.Enabled = True Else cmdExcluir.Enabled = False

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

    If txtConsulta.text <> "" Then
        Sqlconsulta = " Loja.descricao like '%" & txtConsulta.text & "%'"
    Else
        Sqlconsulta = " 1=1 "
    End If

    Lista (Sqlconsulta)

End Sub









