VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoqueConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estoque"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13125
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
      Left            =   11520
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   11160
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_Estoque 
      Height          =   285
      Left            =   10800
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
      Width           =   6255
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5655
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
         Width           =   6255
      End
      Begin VB.Image cmdConsultar 
         Height          =   315
         Left            =   5760
         Picture         =   "frmEstoqueConsulta.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView ListaEstoques 
      Height          =   6495
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   12615
      _ExtentX        =   22251
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
      Top             =   8130
      Width           =   13125
      _ExtentX        =   23151
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
      Left            =   11760
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
      Picture         =   "frmEstoqueConsulta.frx":030A
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
      Left            =   10560
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
      Picture         =   "frmEstoqueConsulta.frx":0414
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
      Left            =   9360
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
      Picture         =   "frmEstoqueConsulta.frx":0966
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
      Left            =   8160
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
      Picture         =   "frmEstoqueConsulta.frx":0CB8
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
      Width           =   9495
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
      Left            =   9720
      TabIndex        =   17
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   9240
      Left            =   0
      Picture         =   "frmEstoqueConsulta.frx":100A
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   13200
   End
End
Attribute VB_Name = "frmEstoqueConsulta"
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
    Me.Width = 13215
    Me.Height = 8880
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmEstoqueConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEstoqueConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAlterar_Click()
    With frmEstoqueCadastro
        .txtid_estoque.text = txtid_estoque.text
        .txtTipo.text = "A"
        .Show 1
    End With
    cmdConsultar_Click
    ' Lista ("")
End Sub

Private Sub cmdExcluir_Click()
    With frmEstoqueCadastro
        .txtid_estoque.text = txtid_estoque.text
        .txtTipo.text = "E"
        .Show 1
    End With
    cmdConsultar_Click
    'Lista ("")
End Sub

Private Sub cmdIncluir_Click()
    With frmEstoqueCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
    cmdConsultar_Click
    'Lista ("")
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Estoques As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Estoques = CreateObject("ADODB.Recordset")

    If SQconsulta = "" Then
        Sql = "SELECT Estoques.*, estoquesaldo.saldo "
        Sql = Sql & " FROM Estoques "
        Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " order by descricao"
    Else
        Sql = "SELECT Estoques.*, estoquesaldo.saldo "
        Sql = Sql & " FROM Estoques "
        Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " Where " & SQconsulta
        Sql = Sql & " order by descricao"
    End If

    ' abre um Recrodset da Tabela Estoques
    If Estoques.State = 1 Then Estoques.Close
    Estoques.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaEstoques.ColumnHeaders.Clear
    ListaEstoques.ListItems.Clear

    If Estoques.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Estoques(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "descricao(s) encontrado(s): " & Estoques.RecordCount

    ListaEstoques.ColumnHeaders.Add , , "Descrição", 8350
    ListaEstoques.ColumnHeaders.Add , , "Unidade", 1000, lvwColumnCenter
    ListaEstoques.ColumnHeaders.Add , , "Saldo", 1500, lvwColumnRight
    ListaEstoques.ColumnHeaders.Add , , "Preco venda", 1500, lvwColumnRight

    If Estoques.BOF = True And Estoques.EOF = True Then Exit Sub
    While Not Estoques.EOF

        If VarType(Estoques("descricao")) <> vbNull Then Set itemx = ListaEstoques.ListItems.Add(, , Estoques("descricao"))
        If VarType(Estoques("unidade")) <> vbNull Then itemx.SubItems(1) = Estoques("unidade") Else itemx.SubItems(1) = ""
        If VarType(Estoques("saldo")) <> vbNull Then itemx.SubItems(2) = Format(Estoques("saldo"), "###,##0.000") Else itemx.SubItems(2) = ""
        If VarType(Estoques("preco_venda")) <> vbNull Then itemx.SubItems(3) = Format(Estoques("preco_venda"), "###,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Estoques("id_estoque")) <> vbNull Then itemx.Tag = Estoques("id_estoque")
        Estoques.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaEstoques, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Estoques.State = 1 Then Estoques.Close
    Set Estoques = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaEstoques_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_estoque.text = ListaEstoques.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarNome_descricao
    Aguarde_Process Me, False
End Sub

Private Sub consultarNome_descricao()

    If optAtivo.Value = True Then Sqlconsulta = " status = 'A'"
    If optInativo.Value = True Then Sqlconsulta = " status = 'I'"
    If optTodos.Value = True Then Sqlconsulta = " 1=1 "

    If txtConsulta.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and Estoques.descricao like '%" & txtConsulta.text & "%'"
        Sqlconsulta = Sqlconsulta & " or estoques.codigo_est like '%" & txtConsulta.text & "%'"
    End If

    Lista (Sqlconsulta)

End Sub

Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub

