VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaEstoque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Estoque"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   8535
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   7935
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   0
         Width           =   8535
      End
      Begin VB.Image cmdConsultar 
         Height          =   360
         Left            =   8040
         Picture         =   "frmConsultaEstoque.frx":0000
         Stretch         =   -1  'True
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame Cliente 
      Caption         =   "Empresa"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   7200
      Width           =   10935
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto selecionado"
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
         Width           =   10935
      End
      Begin VB.Label lblNome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4200
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   285
      Left            =   5520
      TabIndex        =   0
      Top             =   9000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   9300
      Width           =   11370
      _ExtentX        =   20055
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
   Begin MSComctlLib.ListView ListaEstoque 
      Height          =   6615
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
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
      Left            =   10080
      TabIndex        =   12
      Top             =   8280
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
      Picture         =   "frmConsultaEstoque.frx":030A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCopiar 
      Height          =   615
      Left            =   8880
      TabIndex        =   13
      Top             =   8280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "OK"
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
      Picture         =   "frmConsultaEstoque.frx":0414
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado da consulta"
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
      TabIndex        =   11
      Top             =   240
      Width           =   7935
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
      Left            =   8160
      TabIndex        =   10
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   11160
      Left            =   0
      Picture         =   "frmConsultaEstoque.frx":0766
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   11520
   End
End
Attribute VB_Name = "frmConsultaEstoque"
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
Dim strChave As String
Dim Sql As String
Dim SQsort As String


Private Sub cmdCopiar_Click()

    IDEstoque = txtid_estoque.text

    Unload Me
    Set frmConsultaEstoque = Nothing

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmConsultaEstoque = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarDescricao_Estoque
    Aguarde_Process Me, False
End Sub



Private Sub Form_Activate()
' muda o curso para o normal

    Me.Width = 11400
    Me.Height = 9795

    Centerform Me

    MenuPrincipal.DesabilitaMenu
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadPicture(ICONBD)
    '    cmdConsultar_Click
    strChave = "N"
End Sub

'--------------------------- define dados da lista grid VENDA
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Estoque = CreateObject("ADODB.Recordset")

    Sql = "SELECT Estoques.codigo_est,  estoques.descricao, estoquesaldo.saldo, estoques.unidade , estoques.preco_venda, "
    Sql = Sql & " estoques.id_estoque"
    Sql = Sql & " FROM Estoques "
    Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
    Sql = Sql & " Where " & SQconsulta
    Sql = Sql & " order by descricao"
    Sql = Sql & " limit 50"

    ' abre um Recrodset da Tabela Estoque
    If Estoque.State = 1 Then Estoque.Close
    Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaEstoque.ColumnHeaders.Clear
    ListaEstoque.ListItems.Clear

    If Estoque.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Estoque(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "Estoque(s) encontrado(s): " & Estoque.RecordCount

    ListaEstoque.ColumnHeaders.Add , , "Código", 1100
    ListaEstoque.ColumnHeaders.Add , , "Estoque", 5400
    ListaEstoque.ColumnHeaders.Add , , "Unidade", 1000, lvwColumnCenter
    ListaEstoque.ColumnHeaders.Add , , "Saldo", 1500, lvwColumnRight
    ListaEstoque.ColumnHeaders.Add , , "Preço", 1500, lvwColumnRight

    If Estoque.BOF = True And Estoque.EOF = True Then Exit Sub
    While Not Estoque.EOF
        If VarType(Estoque("codigo_est")) <> vbNull Then Set itemx = ListaEstoque.ListItems.Add(, , Estoque("codigo_est"))
        If VarType(Estoque("descricao")) <> vbNull Then itemx.SubItems(1) = Estoque("descricao") Else itemx.SubItems(1) = ""
        If VarType(Estoque("unidade")) <> vbNull Then itemx.SubItems(2) = Estoque("unidade") Else itemx.SubItems(2) = ""
        If VarType(Estoque("saldo")) <> vbNull Then itemx.SubItems(3) = Format(Estoque("saldo"), "###,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Estoque("preco_venda")) <> vbNull Then itemx.SubItems(4) = Format(Estoque("preco_venda"), "###,##0.00") Else itemx.SubItems(4) = ""
        If VarType(Estoque("id_Estoque")) <> vbNull Then itemx.Tag = Estoque("id_Estoque")
        Estoque.MoveNext

    Wend

    'Zebra o listview
    If LVZebra(ListaEstoque, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If strChave = "S" Then ListaEstoque.SetFocus

    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConsultaEstoque = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub ListaEstoque_KeyPress(KeyAscii As Integer)
    On Error GoTo trata_erro

    If KeyAscii = vbKeyReturn Then

        IDEstoque = txtid_estoque.text

        Unload Me
        Set frmConsultaEstoque = Nothing

    End If
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub ListaEstoque_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    Set Estoque = CreateObject("ADODB.Recordset")

    txtid_estoque.text = ListaEstoque.SelectedItem.Tag

    If txtid_estoque.text <> "" Then
        Sql = " SELECT id_Estoque, descricao FROM Estoques where id_Estoque = '" & txtid_estoque.text & "'"

        If Estoque.State = 1 Then Estoque.Close
        Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Estoque.RecordCount > 0 Then
            If VarType(Estoque("descricao")) <> vbNull Then lblNome.Caption = Estoque("descricao") Else lblNome.Caption = ""
        End If
    End If

    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call consultarDescricao_Estoque
End Sub

Private Sub consultarDescricao_Estoque()

    Sqlconsulta = " estoques.status = 'A'"

    If txtConsulta.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and estoques.descricao like '%" & txtConsulta.text & "%'"
        Sqlconsulta = Sqlconsulta & " or estoques.codigo_est like '%" & txtConsulta.text & "%'"
        strChave = "S"
    Else
        strChave = "N"
    End If

    Lista (Sqlconsulta)


End Sub













