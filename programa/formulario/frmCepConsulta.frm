VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCepConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEP"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   8760
      TabIndex        =   14
      Top             =   5520
      Width           =   2415
      Begin VB.TextBox txtcep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   120
         Picture         =   "frmCepConsulta.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CEP encontrado"
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
         TabIndex        =   17
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.TextBox txtid_estado 
      Height          =   285
      Left            =   6360
      TabIndex        =   11
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   8415
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtlogradouro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image cmdConsultar 
         Height          =   600
         Left            =   7560
         Picture         =   "frmCepConsulta.frx":107F
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label30 
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
         TabIndex        =   16
         Top             =   0
         Width           =   8415
      End
      Begin VB.Label Label6 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Logradouro"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtid_endereco 
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3840
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8205
      Width           =   11460
      _ExtentX        =   20214
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
   Begin MSComctlLib.ListView ListaCEP 
      Height          =   5055
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8916
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
      Left            =   10200
      TabIndex        =   20
      Top             =   7440
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
      Picture         =   "frmCepConsulta.frx":1389
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
      Left            =   8160
      TabIndex        =   19
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
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
      TabIndex        =   18
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   9120
      Left            =   0
      Picture         =   "frmCepConsulta.frx":1493
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   11640
   End
End
Attribute VB_Name = "frmCepConsulta"
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
Dim SQsort As String


Private Sub cmdSair_Click()
    Unload Me
    Set frmCepConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdConsultar_Click()
'    Aguarde_Process Me, True
    Call consultarcaracteristica
    '    Aguarde_Process Me, False
End Sub

Private Sub consultarcaracteristica()
    Sqlconsulta = "1=1"
    If txtCidade.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and cep_cidade.cidade like '%" & txtCidade.text & "%'"
    End If

    If txtBairro.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and cep_bairro.bairro like '%" & txtBairro.text & "%'"
    End If

    If txtlogradouro.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and cep_endereco.logradouro like '%" & txtlogradouro.text & "%'"
    End If

    If txtEndereco.text <> "" Then
        Sqlconsulta = Sqlconsulta & " and cep_endereco.endereco like '%" & txtEndereco.text & "%'"
    End If

    If Sqlconsulta = "1=1" Then
        MsgBox ("Favor selecionar um item para consulta..."), vbInformation
        Exit Sub
    End If

    Lista (Sqlconsulta)

End Sub

Private Sub cmdTipo_Click()

End Sub

Private Sub Form_Activate()
' muda o curso para o normal


End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadPicture(ICONBD)
    Me.Width = 11550
    Me.Height = 9060
    Centerform Me
    MenuPrincipal.DesabilitaMenu

End Sub

'--------------------------- define dados da lista grid VENDA
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim CEP As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set CEP = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = "cep_bairro.bairro"

    Sql = "select cep_endereco.id_endereco, cep_endereco.id_bairro, cep_endereco.endereco_completo, cep_endereco.cep, " _
        & " cep_cidade.id_cidade, cep_cidade.id_estado, cep_cidade.cidade, " _
        & " cep_bairro.id_bairro, cep_bairro.id_cidade, cep_bairro.bairro, " _
        & " cep_estados.id_estado , cep_estados.uf " _
        & " From " _
        & " cep_endereco " _
        & " left join cep_cidade on cep_endereco.id_cidade = cep_cidade.id_cidade " _
        & " left join cep_bairro on cep_endereco.id_bairro = cep_bairro.id_bairro " _
        & " left join cep_estados on cep_cidade.id_estado = cep_estados.id_estado " _
        & " where " _
        & SQconsulta & " order by " & SQsort _
        & " limit 100"

    ' abre um Recrodset da Tabela CEP
    If CEP.State = 1 Then CEP.Close
    CEP.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaCEP.ColumnHeaders.Clear
    ListaCEP.ListItems.Clear

    If CEP.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = " CEP(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = " CEP encontrado(s): " & CEP.RecordCount

    ListaCEP.ColumnHeaders.Add , , "ID", 700
    ListaCEP.ColumnHeaders.Add , , "Endereço", 4000
    ListaCEP.ColumnHeaders.Add , , "Cidade", 2000
    ListaCEP.ColumnHeaders.Add , , "Bairro", 2000
    ListaCEP.ColumnHeaders.Add , , "Estado", 800
    ListaCEP.ColumnHeaders.Add , , "CEP", 1000
    If CEP.BOF = True And CEP.EOF = True Then Exit Sub
    While Not CEP.EOF
        If VarType(CEP("id_endereco")) <> vbNull Then Set itemx = ListaCEP.ListItems.Add(, , CEP("id_endereco"))
        If VarType(CEP("endereco_completo")) <> vbNull Then itemx.SubItems(1) = CEP("endereco_completo") Else itemx.SubItems(1) = ""
        If VarType(CEP("cidade")) <> vbNull Then itemx.SubItems(2) = CEP("cidade") Else itemx.SubItems(2) = ""
        If VarType(CEP("bairro")) <> vbNull Then itemx.SubItems(3) = CEP("bairro") Else itemx.SubItems(3) = ""
        If VarType(CEP("uf")) <> vbNull Then itemx.SubItems(4) = CEP("uf") Else itemx.SubItems(4) = ""
        If VarType(CEP("cep")) <> vbNull Then itemx.SubItems(5) = CEP("cep") Else itemx.SubItems(5) = ""
        If VarType(CEP("id_endereco")) <> vbNull Then itemx.Tag = CEP("id_endereco")
        CEP.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaCEP, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If CEP.State = 1 Then CEP.Close
    Set CEP = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCepConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub ListaCEP_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim CEP As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set CEP = CreateObject("ADODB.Recordset")

    txtid_endereco.text = ListaCEP.SelectedItem.Tag

    Sql = "select cep_endereco.id_endereco, cep_endereco.id_bairro, cep_endereco.endereco_completo, cep_endereco.cep, " _
        & " cep_cidade.id_cidade, cep_cidade.id_estado, cep_cidade.cidade, " _
        & " cep_bairro.id_bairro, cep_bairro.id_cidade, cep_bairro.bairro, " _
        & " cep_estados.id_estado , cep_estados.uf " _
        & " From " _
        & " cep_endereco " _
        & " left join cep_cidade on cep_endereco.id_cidade = cep_cidade.id_cidade " _
        & " left join cep_bairro on cep_endereco.id_bairro = cep_bairro.id_bairro " _
        & " left join cep_estados on cep_cidade.id_estado = cep_estados.id_estado " _
        & " where id_endereco = '" & txtid_endereco.text & "'"

    ' abre um Recrodset da Tabela CEP
    If CEP.State = 1 Then CEP.Close
    CEP.Open Sql, banco, adOpenKeyset, adLockOptimistic

    If CEP.RecordCount > 0 Then
        If VarType(CEP("cep")) <> vbNull Then txtCep.text = CEP("cep") Else txtCep.text = ""
        CEP.MoveNext
    End If

    If CEP.State = 1 Then CEP.Close
    Set CEP = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub txtConsultar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call consultarcaracteristica
End Sub








