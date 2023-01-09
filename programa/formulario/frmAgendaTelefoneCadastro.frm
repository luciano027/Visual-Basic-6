VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgendaTelefoneCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de Telefone"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   5415
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox txtAtividade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtTelefone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txttelefone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtcelular 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox txtObs 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3413
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmAgendaTelefoneCadastro.frx":0000
      End
      Begin VB.Image Image2 
         Height          =   1935
         Left            =   2880
         Picture         =   "frmAgendaTelefoneCadastro.frx":0082
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Registro na agenda"
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
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Observação"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.TextBox txtid_telefone 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14888
            MinWidth        =   14888
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   4440
      TabIndex        =   19
      Top             =   5520
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
      Picture         =   "frmAgendaTelefoneCadastro.frx":1A21
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
      Left            =   3240
      TabIndex        =   20
      Top             =   5520
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
      Picture         =   "frmAgendaTelefoneCadastro.frx":1B2B
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdGravar 
      Height          =   615
      Left            =   2040
      TabIndex        =   21
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Gravar"
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
      Picture         =   "frmAgendaTelefoneCadastro.frx":207D
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   6600
      Left            =   0
      Picture         =   "frmAgendaTelefoneCadastro.frx":25CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "frmAgendaTelefoneCadastro"
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

Private Sub cmdAlterar_Click()

End Sub

Private Sub Form_Load()

    Set Me.Icon = LoadPicture(ICONBD)
    chave = 1
End Sub

Private Sub Form_Activate()
    On Error GoTo trata_erro

    Me.Width = 5700
    Me.Height = 7140
    Centerform Me

    If txtTipo.text = "A" Then
        If chave = 1 Then atualizatela
    End If

    If txtTipo.text = "E" Then
        If chave = 1 Then atualizatela
    End If

    If txtTipo.text = "I" Then cmdGravar.Enabled = True

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
Private Sub atualizatela()
    On Error GoTo trata_erro
    Dim telefone As ADODB.Recordset
    ' conecta ao banco de dados
    Sql = "SELECT * FROM telefone WHERE id_telefone = '" & txtid_telefone.text & "'"
    Set telefone = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela telefone
    If telefone.State = 1 Then telefone.Close
    telefone.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If telefone.RecordCount > 0 Then
        If VarType(telefone("nome")) <> vbNull Then txtNome.text = telefone("nome") Else txtNome.text = ""
        If VarType(telefone("telefone")) <> vbNull Then txtTelefone.text = telefone("telefone") Else txtTelefone.text = ""
        If VarType(telefone("telefone2")) <> vbNull Then txttelefone2.text = telefone("telefone2") Else txttelefone2.text = ""
        If VarType(telefone("celular")) <> vbNull Then txtcelular.text = telefone("celular") Else txtcelular.text = ""
        If VarType(telefone("email")) <> vbNull Then txtemail.text = telefone("email") Else txtemail.text = ""
        If VarType(telefone("atividade")) <> vbNull Then txtAtividade.text = telefone("atividade") Else txtAtividade.text = ""
        If VarType(telefone("id_telefone")) <> vbNull Then txtid_telefone.text = telefone("id_telefone") Else txtid_telefone.text = ""
        If VarType(telefone("obs")) <> vbNull Then txtObs.text = telefone("obs") Else txtObs.text = ""
        GravaLog ("telefone Telefone. Usuario:" & UsuBD & "Consulta:" & txtNome.text)
        chave = 2
        txtTipo.text = "A"

        '      If mAgendaTEx = 1 Then cmdExcluir.Enabled = True Else cmdExcluir.Enabled = False
        '      If mAgendaTAl = 1 Then cmdGravar.Enabled = True Else cmdGravar.Enabled = False
    Else
        txtTipo.text = "I"
        '       If mAgendaTIn = 1 Then cmdGravar.Enabled = True Else cmdGravar.Enabled = False
    End If
    If telefone.State = 1 Then telefone.Close
    Set telefone = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro

    If txtid_telefone.text <> "" Then
        confirma = MsgBox("Confirma Exclusão", vbQuestion + vbYesNo)
        If confirma = vbYes Then

            Sqlconsulta = "id_telefone = '" & txtid_telefone.text & "'"

            sqlDeletar "telefone", Sqlconsulta, Me, "S"

            Unload Me

        End If
    End If

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Dim table As String


    ' Rotina de gravacao de inclusao dos dados
    If txtTipo.text = "I" Then

        If txtNome.text <> "" Then
            campo = "nome"
            Scampo = "'" & txtNome.text & "'"
        Else
            MsgBox ("Data não pode ficar em branco")
            txtNome.text = ""
            txtNome.SetFocus
        End If

        If txtTelefone.text <> "" Then
            campo = campo & ", telefone"
            Scampo = Scampo & ", '" & txtTelefone.text & "'"
        End If

        If txttelefone2.text <> "" Then
            campo = campo & ", telefone2"
            Scampo = Scampo & ", '" & txttelefone2.text & "'"
        End If

        If txtcelular.text <> "" Then
            campo = campo & ", celular"
            Scampo = Scampo & ", '" & txtcelular.text & "'"
        End If

        If txtemail.text <> "" Then
            campo = campo & ", email"
            Scampo = Scampo & ", '" & txtemail.text & "'"
        End If

        If txtAtividade.text <> "" Then
            campo = campo & ", atividade"
            Scampo = Scampo & ", '" & txtAtividade.text & "'"
        End If

        If txtObs.text <> "" Then
            campo = campo & ", obs"
            Scampo = Scampo & ", '" & txtObs.text & "'"
        End If

        sqlIncluir "telefone", campo, Scampo, Me, "S"

        txtTipo.text = ""

    End If
    If txtTipo.text = "A" Then

        Sqlconsulta = "id_telefone = '" & txtid_telefone.text & "'"

        campo = "nome = '" & txtNome.text & "'"

        If txtTelefone.text <> "" Then campo = campo & ", telefone ='" & txtTelefone.text & "'"
        If txttelefone2.text <> "" Then campo = campo & ", telefone2 ='" & txttelefone2.text & "'"
        If txtcelular.text <> "" Then campo = campo & ", celular ='" & txtcelular.text & "'"
        If txtemail.text <> "" Then campo = campo & ", email = '" & txtemail.text & "'"
        If txtAtividade.text <> "" Then campo = campo & ", atividade = '" & txtAtividade.text & "'"
        If txtObs.text <> "" Then campo = campo & ", obs = '" & txtObs.text & "'"

        ' Aletar dos dados da tabela tabela_cid10
        sqlAlterar "telefone", campo, Sqlconsulta, Me, "S"

        txtTipo.text = ""

    End If

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmAgendaTelefoneCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAgendaTelefoneCadastro = Nothing
End Sub


'----------------------- campos do formulario-------------------------------
'--- txtnome
Private Sub txtnome_GotFocus()
    txtNome.BackColor = &H80FFFF
End Sub
Private Sub txtnome_LostFocus()
    txtNome.BackColor = &H80000014
    If Len(txtNome.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtNome.text)
        txtNome.SetFocus
    End If
    If txtNome.text = "" Then
        MsgBox ("Valor não pode ficar em branco..."), vbExclamation
        txtNome.SetFocus
    End If
End Sub
Private Sub txtnome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtTelefone.SetFocus
End Sub

'--- txtTelefone
Private Sub txtTelefone_GotFocus()
    txtTelefone.BackColor = &H80FFFF
End Sub
Private Sub txtTelefone_LostFocus()
    txtTelefone.BackColor = &H80000014
    If txtTelefone.text <> "" Then
        txtTelefone.text = SoNumero(txtTelefone.text)
        txtTelefone.text = FormataTelefone(txtTelefone.text)
    End If
End Sub
Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txttelefone2.SetFocus
End Sub

'--- txtTelefone
Private Sub txtTelefone2_GotFocus()
    txttelefone2.BackColor = &H80FFFF
End Sub
Private Sub txtTelefone2_LostFocus()
    txttelefone2.BackColor = &H80000014
    If txttelefone2.text <> "" Then
        txttelefone2.text = SoNumero(txttelefone2.text)
        txttelefone2.text = FormataTelefone(txttelefone2.text)
    End If
End Sub
Private Sub txtTelefone2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtcelular.SetFocus
End Sub

'--- txtcelular
Private Sub txtcelular_GotFocus()
    txtcelular.BackColor = &H80FFFF
End Sub
Private Sub txtcelular_LostFocus()
    txtcelular.BackColor = &H80000014
    If txtcelular.text <> "" Then
        txtcelular.text = SoNumero(txtcelular.text)
        txtcelular.text = FormataTelefone(txtcelular.text)
    End If
End Sub
Private Sub txtcelular_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtemail.SetFocus
End Sub

'--- txtEmail
Private Sub txtemail_GotFocus()
    txtemail.BackColor = &H80FFFF
End Sub
Private Sub txtemail_LostFocus()
    txtemail.BackColor = &H80000014
    If Len(txtemail.text) > 35 Then
        MsgBox "Comprimento do campo e de 35 digitos, voce digitou " & Len(txtemail.text)
        txtemail.SetFocus
    End If
End Sub
Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtAtividade.SetFocus
End Sub

'--- txtAtividade
Private Sub txtAtividade_GotFocus()
    txtAtividade.BackColor = &H80FFFF
End Sub
Private Sub txtAtividade_LostFocus()
    txtAtividade.BackColor = &H80000014
    If Len(txtAtividade.text) > 35 Then
        MsgBox "Comprimento do campo e de 35 digitos, voce digitou " & Len(txtAtividade.text)
        txtAtividade.SetFocus
    End If
End Sub
Private Sub txtAtividade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtObs.SetFocus
End Sub

