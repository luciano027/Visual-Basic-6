VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6450
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   5685
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   4905
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   5685
   End
   Begin VB.TextBox txtid_login 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   150
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   4080
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
      Picture         =   "frmLogin.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdOK 
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   4080
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
      Picture         =   "frmLogin.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   0
      Left            =   0
      Picture         =   "frmLogin.frx":045C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6450
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   8880
      Index           =   1
      Left            =   -1920
      Picture         =   "frmLogin.frx":6ED9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11760
   End
End
Attribute VB_Name = "frmLogin"
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
Dim test As Boolean
Dim troq As String
Public MicroBD As String
Public TipoBD As String
Public NOLogin As String
Public mContador As Integer

Private Sub cmdCancel_Click()
    Unload Me
    Set frmLogin = Nothing
    End
End Sub

Private Sub cmdOK_Click()
    okEnter
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
    okEnter
End Sub
Private Sub okEnter()
    On Error GoTo trata_erro
    Dim login As ADODB.Recordset
    Set login = CreateObject("ADODB.Recordset")


    Sql = "SELECT * FROM Login WHERE login = '" & txtLogin.text & "' and senha ='" & txtPassword.text & "'"
    If login.State = 1 Then login.Close
    With login
        .CursorType = adOpenStatic    'Este é o unico tipo de cursor a ser usado com um cursor localizado no lado do cliente
        .CursorLocation = adUseClient    'estamos usando o cursor no cliente
        .LockType = adLockPessimistic    'Isto garente que o registros que esta sendo editado pode ser salvo
        .Source = Sql    'altere para tabela que desejar a fonte de dados usamos uma instrucal SQL
        .ActiveConnection = banco  'O recordset precisa saber qual a conexao em uso
        .Open    'abre o recordset com isto o evento MoveComplete sera disparado
    End With
    If login.RecordCount > 0 Then
        Me.Visible = False
        If VarType(login("id_login")) <> vbNull Then txtid_login.text = login("id_login")
        If VarType(login("cadastro")) <> vbNull Then mCadastro = login("cadastro")

        If VarType(login("Utilitarios")) <> vbNull Then mUtilitarios = login("Utilitarios")
        If VarType(login("AgendaTelefone")) <> vbNull Then mAgendaT = login("AgendaTelefone")
        If VarType(login("AgendaTelefoneIn")) <> vbNull Then mAgendaTIn = login("AgendaTelefoneIn")
        If VarType(login("AgendaTelefoneAl")) <> vbNull Then mAgendaTAl = login("AgendaTelefoneAl")
        If VarType(login("AgendaTelefoneEx")) <> vbNull Then mAgendaTEx = login("AgendaTelefoneEx")
        If VarType(login("AgendaTelefoneCon")) <> vbNull Then mAgendaTCo = login("AgendaTelefoneCon")

        If VarType(login("usuarios")) <> vbNull Then mUsuarios = login("usuarios")
        If VarType(login("ConsultaCEP")) <> vbNull Then mConsultaCEP = login("ConsultaCEP")
        If VarType(login("configuracao")) <> vbNull Then mConfiguracao = login("configuracao")
        If VarType(login("backup")) <> vbNull Then mbackup = login("backup")

        '  Call AbilitaMenu
        LoginM = txtLogin.text
        UsuBD = txtLogin.text
        IdLogin = txtid_login.text
        MenuPrincipal.MDIStatus.Panels.Item(3).text = UsuBD

        busca_dados_Empresa

        MenuPrincipal.AbilidataMenu

        If login.State = 1 Then login.Close
        Set login = Nothing

        Unload Me
    Else
        If login.State = 1 Then login.Close
        Set login = Nothing
        MsgBox ("Login não encontrado...")
        If mContador = 4 Then
            End
        Else
            Me.Visible = True
            mContador = mContador + 1
            txtLogin.text = ""
            txtPassword.text = ""
            txtLogin.SetFocus
        End If
    End If


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub busca_dados_Empresa()
    On Error GoTo trata_erro

    Dim Configuracao As ADODB.Recordset
    ' conecta ao banco de dados

    Set Configuracao = CreateObject("ADODB.Recordset")    '''
    ' abre um Recrodset da Tabela configuracao
    Sql = " SELECT Configuracao.*, cep_endereco.id_endereco, cep_endereco.id_bairro, cep_endereco.endereco_completo, cep_endereco.cep,"
    Sql = Sql & " cep_cidade.id_cidade, cep_cidade.id_estado, cep_cidade.cidade,"
    Sql = Sql & " cep_bairro.id_bairro, cep_bairro.id_cidade, cep_bairro.bairro,"
    Sql = Sql & " cep_estados.id_estado , cep_estados.uf"
    Sql = Sql & " From"
    Sql = Sql & " configuracao"
    Sql = Sql & " LEFT JOIN cep_endereco ON configuracao.cep = cep_endereco.cep"
    Sql = Sql & " LEFT JOIN cep_cidade ON cep_endereco.id_cidade = cep_cidade.id_cidade"
    Sql = Sql & " LEFT JOIN cep_bairro ON cep_endereco.id_bairro = cep_bairro.id_bairro"
    Sql = Sql & " LEFT JOIN cep_estados ON cep_cidade.id_estado = cep_estados.id_estado"
    Sql = Sql & " where "
    Sql = Sql & " id_config is not null"

    If Configuracao.State = 1 Then Configuracao.Close
    Configuracao.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Configuracao.RecordCount > 0 Then
        If VarType(Configuracao("empresa")) <> vbNull Then mEmpresa = Configuracao("empresa") Else mEmpresa = ""
        If VarType(Configuracao("endereco_completo")) <> vbNull Then mEndereco = Configuracao("endereco_completo") Else mEndereco = ""
        If VarType(Configuracao("bairro")) <> vbNull Then mBairro = Configuracao("bairro") Else mBairro = ""
        If VarType(Configuracao("cidade")) <> vbNull Then mCidade = Configuracao("cidade") Else mCidade = ""
        If VarType(Configuracao("numero")) <> vbNull Then mNumero = Configuracao("numero") Else mNumero = ""
        If VarType(Configuracao("cep")) <> vbNull Then mCEP = Configuracao("cep") Else mCEP = ""
        If VarType(Configuracao("uf")) <> vbNull Then mUF = Configuracao("uf") Else mUF = ""
        If VarType(Configuracao("telefone1")) <> vbNull Then mTelefone = Configuracao("telefone1") Else mTelefone = ""
        If VarType(Configuracao("logomarca")) <> vbNull Then mLogoMarcar = Configuracao("logomarca") Else mLogoMarcar = ""
        If VarType(Configuracao("id_lojacamara")) <> vbNull Then mIDLojaCamara = Configuracao("id_lojacamara") Else mIDLojaCamara = ""
        If VarType(Configuracao("id_lojaAlmoxa")) <> vbNull Then mIDLojaAlmoxa = Configuracao("id_lojaalmoxa") Else mIDLojaAlmoxa = ""

        If VarType(Configuracao("smtp_smtp")) <> vbNull Then mSmtp_smtp = Configuracao("smtp_smtp") Else mSmtp_smtp = ""
        If VarType(Configuracao("smtp_email")) <> vbNull Then mSmtp_email = Configuracao("smtp_email") Else mSmtp_email = ""
        If VarType(Configuracao("smtp_senha")) <> vbNull Then mSmtp_senha = Configuracao("smtp_senha") Else mSmtp_senha = ""
        If VarType(Configuracao("smtp_porta")) <> vbNull Then mSmtp_porta = Configuracao("smtp_porta") Else mSmtp_porta = ""

    End If

    If Configuracao.State = 1 Then Configuracao.Close
    Set Configuracao = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmLogin = Nothing
    End
End Sub

Private Sub Form_Activate()
    txtLogin.SetFocus
    mChaveInicial = "1"
    mContador = 1
End Sub

Private Sub Form_Load()
'    Set Me.Icon = LoadPicture(ICONBD)
    Me.Width = 6540
    Me.Height = 5625
    Centerform Me
    ' MenuPrincipal.DesabilitaMenu
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub txtlogin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtPassword.SetFocus
End Sub
