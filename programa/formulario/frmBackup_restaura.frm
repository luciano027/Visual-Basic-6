VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup_restaura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backu / Restaura"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmBackup_restaura.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Local Backup"
      Height          =   4380
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   3705
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3255
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Realizar Backup "
      Height          =   3615
      Left            =   3720
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Status do Processo : "
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Total de Tabelas :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Tabela Atual :"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Tabela :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Registros :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblTotalTabelas 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblTabelaAtual 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblNomeTabela 
         Caption         =   "Aguardando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblQtdeRegistros 
         Caption         =   "Aguardando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações do Servidor"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "2562"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "root"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtBanco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Text            =   "Botelho"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtServidor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "localhost"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Usuário"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Senha"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Servidor"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Banco"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".sql"
      DialogTitle     =   "Localizar Arquivo"
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   5970
      Width           =   7755
      _ExtentX        =   13679
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
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   3960
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtNovoBanco 
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optNovoBanco 
         Caption         =   "Novo Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbBanco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optOutroBanco 
         Caption         =   "Outro Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   500
         Width           =   1215
      End
      Begin VB.OptionButton optOriginal 
         Caption         =   "Banco Original"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   6600
      TabIndex        =   31
      Top             =   5280
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
      Picture         =   "frmBackup_restaura.frx":1ABD
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton btnBackup 
      Height          =   615
      Left            =   5400
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Iniciar"
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
      Picture         =   "frmBackup_restaura.frx":1BC7
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton btnParar 
      Height          =   615
      Left            =   4200
      TabIndex        =   33
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Parar"
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
      Picture         =   "frmBackup_restaura.frx":1F19
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
End
Attribute VB_Name = "frmBackup_restaura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim cn As ADODB.Connection
Dim mLocal As String
Dim strArquivo As String
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&


Private Sub RemoveMenus()
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Private Sub btnBackup_Click()
    Dim arqbackup As String
    Dim sBanco As String
    Dim mArquivo As String


    If mLocal = "" Then
        MsgBox ("Favor selecionar um local para o Backup..."), vbInformation
        Exit Sub
    End If

    mArquivo = mLocal & "\Backup_" & Format(Date, "dd_MM_yy") & "_" & Format(Time, "hh_mm_ss") & ".sql"

    If Len(txtUser) = 0 Or Len(txtBanco) = 0 Or Len(txtSenha) = 0 Or Len(txtServidor) = 0 Then
        MsgBox "Verifique as informações do servidor", vbCritical
        Exit Sub
    End If

    sBanco = txtBanco    'GetSetting("BsControl", "Servidor", "Shema", Valor)
    Call conex(bd, sBanco)
    Call MySQLBackup(mArquivo, bd, lstLog)

    ' GravaLog ("Backup. Usuario:" & Principal.UsuBD)
End Sub
Function conex(cnn As ADODB.Connection, Optional banco As String) As Boolean
    On Error Resume Next
    Set bd = Nothing
    bd.Close
    Err.Clear
    Set cnn = New ADODB.Connection
    With cnn
        If Len(banco) = 0 Then
            .ConnectionTimeout = 60
            .CommandTimeout = 400
            .CursorLocation = adUseClient
            .Open "Driver={MySQL ODBC 5.1 Driver};" & _
                  "user=" & txtUser & _
                  ";password=" & txtSenha & _
                  ";server=" & txtServidor & _
                  ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)



            If .State = 1 Then
                conex = True
            Else
                conex = False
                MsgBox "Não foi possivel se registrar no sistema!!!", vbCritical
                End
            End If
        Else
            .ConnectionTimeout = 60
            .CommandTimeout = 400
            .CursorLocation = adUseClient
            .Open "Driver={MySQL ODBC 5.1 Driver};" & _
                  "user=" & txtUser & _
                  ";password=" & txtSenha & _
                  ";database=" & txtBanco & _
                  ";server=" & txtServidor & _
                  ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            If .State = 1 Then
                conex = True
            Else
                conex = False
                MsgBox "Não foi possivel se registrar no sistema!!!", vbCritical
                End
            End If
        End If
    End With
End Function

Private Sub btnParar_Click()
    sStop = True
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmBackup_restaura = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub Form_Activate()
    Set Me.Icon = LoadPicture(ICONBD)

    Me.Width = 7845
    Me.Height = 6825

    Centerform Me

    MenuPrincipal.DesabilitaMenu
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmBackup_restaura = Nothing
    MenuPrincipal.AbilidataMenu
End Sub



Private Sub optNovoBanco_Click()
    If Len(txtSenha) = 0 Or Len(txtUser) = 0 Or Len(txtServidor) = 0 Or Len(txtBanco) = 0 Then
        MsgBox "É necessário preencher os dados no painel ao lado", vbCritical
        txtSenha.SetFocus
        optNovoBanco.Value = False
        Exit Sub
    End If
    cmbBanco.Visible = False
    txtNovoBanco.Visible = True
    txtNovoBanco.SetFocus
End Sub
Private Sub optOriginal_Click()
    If Len(txtSenha) = 0 Or Len(txtUser) = 0 Or Len(txtServidor) = 0 Or Len(txtBanco) = 0 Then
        MsgBox "É necessário preencher os dados no painel ao lado", vbCritical
        txtSenha.SetFocus
        optOriginal.Value = False
        Exit Sub
    End If
    cmbBanco.Visible = False
    txtNovoBanco.Visible = False
End Sub

Private Sub optOutroBanco_Click()
    If Len(txtSenha) = 0 Or Len(txtUser) = 0 Or Len(txtServidor) = 0 Or Len(txtBanco) = 0 Then
        MsgBox "É necessário preencher os dados no painel ao lado", vbCritical
        txtSenha.SetFocus
        optOutroBanco.Value = False
        Exit Sub
    End If
    txtNovoBanco.Visible = False
    cmbBanco.Visible = True
    cmbBanco.Enabled = True
    Sql = "SHOW DATABASES"
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Call conex(bd)
    rs.Open Sql, bd, 3, 3
    cmbBanco.Clear
    Do While Not rs.EOF
        cmbBanco.AddItem rs!Database
        rs.MoveNext
    Loop
    cmbBanco.ListIndex = 1
    cmbBanco.SetFocus
End Sub

Private Sub Dir1_Change()
    mLocal = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

