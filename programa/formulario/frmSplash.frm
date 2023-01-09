VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   60
   End
   Begin VB.Timer Timer2 
      Interval        =   70
      Left            =   0
      Top             =   540
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7200
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   3240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   2400
         Picture         =   "frmSplash.frx":0ECA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Atenção: Este Software está protegido pela lei de direitos autorais. Não reproduzir ilegalmente."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   3660
         Width           =   6495
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "(c) LMA Inc. 2014"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H8000000E&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   2
         Top             =   2820
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Carregando . . ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   " Microsoft Windows 2000/XP/Seven"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3990
         TabIndex        =   3
         Top             =   2640
         Width           =   2865
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Sistemas de Vendas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   6885
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SqlConsulta As String

Private Sub Form_KeyPress(KeyAscii As Integer)
'Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo trata_erro
    ' Set Me.Icon = LoadPicture(ICONBD)
    lblVersion.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
    Me.MousePointer = 11    'vbHourglass

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    busca_dados_config
    MenuPrincipal.Show
End Sub

Private Sub Frame1_Click()
'Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub

Private Sub Timer2_Timer()
    If ProgressBar <= 98 Then
        ProgressBar = ProgressBar + 2
    Else
        Conectar    ' Conectar ao banco de dados Mysql CLINICA
    End If
End Sub

Private Sub busca_dados_config()
    On Error GoTo trata_erro
    '   Dim Sql As String
    '   Dim ConfigSistema As ADODB.Recordset
    ' conecta ao banco de dados

    '   Set ConfigSistema = CreateObject("ADODB.Recordset")    '''
    ' abre um Recrodset da Tabela Config Sistema
    '   Sql = " select * from ConfigSistema  where id_config is not null"

    '   If ConfigSistema.State = 1 Then ConfigSistema.Close
    '    ConfigSistema.Open Sql, banco, adOpenKeyset, adLockOptimistic
    '    If ConfigSistema.RecordCount > 0 Then
    '        If VarType(ConfigSistema("dataSistema")) <> vbNull Then DtaSistema = ConfigSistema("dataSistema")
    '    End If


    '    If ConfigSistema.State = 1 Then ConfigSistema.Close
    '    Set ConfigSistema = Nothing


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

