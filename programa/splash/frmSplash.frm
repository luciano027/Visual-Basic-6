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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7200
      Begin VB.TextBox txtLog 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   3240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Sistema de Vendas"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   6885
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   2400
         Picture         =   "frmSplash.frx":030A
         Stretch         =   -1  'True
         Top             =   360
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
      Begin VB.Label lblMsg 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   3000
         Width           =   3975
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
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlconsulta As String
Private PodeFechar As Boolean
Private LocalFile As String
Private NetFile As String
Private LocalDateFile As String * 17
Private NetDateFile As String * 17
Private NomeArquivoBackup As String
Private Arqupdate As String


Private Sub Form_KeyPress(KeyAscii As Integer)
'Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo trata_erro
    lblVersion.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
    Me.MousePointer = 11    'vbHourglass

    Arqupdate = ReadINI("Arquivos", "ArqUpDate", App.Path & "\vendas.ini")
    LocalFile = App.Path & "\vendas.exe"
    NetFile = Arqupdate

    'Inicia o processo de verificação entre datas de arquivos
    VerificaSeExisteAtualizacao

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PodeFechar = True Then
        Unload Me
    Else
        Cancel = 1
    End If


End Sub


Private Sub FinalizaAtualizador()

    Me.MousePointer = 0
    PodeFechar = True
    Set frmSplash = Nothing
    Unload Me
    End

End Sub


Private Function VerificaSeExisteAtualizacao() As Boolean

       CopiaArquivos
       
       AbrirPrograma

       FinalizaAtualizador
       
       
End Function

Private Sub AtualizaPrograma()

    ProgressBar.Value = 20
    lblMsg.Caption = "Iniciando atualização..."
    Escreve_Log ("Iniciando atualização...")
    lblMsg.Refresh
    Sleep (2000)

    ProgressBar.Value = 30
    lblMsg.Caption = "Fechando aplicação..."
    Escreve_Log ("Fechando aplicação " & LocalFile)
    lblMsg.Refresh
    Sleep (2000)
    FechaPrograma

    ProgressBar.Value = 50
    lblMsg.Caption = "Efetuando backup dos arquivos existentes..."
    CriaBackup
    Escreve_Log ("Efetuando backup dos arquivos existentes..." & vbCrLf & _
                 String(21, ".") & "Arquivo de backup: " & NomeArquivoBackup)
    lblMsg.Refresh
    Sleep (2000)

    ProgressBar.Value = 60
    lblMsg.Caption = "Copiando os novos arquivos..."
    Escreve_Log ("Copiando os novos arquivos... " & _
                 vbCrLf & String(21, ".") & "Arquivo local: " & LocalFile & " - ok!" & _
                 vbCrLf & String(21, ".") & "Arquivo atualização: " & NetFile) & " - ok! "
    lblMsg.Refresh
    Sleep (2000)
    CopiaArquivos

    ProgressBar.Value = 85
    lblMsg.Caption = "Finalizando atualização e iniciando o programa..."
    Escreve_Log ("Finalizando atualização e iniciando o programa...")
    Escreve_Log ("Instalação finalizada com sucesso!")
    lblMsg.Refresh
    Sleep (2000)

    ProgressBar.Value = 100
    lblMsg.Caption = "Gerando Log do processo..."
    lblMsg.Refresh
    Sleep (1000)
    GeraLog ("Log.txt")

    AbrirPrograma

End Sub

Private Sub GeraLog(ByVal strFile As String)
    Dim FSO As New FileSystemObject
    Dim iARQ As TextStream

    On Error GoTo Erro_GeraLog

    Set FSO = New FileSystemObject

    If Dir$(App.Path & "\" & "Log.txt") <> vbNullString Then
        Set iARQ = FSO.OpenTextFile(App.Path & "\" & strFile, ForAppending)
    Else
        Set iARQ = FSO.CreateTextFile(App.Path & "\" & strFile, False)
    End If

    iARQ.WriteLine txtLog.Text

    Set FSO = Nothing
    Set iARQ = Nothing

    Exit Sub

Erro_GeraLog:
    MsgBox "Erro ao gerar o Log!", vbCritical, "ERRO!"
    FinalizaAtualizador
End Sub

Private Sub Escreve_Log(strLog As String)

    If txtLog.Text = vbNullString Then
        txtLog.Text = vbCrLf & String(37, "=") & vbCrLf
        txtLog.Text = txtLog.Text & "* LOG GERADO EM " & Now() & " *" & vbCrLf
        txtLog.Text = txtLog.Text & String(37, "=") & vbCrLf & vbCrLf
        txtLog.Text = txtLog.Text & "[" & Now() & "] " & strLog & vbCrLf
    Else
        txtLog.Text = txtLog.Text & "[" & Now() & "] " & strLog & vbCrLf
    End If

End Sub


Private Sub AbrirPrograma()

    ShellExecute Me.hWnd, "OPEN", LocalFile, "", "", SW_SHOW

End Sub

Private Sub CriaBackup()
    Dim ExtensaoArquivo As String

    On Error GoTo Erro_CriaBackup

    'Define nome do Backup
    NomeArquivoBackup = Mid(LocalFile, 1, Len(LocalFile) - 4)
    ExtensaoArquivo = Right(LocalFile, 4)
    NomeArquivoBackup = NomeArquivoBackup & AdicionaZero(Day(Now())) & _
                        AdicionaZero(Month(Now())) & Right(Year(Now()), 2) & "_" & _
                        AdicionaZero(Hour(Now())) & AdicionaZero(Minute(Now())) & _
                        ExtensaoArquivo

    'Gera o Backup
    FileCopy LocalFile, NomeArquivoBackup

    'Deleta o arquivo atual
    Kill LocalFile

    Exit Sub

Erro_CriaBackup:
    MsgBox "Erro ao criar o backup!", vbCritical, "ERRO!"
    Escreve_Log ("Erro ao criar o backup!")
    Escreve_Log (Err.Number & " - " & Err.Description)
    GeraLog ("Log.txt")
    FinalizaAtualizador
End Sub

Private Sub CopiaArquivos()
    On Error GoTo Erro_CopiaArquivo

    'Copia o novo arquivo
    FileCopy NetFile, LocalFile

    Exit Sub

Erro_CopiaArquivo:
    MsgBox "Erro ao copiar o arquivo!", vbCritical, "ERRO!"
    Escreve_Log ("Erro ao copiar o arquivo!")
    Escreve_Log (Err.Number & " - " & Err.Description)
    GeraLog ("Log.txt")
    FinalizaAtualizador
End Sub

Private Function AdicionaZero(strNumero As Integer) As String

    If strNumero < 10 Then
        AdicionaZero = "0" & strNumero
    Else
        AdicionaZero = strNumero
    End If

End Function
Private Sub FechaPrograma()
    Dim WinHndl As Long

    On Error GoTo Erro_Handle

    'Obtém o Handle da janela que está sendo informada abaixo
    WinHndl = FindWindow(vbNullString, "Arquivo.txt - Bloco de Notas")

    If WinHndl > 0 Then
        PostMessage WinHndl, WM_CLOSE, 0&, 0&
    End If

    Exit Sub

Erro_Handle:
    MsgBox "Impossível definir o Handle!", vbCritical, "ERRO!"
    Escreve_Log ("Impossível definir o Handle!")
    Escreve_Log (Err.Number & " - " & Err.Description)
    GeraLog ("Log.txt")
    FinalizaAtualizador
End Sub

Private Sub DefineDataArquivos()
    On Error GoTo Erro_DefineData

    LocalDateFile = CDate(FileDateTime(LocalFile))
    NetDateFile = CDate(FileDateTime(NetFile))

    Exit Sub

Erro_DefineData:
    MsgBox "Erro em definir as datas!", vbCritical, "ERRO!"
    Escreve_Log ("Erro em definir as datas!")
    Escreve_Log (Err.Number & " - " & Err.Description)
    GeraLog ("Log.txt")
    FinalizaAtualizador
End Sub


