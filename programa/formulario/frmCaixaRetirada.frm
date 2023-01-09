VERSION 5.00
Begin VB.Form frmCaixaRetirada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa Retirada"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6570
      TabIndex        =   11
      Top             =   3270
      Width           =   6630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6255
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3240
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dados Caixa - Saida"
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
         TabIndex        =   8
         Top             =   0
         Width           =   6255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Historico"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.TextBox txtChave 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin Vendas.VistaButton cmdsair 
      Height          =   615
      Left            =   5400
      TabIndex        =   12
      Top             =   2400
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
      Picture         =   "frmCaixaRetirada.frx":0000
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
      Left            =   3840
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Retirar..."
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
      Picture         =   "frmCaixaRetirada.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   0
      Picture         =   "frmCaixaRetirada.frx":0B7C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "frmCaixaRetirada"
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


Private Sub cmdsairPG_Click()

End Sub

Private Sub Form_Activate()
    txtHistorico.SetFocus
    txtData.text = Format(Date, "DD/MM/YYYY")
End Sub

Private Sub Form_Load()
    Me.Width = 6720
    Me.Height = 4080

    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmCaixaRetirada = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCaixaRetirada = Nothing
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Dim table As String


    If txtHistorico.text <> "" Then
        campo = " historico"
        Scampo = "'" & Mid(txtHistorico.text, 1, 100) & "'"
    Else
        MsgBox ("Historico não pode ficar em branco..")
        txtHistorico.SetFocus
        Exit Sub
    End If

    If txtValor.text <> "" Then
        campo = campo & ", valorcaixadinheiro"
        Scampo = Scampo & ", '-" & FormatValor(txtValor.text, 1) & "'"
    Else
        MsgBox ("Valor não pode ficar em branco..")
        txtHistorico.SetFocus
        Exit Sub
    End If

    campo = campo & ", id_vendedor"
    Scampo = Scampo & ", ' " & txtid_vendedor.text & "'"

    campo = campo & ", datacaixa"
    Scampo = Scampo & ", '" & Format(txtData.text, "YYYYMMDD") & "'"

    campo = campo & ", status"
    Scampo = Scampo & ", 'P'"

    ' Incluir valor na tabela desc_gru
    sqlIncluir "caixa", campo, Scampo, Me, "S"

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


'---------------------------------------------------------------
'----------------------- campos do formulario-------------------------------

'------------ nome
Private Sub txthistorico_GotFocus()
    txtHistorico.BackColor = &H80FFFF
End Sub
Private Sub txthistorico_LostFocus()
    txtHistorico.BackColor = &H80000014
    If Len(txtHistorico.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtHistorico.text)
        txtHistorico.SetFocus
    End If
End Sub
Private Sub txthistorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtValor.SetFocus
End Sub

'''--- txtdata
'Private Sub txtdata_GotFocus()
'    txtData.BackColor = &H80FFFF
'End Sub
'Private Sub txtdata_LostFocus()
'    txtData.BackColor = &H80000014
'    If txtData.text <> "" Then
'        txtData.text = SoNumero(txtData.text)
'        txtData.text = Mid(txtData.text, 1, 2) & "/" & Mid(txtData.text, 3, 2) & "/" & Mid(txtData.text, 5, 4)
'        If Not IsDate(txtData.text) Then
'            MsgBox ("Data Invalida...")
'            txtData.SetFocus
'        End If
'    End If
'End Sub
'Private Sub txtdata_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then txtValor.SetFocus
'End Sub

'--- txtValor
Private Sub txtValor_GotFocus()
    txtValor.BackColor = &H80FFFF
End Sub
Private Sub txtValor_LostFocus()
    txtValor.BackColor = &H80000014
    txtValor.text = Format(txtValor.text, "###,##0.00")
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub













