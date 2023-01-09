VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagarPagamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtDataPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Valor Pago (R$)"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pagamento"
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
         TabIndex        =   6
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.TextBox txtid_ContasPagarItem 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2085
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   30763
            MinWidth        =   30763
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
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
      Picture         =   "frmPagarPagamento.frx":0000
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
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "  Pagar"
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
      Picture         =   "frmPagarPagamento.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "frmPagarPagamento.frx":065C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5355
   End
End
Attribute VB_Name = "frmPagarPagamento"
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



Private Sub Form_Load()
    Me.Width = 5430
    Me.Height = 2835
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub
Private Sub cmdSair_Click()
    Unload Me
    Set frmPagarPagamento = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPagarPagamento = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "id_contaspagaritem = '" & txtid_ContasPagarItem.text & "'"

    If txtDataPagto.text <> "" Then campo = "DataPagto = '" & Format(txtDataPagto.text, "YYYYMMDD") & "'"
    If txtValorPago.text <> "" Then campo = campo & ", ValorPago = '" & FormatValor(txtValorPago.text, 1) & "'"

    sqlAlterar "ContaspagarItem", campo, Sqlconsulta, Me, "N"

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'----------------------------------------------------------

'--- txtDataPagto
Private Sub txtdataPagto_GotFocus()
    txtDataPagto.BackColor = &H80FFFF
End Sub
Private Sub txtdataPagto_LostFocus()
    txtDataPagto.BackColor = &H80000014
    txtDataPagto.text = SoNumero(txtDataPagto.text)
    If txtDataPagto.text <> "" Then
        txtDataPagto.text = Mid(txtDataPagto.text, 1, 2) & "/" & Mid(txtDataPagto.text, 3, 2) & "/" & Mid(txtDataPagto.text, 5, 4)
    End If
End Sub
Private Sub txtdataPagto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtValorPago.SetFocus
End Sub

'--- txtValorPago
Private Sub txtValorPago_GotFocus()
    txtValorPago.BackColor = &H80FFFF
End Sub
Private Sub txtValorPago_LostFocus()
    txtValorPago.BackColor = &H80000014
    txtValorPago.text = Format(txtValorPago.text, "###,##0.00")
End Sub
Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub


