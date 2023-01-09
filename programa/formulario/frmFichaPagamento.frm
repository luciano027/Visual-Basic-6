VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaPagamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Pagamento"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTotalPagar 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2175
      Begin VB.TextBox txtCaixaDinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   120
         TabIndex        =   15
         Top             =   345
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estoque"
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
         Left            =   -1560
         TabIndex        =   17
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinheiro (R$)"
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
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   2055
      Begin VB.TextBox txtCaixaCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão (R$)"
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
         TabIndex        =   13
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estoque"
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
         Left            =   -1560
         TabIndex        =   12
         Top             =   -1560
         Width           =   7695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estoque"
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
         Left            =   -1560
         TabIndex        =   9
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total a pagar (R$)"
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
         Width           =   4335
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.TextBox txtid_Prazo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCliente 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      Picture         =   "frmFichaPagamento.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComctlLib.StatusBar StatusBarVendedor 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   4680
      _ExtentX        =   8255
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
   Begin Vendas.VistaButton cmdGravar 
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "Pagar"
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
      Picture         =   "frmFichaPagamento.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "frmFichaPagamento.frx":065C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmFichaPagamento"
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
Dim mID_prazo As String
Dim mDinheiro As Double
Dim mCartao As Double

Private Sub Form_Activate()
    txtCaixaDinheiro.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 4770
    Me.Height = 3990
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaPagamento = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFichaPagamento = Nothing
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro
    Dim mPagar As Double
    Dim mPagamento As Double

    mPagar = txtTotalPagar.text
    mPagamento = lblTotalPagar.Caption

    If mPagar > mPagamento Then
        MsgBox ("Valor do pagamento menor que o valor do Debito R$:") & Format(txtTotalPagar.text, "###,##0.00"), vbInformation
        Exit Sub
    End If

    If mPagar < mPagamento Then
        MsgBox ("Valor do pagamento maior que o valor do Debito R$:") & Format(txtTotalPagar.text, "###,##0.00"), vbInformation
        Exit Sub
    End If



    confirma = MsgBox("Confirma Pagamento da Conta do Cliente: " & txtCliente.text, vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        Sqlconsulta = "id_prazo = '" & txtid_prazo.text & "'"

        sqlDeletar "Prazo", Sqlconsulta, Me, "N"
        sqlDeletar "Prazoitem", Sqlconsulta, Me, "N"
        sqlDeletar "Prazopagto", Sqlconsulta, Me, "N"

        campo = "datacaixa"
        Scampo = "'" & Format(Now, "YYYYMMDD") & "'"

        If txtCaixaDinheiro.text <> "" Then
            campo = campo & ", valorcaixadinheiro"
            Scampo = Scampo & ", '" & FormatValor(txtCaixaDinheiro.text, 1) & "'"
        End If

        If txtCaixaCartao.text <> "" Then
            campo = campo & ", valorcaixacartao"
            Scampo = Scampo & ", '" & FormatValor(txtCaixaCartao.text, 1) & "'"
        End If

        campo = campo & ", historico"
        Scampo = Scampo & ", '" & Mid("Pagamento conta:" & txtCliente.text, 1, 100) & "'"

        campo = campo & ", id_vendedor"
        Scampo = Scampo & ", '" & txtid_vendedor.text & "'"

        campo = campo & ", status"
        Scampo = Scampo & ", 'P'"

        sqlIncluir "caixa", campo, Scampo, Me, "N"

        frmFichaFicha.txtPagamento.text = "S"

    End If

    Unload Me

    Exit Sub
trata_erro:

End Sub



'--- txtCaixaDinheiro
Private Sub txtCaixaDinheiro_GotFocus()
    txtCaixaDinheiro.BackColor = &H80FFFF
End Sub
Private Sub txtCaixaDinheiro_LostFocus()
    txtCaixaDinheiro.BackColor = &H80000014
    txtCaixaDinheiro.text = Format(txtCaixaDinheiro.text, "###,##0.00")
End Sub
Private Sub txtCaixaDinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtCaixaDinheiro.text = "" Then txtCaixaDinheiro.text = "0,00"
        mDinheiro = txtCaixaDinheiro.text
        lblTotalPagar.Caption = Format(mDinheiro + mCartao, "###,##0.00")
        txtCaixaCartao.SetFocus
    End If
End Sub

'--- txtCaixaCartao
Private Sub txtCaixaCartao_GotFocus()
    txtCaixaCartao.BackColor = &H80FFFF
End Sub
Private Sub txtCaixaCartao_LostFocus()
    txtCaixaCartao.BackColor = &H80000014
    txtCaixaCartao.text = Format(txtCaixaCartao.text, "###,##0.00")
End Sub
Private Sub txtCaixaCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtCaixaCartao.text = "" Then txtCaixaCartao.text = "0,00"
        mCartao = txtCaixaCartao.text
        lblTotalPagar.Caption = Format(mDinheiro + mCartao, "###,##0.00")
        cmdGravar_Click
    End If
End Sub
