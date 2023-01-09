VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Credito"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   4335
      Begin VB.Label lblTotalCredito 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   18
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total de Credito (R$)"
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
         Width           =   4335
      End
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
         TabIndex        =   16
         Top             =   -1560
         Width           =   7695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   2400
      TabIndex        =   11
      Top             =   120
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
         TabIndex        =   12
         Top             =   345
         Width           =   1815
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
         TabIndex        =   14
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label2 
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
   End
   Begin VB.TextBox txtVendedor 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtCliente 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
         TabIndex        =   6
         Top             =   345
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         TabIndex        =   8
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         TabIndex        =   2
         Top             =   -1560
         Width           =   7695
      End
   End
   Begin VB.TextBox txtid_Prazo 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
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
      Picture         =   "frmFichaCredito.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3105
      Width           =   4650
      _ExtentX        =   8202
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "Credito"
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
      Picture         =   "frmFichaCredito.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   -480
      Picture         =   "frmFichaCredito.frx":065C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5115
   End
End
Attribute VB_Name = "frmFichaCredito"
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
    Me.Width = 4740
    Me.Height = 3855
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaCredito = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFichaCredito = Nothing
End Sub

Private Sub cmdGravar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar_Click
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    campo = "datapagto"
    Scampo = "'" & Format(Now, "YYYYMMDD") & "'"

    If lbltotalCredito.Caption <> "" Then
        campo = campo & ", valorPagto"
        Scampo = Scampo & ", '" & FormatValor(lbltotalCredito.Caption, 1) & "'"
    Else
        MsgBox ("Valor não pode ficar em branco.."), vbInformation
        txtCaixaDinheiro.SetFocus
        Exit Sub
    End If

    campo = campo & ", id_prazo"
    Scampo = Scampo & ", '" & txtid_prazo.text & "'"

    campo = campo & ", id_vendedor"
    Scampo = Scampo & ", '" & txtid_vendedor.text & "'"


    ' Incluir valor na tabela EntradaItens
    sqlIncluir "prazopagto", campo, Scampo, Me, "N"

    Buscar_id

    campo = "datacaixa"
    Scampo = "'" & Format(Now, "YYYYMMDD") & "'"

    campo = campo & ", id_vendedor"
    Scampo = Scampo & ", '" & txtid_vendedor.text & "'"

    If txtCaixaDinheiro.text <> "" Then
        campo = campo & ", valorcaixadinheiro"
        Scampo = Scampo & ", '" & FormatValor(txtCaixaDinheiro.text, 1) & "'"
    End If

    If txtCaixaCartao.text <> "" Then
        campo = campo & ", valorcaixacartao"
        Scampo = Scampo & ", '" & FormatValor(txtCaixaCartao.text, 1) & "'"
    End If

    campo = campo & ", id_prazoPagto"
    Scampo = Scampo & ", '" & mID_prazo & "'"

    campo = campo & ", historico"
    Scampo = Scampo & ", '" & Mid("Credito conta:" & txtCliente.text, 1, 100) & "'"

    campo = campo & ", status"
    Scampo = Scampo & ", 'P'"

    sqlIncluir "caixa", campo, Scampo, Me, "N"

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(id_prazopagto) as MaxID "
    Sql = Sql & " FROM prazopagto"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then mID_prazo = Tabela("maxid") Else mID_prazo = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

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
        mDinheiro = txtCaixaDinheiro.text
        lbltotalCredito.Caption = Format(mDinheiro + mCartao, "###,##0.00")
        txtCaixaCartao.SetFocus
    End If
End Sub

'--- txtCaixaCartao
Private Sub txtCaixaCartao_GotFocus()
    txtCaixaCartao.BackColor = &H80FFFF
End Sub
Private Sub txtCaixaCartao_LostFocus()
    txtCaixaCartao.BackColor = &H80000014
    If txtCaixaCartao.text = "" Then txtCaixaCartao.text = "0,00"
    txtCaixaCartao.text = Format(txtCaixaCartao.text, "###,##0.00")
End Sub
Private Sub txtCaixaCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtCaixaCartao.text = "" Then txtCaixaCartao.text = "0,00"
        mCartao = txtCaixaCartao.text
        lbltotalCredito.Caption = Format(mDinheiro + mCartao, "###,##0.00")
        cmdGravar_Click
    End If
End Sub

