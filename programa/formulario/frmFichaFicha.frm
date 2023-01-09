VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaFicha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTipo_acesso 
      Height          =   285
      Left            =   5160
      TabIndex        =   32
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtExtrato 
      Height          =   375
      Left            =   4560
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAcesso 
      Height          =   285
      Left            =   3720
      TabIndex        =   30
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtvendedor 
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   2280
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPagamento 
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "N"
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtid_prazoItem 
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   14
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid_prazo 
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   14055
      Begin VB.Label lblid_cliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   600
         Width           =   11055
      End
      Begin VB.Label lbldataVenda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Venda"
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
         TabIndex        =   5
         Top             =   0
         Width           =   14055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   12135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   14055
      Begin VB.Frame frSenha 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Senha"
         Height          =   855
         Left            =   5040
         TabIndex        =   33
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox txtSenha 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   34
            Top             =   240
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView ListaFicha 
         Height          =   5415
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9551
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10080
         Top             =   7200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   21
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":0CDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":266E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":4000
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":5992
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":7324
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":7FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":8CD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":99B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":A68E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":B36A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":BC46
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":C922
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":D5FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":E2DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":EBBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":F89A
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":10176
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":10E52
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":127E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFichaFicha.frx":1417A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListaPrazo 
         Height          =   5415
         Left            =   11160
         TabIndex        =   12
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   9551
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "Total Compras"
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
         Left            =   8400
         TabIndex        =   27
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label lbltotalGeral 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8400
         TabIndex        =   26
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label lblPagar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   6000
         Width           =   3255
      End
      Begin VB.Label lblCredito 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11160
         TabIndex        =   24
         Top             =   6000
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Total a Pagar"
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
         Left            =   120
         TabIndex        =   23
         Top             =   5760
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "Total Credito"
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
         Left            =   11160
         TabIndex        =   22
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos"
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
         Left            =   11160
         TabIndex        =   11
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label lblConsulta 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Itens Ficha do Cliente"
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
         TabIndex        =   3
         Top             =   0
         Width           =   11175
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8970
      Width           =   14355
      _ExtentX        =   25321
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
   Begin Vendas.VistaButton cmdsairVenda 
      Height          =   615
      Left            =   13080
      TabIndex        =   17
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "frmFichaFicha.frx":14A56
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdExtrato 
      Height          =   615
      Left            =   9840
      TabIndex        =   18
      ToolTipText     =   "Extrato da Venda"
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Extrato"
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
      Picture         =   "frmFichaFicha.frx":14B60
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCredito 
      Height          =   615
      Left            =   6720
      TabIndex        =   19
      ToolTipText     =   "Formas de pagamento"
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      Picture         =   "frmFichaFicha.frx":14C72
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdPagamento 
      Height          =   615
      Left            =   8280
      TabIndex        =   20
      ToolTipText     =   "Formas de pagamento"
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Pagamento"
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
      Picture         =   "frmFichaFicha.frx":156E4
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCancelarItem 
      Height          =   615
      Left            =   11400
      TabIndex        =   21
      ToolTipText     =   "Cancela um item de Compra"
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Cancelar Item"
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
      Picture         =   "frmFichaFicha.frx":16156
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   -120
      Picture         =   "frmFichaFicha.frx":164A8
      Stretch         =   -1  'True
      Top             =   -1680
      Width           =   14640
   End
End
Attribute VB_Name = "frmFichaFicha"
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
Dim mTipo As String


Private Sub cmdCancelarItem_Click()

    mTipo = "D"
    txtSenha.text = ""
    frSenha.Visible = True
    txtSenha.SetFocus

End Sub

Private Sub cmdCredito_Click()

    mTipo = "C"
    txtSenha.text = ""
    frSenha.Visible = True
    txtSenha.SetFocus
    '  credito
End Sub

Private Sub cmdExtrato_Click()
    ClienteNome = lblCliente.Caption
    extrato
End Sub

Private Sub cmdPagamento_Click()

    mTipo = "P"
    txtSenha.text = ""
    frSenha.Visible = True
    frSenha.Caption = "Senha Caixa"
    txtSenha.SetFocus
    ' VendasPagamento
End Sub

Private Sub cmdsairVenda_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Atualiza
End Sub

Private Sub Form_Load()
    Me.Width = 14445
    Me.Height = 9720
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)



    MenuPrincipal.AbilidataMenu
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaFicha = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub Atualiza()
    On Error GoTo trata_erro
    Dim Prazo As ADODB.Recordset

    ' conecta ao banco de dados
    Set Prazo = CreateObject("ADODB.Recordset")

    Sql = " SELECT prazo.*, clientes.id_cliente, clientes.cliente"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN clientes ON clientes.id_cliente = prazo.id_cliente"
    Sql = Sql & " Where prazo.id_prazo = '" & txtid_prazo.text & "'"

    ' abre um Recrodset da Tabela Prazo
    If Prazo.State = 1 Then Prazo.Close
    Prazo.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Prazo.RecordCount > 0 Then
        If VarType(Prazo("id_cliente")) <> vbNull Then lblid_cliente.Caption = Prazo("id_cliente") Else lblid_cliente.Caption = ""
        If VarType(Prazo("cliente")) <> vbNull Then lblCliente.Caption = Prazo("cliente") Else lblCliente.Caption = ""
        If VarType(Prazo("data_venda")) <> vbNull Then lbldataVenda.Caption = Format(Prazo("data_venda"), "DD/MM/YYYY") Else lbldataVenda.Caption = ""
        Lista ("")
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Estoques As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim Gtotal As Double

    ' conecta ao banco de dados
    Set Estoques = CreateObject("ADODB.Recordset")

    Sql = " SELECT prazoitem.*, estoques.id_estoque, estoques.unidade, estoques.descricao,"
    Sql = Sql & " (prazoitem.quantidade * prazoitem.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " prazoitem"
    Sql = Sql & " LEFT JOIN estoques ON prazoitem.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where prazoitem.id_prazo = '" & txtid_prazo.text & "'"
    Sql = Sql & " order by estoques.descricao"


    ' abre um Recrodset da Tabela Estoques
    If Estoques.State = 1 Then Estoques.Close
    Estoques.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaFicha.ColumnHeaders.Clear
    ListaFicha.ListItems.Clear

    ListaFicha.ColumnHeaders.Add , , "Descrição", 4800
    ListaFicha.ColumnHeaders.Add , , "Data ", 1700, lvwColumnCenter
    ListaFicha.ColumnHeaders.Add , , "Quant.", 1500, lvwColumnRight
    ListaFicha.ColumnHeaders.Add , , "Preço ", 1400, lvwColumnRight
    ListaFicha.ColumnHeaders.Add , , "Total", 1200, lvwColumnRight

    Gtotal = 0

    If Estoques.BOF = True And Estoques.EOF = True Then Exit Sub
    While Not Estoques.EOF

        If VarType(Estoques("descricao")) <> vbNull Then Set itemx = ListaFicha.ListItems.Add(, , Estoques("descricao"))
        If VarType(Estoques("dataCompra")) <> vbNull Then itemx.SubItems(1) = Format(Estoques("dataCompra"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Estoques("quantidade")) <> vbNull Then itemx.SubItems(2) = Format(Estoques("quantidade"), "###,##0.000") Else itemx.SubItems(2) = ""
        If VarType(Estoques("preco_venda")) <> vbNull Then itemx.SubItems(3) = Format(Estoques("preco_venda"), "###,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Estoques("total")) <> vbNull Then itemx.SubItems(4) = Format(Estoques("total"), "###,##0.00") Else itemx.SubItems(4) = ""
        If VarType(Estoques("id_prazoitem")) <> vbNull Then itemx.Tag = Estoques("id_prazoitem")
        Gtotal = Gtotal + Estoques("total")

        Estoques.MoveNext


    Wend

    lblTotalGeral.Caption = Format(Gtotal, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaFicha, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Estoques.State = 1 Then Estoques.Close
    Set Estoques = Nothing

    ListaCredito ("")

    lblPagar.Caption = Format(lblTotalGeral.Caption - lblCredito.Caption, "###,##0.00")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub ListaFicha_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_prazoitem.text = ListaFicha.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub ListaCredito(SQconsulta As String)
    On Error GoTo trata_erro
    Dim PrazoPagtos As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim Gtotal As Double

    ' conecta ao banco de dados
    Set PrazoPagtos = CreateObject("ADODB.Recordset")

    Sql = " SELECT prazopagto.*"
    Sql = Sql & " From"
    Sql = Sql & " prazopagto"
    Sql = Sql & " Where prazopagto.id_prazo = '" & txtid_prazo.text & "'"
    Sql = Sql & " order by prazopagto.dataPagto"


    ' abre um Recrodset da Tabela PrazoPagtos
    If PrazoPagtos.State = 1 Then PrazoPagtos.Close
    PrazoPagtos.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaPrazo.ColumnHeaders.Clear
    ListaPrazo.ListItems.Clear

    ListaPrazo.ColumnHeaders.Add , , "Data", 1500
    ListaPrazo.ColumnHeaders.Add , , "Credito", 1200, lvwColumnRight

    Gtotal = 0

    If PrazoPagtos.BOF = True And PrazoPagtos.EOF = True Then Exit Sub
    While Not PrazoPagtos.EOF

        If VarType(PrazoPagtos("datapagto")) <> vbNull Then Set itemx = ListaPrazo.ListItems.Add(, , Format(PrazoPagtos("datapagto"), "DD/MM/YYYY"))
        If VarType(PrazoPagtos("valorPagto")) <> vbNull Then itemx.SubItems(1) = Format(PrazoPagtos("valorPagto"), "###,##0.00") Else itemx.SubItems(1) = ""
        If VarType(PrazoPagtos("id_prazopagto")) <> vbNull Then itemx.Tag = PrazoPagtos("id_prazopagto")
        Gtotal = Gtotal + PrazoPagtos("valorPagto")

        PrazoPagtos.MoveNext


    Wend

    lblCredito.Caption = Format(Gtotal, "###,##0.00")

    ' lblPagar.Caption = Format(lbltotalGeral.Caption - lblCredito.Caption, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaPrazo, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If PrazoPagtos.State = 1 Then PrazoPagtos.Close
    Set PrazoPagtos = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Cancelar()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim mQuantidade As Double
    Dim mdataCompra As Date
    Dim mTotalVenda As Double
    Dim mIdestoque As String

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    If txtid_prazoitem.text <> "" Then

        confirma = MsgBox("Confirma a exclusão do item ", vbQuestion + vbYesNo, "Incluir")
        If confirma = vbYes Then

            Sql = " select prazoitem.quantidade, prazoitem.dataCompra, prazoitem.id_estoque,"
            Sql = Sql & " (prazoitem.quantidade*prazoitem.preco_venda) as TotalVenda"
            Sql = Sql & " From"
            Sql = Sql & " prazoitem"
            Sql = Sql & " where "
            Sql = Sql & " prazoitem.id_prazoitem = '" & txtid_prazoitem.text & "'"

            If Tabela.State = 1 Then Tabela.Close
            Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Tabela.RecordCount > 0 Then
                If VarType(Tabela("quantidade")) <> vbNull Then mQuantidade = Tabela("quantidade") Else mQuantidade = 0
                If VarType(Tabela("datacompra")) <> vbNull Then mdataCompra = Tabela("datacompra") Else mdataCompra = ""
                If VarType(Tabela("totalvenda")) <> vbNull Then mTotalVenda = Tabela("totalvenda") Else mTotalVenda = 0
                If VarType(Tabela("id_estoque")) <> vbNull Then mIdestoque = Tabela("id_estoque") Else mIdestoque = ""
            End If

            ' ----------------------- Exclui item tabela prazoitem
            Sqlconsulta = " prazoitem.id_prazoitem = '" & txtid_prazoitem.text & "'"
            sqlDeletar "Prazoitem", Sqlconsulta, Me, "N"

            '----------------------- Inlcui saldo no estoque
            Sql = " select estoquesaldo.saldo, estoquesaldo.id_estoque"
            Sql = Sql & " From"
            Sql = Sql & " estoquesaldo"
            Sql = Sql & " where "
            Sql = Sql & " estoquesaldo.id_estoque = '" & mIdestoque & "'"

            If Tabela.State = 1 Then Tabela.Close
            Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Tabela.RecordCount > 0 Then

                mSaldo = Tabela("saldo")

                Sqlconsulta = "id_estoque = '" & mIdestoque & "'"

                mSaldo = mSaldo + mQuantidade
                '-------------- Alteara saldo
                campo = "saldo = '" & FormatValor(mSaldo, 1) & "'"
                sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"

            Else

                campo = "id_estoque"
                Scampo = "'" & mIdestoque & "'"

                campo = campo & ", saldo"
                Scampo = Scampo & ", '" & FormatValor(mQuantidade, 1) & "'"

                sqlIncluir "Estoquesaldo", campo, Scampo, Me, "N"

            End If

        End If
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub sair()
    Unload Me
    Set frmFichaFicha = Nothing
End Sub

Private Sub credito()
    With frmFichaCredito
        .txtid_prazo.text = txtid_prazo.text
        .txtCliente.text = lblCliente.Caption
        .txtid_vendedor.text = txtid_vendedor.text
        .txtVendedor.text = txtVendedor.text
        .Show 1
    End With
    Lista ("")
End Sub

Private Sub VendasPagamento()
    On Error GoTo trata_erro
    ' conecta ao banco de dados
    With frmFichaPagamento
        .txtid_vendedor.text = txtid_vendedor.text
        .txtid_prazo.text = txtid_prazo.text
        .txtCliente.text = lblCliente.Caption
        .txtTotalPagar.text = lblPagar.Caption
        .lblTotalPagar.Caption = lblPagar.Caption
        .StatusBarVendedor.Panels.Item(1).text = "Caixa: " & txtVendedor.text
        .Show 1
    End With


    If txtPagamento.text = "S" Then
        MsgBox ("Pagamento efetuado com sucesso.."), vbInformation
        Unload Me
    End If

    Lista ("")

    Exit Sub

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub extrato()
'---------------------------------------  Vendas Extrato -----------------------------------
    On Error GoTo trata_erro
    Dim intLinhaInicial As Integer
    Dim intLinhafinal As Integer
    Dim intNroPagina As Integer
    Dim intX As Integer
    Dim strCustFileName As String
    Dim strBackSlash As String
    Dim intCustFileNbr As Integer

    Dim strFirstName As String
    Dim strLastName As String
    Dim strAddr As String
    Dim strCity As String
    Dim strState As String
    Dim strZip As String

    Dim mDebito As Double
    Dim mCredito As Double
    Dim mAPagar As Double
    Dim mDaDos As Integer
    Dim mCabecarioDados As String
    Dim mArquivo As String
    Dim mDescricao As String

    Dim strCliente As String

    Dim bRet As Boolean


    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    If lblid_cliente.Caption <> "" Then
        Sql = "SELECT SUM(valorpagto) AS totalcredito, prazo.id_prazo, prazo.id_prazo,"
        Sql = Sql & " clientes.id_cliente, clientes.cliente, clientes.tel2"
        Sql = Sql & " From"
        Sql = Sql & " prazopagto"
        Sql = Sql & " left join prazo ON  prazopagto.id_prazo = prazo.id_prazo"
        Sql = Sql & " left join clientes on prazo.id_cliente = clientes.id_cliente"
        Sql = Sql & " Where"
        Sql = Sql & " Prazo.id_cliente = '" & lblid_cliente.Caption & "'"

        If Tabela.State = 1 Then Tabela.Close
        Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Tabela.RecordCount > 0 Then
            If VarType(Tabela("totalcredito")) <> vbNull Then mCredito = Tabela("totalcredito")
            If VarType(Tabela("cliente")) <> vbNull Then strCliente = Tabela("Cliente")
            '  If VarType(Tabela("tel2")) <> vbNull Then strCliente = strCliente & " - " & Tabela("tela2")
        Else
            mCredito = 0
        End If
    Else
        mCredito = 0
    End If

    strCliente = ClienteNome

    Sql = " SELECT prazoitem.*, estoques.id_estoque, estoques.unidade, estoques.descricao, estoques.codigo_est,"
    Sql = Sql & " (prazoitem.quantidade * prazoitem.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " prazoitem"
    Sql = Sql & " LEFT JOIN estoques ON prazoitem.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where prazoitem.id_prazo = '" & txtid_prazo.text & "'"
    Sql = Sql & " order by estoques.descricao"


    ' ------------------- Verificar a existencia do arquivo ------------------------
    mArquivo = Dir(strgExtrato)
    If mArquivo = "maq" & MicroBD & ".txt" Then
        Kill (strgExtrato)
    End If
    ' ---------------------------------------------------------------------------------

    intLinhaInicial = 1
    intLinhafinal = 1  '-----> 19
    intNroPagina = 0

    txtExtrato.text = ""

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not Tabela.EOF
        ' read and print all the records in the input file
        If intLinhaInicial = 1 Then
            GoSub Cabecario
            intLinhafinal = 1
            intLinhaInicial = 1
        End If
        If intLinhafinal = 16 Then
            GoSub Cabecario
            intLinhafinal = 1
            intLinhaInicial = 1
        End If
        ' print a line of data
        mDaDos = 30 - Len(Mid(Tabela("descricao"), 1, 30))
        If mDaDos < 0 Then mDaDos = Len(Tabela("descricao")) - 30


        If mDaDos = 0 Then
            mDescricao = Mid(Tabela("descricao"), 1, 30) & "-" & Format(Tabela("dataCompra"), "DD/MM/YY")
        Else
            mDescricao = Mid(Tabela("descricao") & Space(mDaDos), 1, 30) & "-" & Format(Tabela("dataCompra"), "DD/MM/YY")
        End If

        txtExtrato.text = txtExtrato.text & Mid(Tabela("codigo_est"), 1, 6) & Space(1)
        txtExtrato.text = txtExtrato.text & mDescricao & Space(3)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("quantidade"), "###,##0.00"), 10) & Space(1)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("Preco_venda"), "###,##0.00"), 10) & Space(1)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("total"), "###,##0.00"), 10) & vbCrLf

        mDebito = mDebito + Tabela("total")

        intLinhaInicial = intLinhaInicial + 1
        intLinhafinal = intLinhafinal + 1

        If intLinhaInicial = 15 Then intLinhaInicial = 1

        Tabela.MoveNext
    Wend


    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    GoSub Rodape


    GeraTXT (strgExtrato)

    Call Shell(ArqImprime, vbNormalFocus)


    Exit Sub


    ' internal subroutine to print report headings
    '------------
Cabecario:
    If intNroPagina > 0 Then
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & String(80, "-") & vbCrLf
        txtExtrato.text = txtExtrato.text & "                                                                   continua....." & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf

    Else
        txtExtrato.text = txtExtrato.text & Chr(27) & Chr(120) & Chr(48) & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf

    End If
    ' increment the page counter
    intNroPagina = intNroPagina + 1

    ' Print 4 blank lines, which provides a for top margin. These four lines do NOT
    ' count toward the limit of 60 lines.

    txtExtrato.text = txtExtrato.text & "Papelaria" & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "......................................................." & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "dd/mm/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Codigo Descricao                          Data  Quant.     Preco      Total  " & vbCrLf
    txtExtrato.text = txtExtrato.text & "------ ---------------------------------------- ---------- ---------- ----------" & vbCrLf

    Return

Rodape:

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    mAPagar = mDebito - mCredito


    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Debito..R$ " & Alinhar(Format(mDebito, "###,##0.00"), 10) & vbCrLf
    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Credito.R$ " & Alinhar(Format(mCredito, "###,##0.00"), 10) & vbCrLf
    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Pagar...R$ " & Alinhar(Format(mAPagar, "###,##0.00"), 10) & vbCrLf


    Do While intLinhafinal < 19
        txtExtrato.text = txtExtrato.text & vbCrLf
        intLinhafinal = intLinhafinal + 1
    Loop



    Return


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


Private Sub GeraTXT(ByVal strFile As String)
    Dim FSO As New FileSystemObject
    Dim iARQ As TextStream

    On Error GoTo Erro_GeraLog

    Set FSO = New FileSystemObject

    If Dir$(strgExtrato) <> vbNullString Then
        Set iARQ = FSO.OpenTextFile(strFile, ForAppending)
    Else
        Set iARQ = FSO.CreateTextFile(strFile, False)
    End If

    iARQ.WriteLine txtExtrato.text

    Set FSO = Nothing
    Set iARQ = Nothing

    Exit Sub

Erro_GeraLog:
    MsgBox "Erro ao gerar !", vbCritical, "ERRO!"
    ' FinalizaAtualizador
End Sub


Private Function Alinhar(texto As String, Largura As Integer)
    Alinhar = String(Largura - Len(texto), " ") & texto
End Function

'-------- senha
Private Sub txtsenha_GotFocus()
    txtSenha.BackColor = &H80FFFF
End Sub
Private Sub txtsenha_LostFocus()
    txtSenha.BackColor = &H80000014
End Sub
Private Sub txtsenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then verifica_senha
End Sub

Private Sub verifica_senha()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim mTipoAcesso As String
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")

    frSenha.Visible = False


    Sql = "Select vendedores.tipo_acesso, vendedores.vendedor "
    Sql = Sql & " from "
    Sql = Sql & " vendedores"
    Sql = Sql & " where vendedores.acesso = '" & txtSenha.text & "'"

    ' abre um Recrodset da Tabela Tabela
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("tipo_acesso")) <> vbNull Then mTipoAcesso = Tabela("tipo_acesso") Else mTipoAcesso = ""
        If VarType(Tabela("tipo_acesso")) <> vbNull Then txtVendedor.text = Tabela("vendedor") Else mTipoAcesso = ""
        Else
        Exit Sub
    End If



    If mTipoAcesso = "P" And mTipo = "C" Then
        credito
        Exit Sub
    End If

    If mTipoAcesso = "P" And mTipo = "P" Then
        VendasPagamento
        Exit Sub
    End If

    If mTipoAcesso = "A" And mTipo = "D" Then
         Cancelar
        Exit Sub
    End If

  
    SemAcesso

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub SemAcesso()
    MsgBox ("Você não tem autorização para este tipo de acesso.."), vbInformation
End Sub

