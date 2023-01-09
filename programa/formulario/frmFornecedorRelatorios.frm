VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmFornecedorRelatorios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatorios"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Relatorios"
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.OptionButton optInventario 
         Caption         =   "Inventario"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.OptionButton optProdutoCadastrado 
         Caption         =   "Produtos Cadastrado"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton optTabelaPreeco 
         Caption         =   "Tabela de Pre�os"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton optTabelaCompra 
         Caption         =   "Tabela Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   4095
      End
      Begin VB.OptionButton optSaldoMinimo 
         Caption         =   "Saldo Minimo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   3975
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3690
      Width           =   5205
      _ExtentX        =   9181
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
      TabIndex        =   7
      Top             =   3000
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
      Picture         =   "frmFornecedorRelatorios.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdRelatorios 
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "  Imprimir"
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
      Picture         =   "frmFornecedorRelatorios.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "frmFornecedorRelatorios.frx":021C
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   5235
   End
End
Attribute VB_Name = "frmFornecedorRelatorios"
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


Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Me.Width = 5130
    Me.Height = 4455
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFornecedorRelatorios = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFornecedorRelatorios = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

