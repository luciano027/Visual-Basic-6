VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendas"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtExtrato 
      Height          =   375
      Left            =   1560
      TabIndex        =   46
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPrazo 
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtObservacao 
      Height          =   375
      Left            =   2520
      TabIndex        =   44
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAcesso 
      Height          =   375
      Left            =   3120
      TabIndex        =   43
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtVendedor 
      Height          =   375
      Left            =   3720
      TabIndex        =   42
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   375
      Left            =   4200
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   20
      Top             =   7200
      Width           =   9735
      Begin VB.TextBox txtid_estoque 
         Height          =   375
         Left            =   4560
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSaldo 
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   435
         Width           =   5655
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin Vendas.VistaButton cmdIncluirItem 
         Height          =   375
         Left            =   9180
         TabIndex        =   30
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   ""
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
         Picture         =   "frmVendas.frx":0000
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label Label10 
         Caption         =   "Descrição"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   975
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
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   9735
      End
      Begin VB.Image cmdConsultar 
         Height          =   375
         Left            =   5760
         Picture         =   "frmVendas.frx":0352
         Stretch         =   -1  'True
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Quantidade"
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblPrecovenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7080
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8160
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Preço Venda"
         Height          =   255
         Left            =   7080
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   255
         Left            =   8160
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   285
      Left            =   9000
      TabIndex        =   17
      Text            =   "A"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7035
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   12855
      Begin VB.TextBox txtid_venda 
         Height          =   285
         Left            =   9480
         TabIndex        =   14
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   1215
         TabIndex        =   11
         Top             =   6720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListaProdutos 
         Height          =   5415
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
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
         Caption         =   "Total  (R$)"
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
         Left            =   9960
         TabIndex        =   18
         Top             =   5880
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produtos "
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
         Width           =   12855
      End
      Begin VB.Label lbltotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   9960
         TabIndex        =   19
         Top             =   6120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12855
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11040
         TabIndex        =   50
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblLimite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11040
         TabIndex        =   47
         Top             =   600
         Width           =   1695
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
         TabIndex        =   9
         Top             =   0
         Width           =   12855
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
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblid_venda 
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   7215
      End
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
         Left            =   4080
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nº venda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbltotalcredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9105
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   27236
            MinWidth        =   27236
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtid_saida 
      Height          =   285
      Left            =   9720
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtEstado_venda 
      Height          =   285
      Left            =   9600
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin Vendas.VistaButton cmdsairVenda 
      Height          =   615
      Left            =   11760
      TabIndex        =   33
      Top             =   8400
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
      Picture         =   "frmVendas.frx":065C
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
      Left            =   7920
      TabIndex        =   34
      ToolTipText     =   "Extrato da Venda"
      Top             =   8400
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
      Picture         =   "frmVendas.frx":0766
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCancelarVenda 
      Height          =   615
      Left            =   6360
      TabIndex        =   35
      ToolTipText     =   "Excluir uma venda inteira"
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Excluir"
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
      Picture         =   "frmVendas.frx":0878
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
      Left            =   4800
      TabIndex        =   36
      ToolTipText     =   "Formas de pagamento"
      Top             =   8400
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
      Picture         =   "frmVendas.frx":0C55
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
      Left            =   3240
      TabIndex        =   37
      ToolTipText     =   "Cancela um item de Compra"
      Top             =   8400
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
      Picture         =   "frmVendas.frx":16C7
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCompras 
      Height          =   615
      Left            =   1680
      TabIndex        =   38
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Compras"
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
      Picture         =   "frmVendas.frx":1A19
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdCliente 
      Height          =   615
      Left            =   120
      TabIndex        =   39
      ToolTipText     =   "Define o Cliente para venda a Prazo"
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Cliente"
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
      Picture         =   "frmVendas.frx":1D6B
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdOrcamento 
      Height          =   615
      Left            =   9480
      TabIndex        =   40
      ToolTipText     =   "Orcaçmento"
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Orçamento"
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
      Picture         =   "frmVendas.frx":29BD
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   9225
      Left            =   0
      Picture         =   "frmVendas.frx":2D0F
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   14715
   End
End
Attribute VB_Name = "frmVendas"
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
Dim mControlarSaldo As String
Dim mDescricao As String

Private Sub Form_Activate()
    VerificarMaquina
    If ChaveM = "S" Then txtDescricao.SetFocus Else txtQuantidade.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 13260
    Me.Height = 9880
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    ChaveM = "S"

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendas = Nothing
    MenuPrincipal.AbilidataMenu
End Sub
'--------------------------------------  Botoes ----------------------------------------------------------------
Private Sub cmdCliente_Click()
    VendasClientes
End Sub

Private Sub cmdCompras_Click()
    VendasCompras
End Sub

Private Sub cmdCancelarItem_Click()
    vendasCancelarItem
End Sub

Private Sub cmdPagamento_Click()
    VendasPagamento
End Sub
Private Sub cmdCancelarVenda_Click()
    VendasCancelar
End Sub
Private Sub cmdExtrato_Click()
    ClienteNome = lblCliente.Caption
    VendasExtrato
End Sub
Private Sub cmdOrcamento_Click()
    Vendaorcamento
End Sub


Private Sub cmdsairVenda_Click()
    Unload Me
End Sub

'-----------------------------------------------------------------------------------------------------------------
Private Sub VerificarMaquina()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = " SELECT vendas.status, vendas.maquina, vendas.id_cliente, clientes.id_cliente, clientes.cliente, clientes.observacao, clientes.limite, "
    Sql = Sql & " vendas.id_venda, vendas.datavenda"
    Sql = Sql & " From"
    Sql = Sql & " vendas"
    Sql = Sql & " left join clientes on vendas.id_cliente = clientes.id_cliente"
    Sql = Sql & " Where vendas.status = 'A' and vendas.maquina = '" & MicroBD & "'"

    ' abre um Recrodset da Tabela Tabela
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        lblid_venda.Caption = Tabela("id_venda")
        If VarType(Tabela("id_cliente")) <> vbNull Then lblid_cliente.Caption = Tabela("id_cliente")
        If VarType(Tabela("cliente")) <> vbNull Then lblCliente.Caption = Tabela("cliente")
        If VarType(Tabela("datavenda")) <> vbNull Then lbldataVenda.Caption = Format(Tabela("datavenda"), "DD/MM/YYYY")
        If VarType(Tabela("observacao")) <> vbNull Then txtObservacao.text = Tabela("observacao") Else txtObservacao.text = ""
        If VarType(Tabela("limite")) <> vbNull Then lblLimite.Caption = Tabela("limite") Else lblLimite.Caption = "0"
        Lista ("")
        If lblid_cliente.Caption <> "" Then Atualiza

    Else
        If ChaveM = "N" Then Exit Sub
        VendasCompras
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

    Sql = " SELECT saida.id_venda, saida.id_estoque, saida.quantidade, saida.preco_venda, saida.id_saida,"
    Sql = Sql & " estoques.id_estoque, estoques.descricao, estoques.unidade,"
    Sql = Sql & " (saida.quantidade*saida.preco_venda) AS total, vendas.id_venda"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " LEFT JOIN vendas on saida.id_venda = vendas.id_venda"
    Sql = Sql & " Where vendas.id_venda = '" & lblid_venda.Caption & "'"
    Sql = Sql & " order by descricao"


    ' abre um Recrodset da Tabela Estoques
    If Estoques.State = 1 Then Estoques.Close
    Estoques.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaProdutos.ColumnHeaders.Clear
    ListaProdutos.ListItems.Clear

    ListaProdutos.ColumnHeaders.Add , , "Descrição", 6800
    ListaProdutos.ColumnHeaders.Add , , "Unid", 1000, lvwColumnCenter
    ListaProdutos.ColumnHeaders.Add , , "Quant.", 1500, lvwColumnRight
    ListaProdutos.ColumnHeaders.Add , , "Preço(R$)", 1500, lvwColumnRight
    ListaProdutos.ColumnHeaders.Add , , "Total", 1800, lvwColumnRight

    Gtotal = 0

    If Estoques.BOF = True And Estoques.EOF = True Then Exit Sub
    While Not Estoques.EOF

        If VarType(Estoques("descricao")) <> vbNull Then Set itemx = ListaProdutos.ListItems.Add(, , Estoques("descricao"))
        If VarType(Estoques("unidade")) <> vbNull Then itemx.SubItems(1) = Estoques("unidade") Else itemx.SubItems(1) = ""
        If VarType(Estoques("quantidade")) <> vbNull Then itemx.SubItems(2) = Format(Estoques("quantidade"), "###,##0.000") Else itemx.SubItems(2) = ""
        If VarType(Estoques("preco_venda")) <> vbNull Then itemx.SubItems(3) = Format(Estoques("preco_venda"), "###,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Estoques("total")) <> vbNull Then itemx.SubItems(4) = Format(Estoques("total"), "###,##0.00") Else itemx.SubItems(4) = ""
        If VarType(Estoques("id_saida")) <> vbNull Then itemx.Tag = Estoques("id_saida")
        Gtotal = Gtotal + Estoques("total")

        Estoques.MoveNext

    Wend

    lblTotalGeral.Caption = Format(Gtotal, "###,##0.00")

    If lblid_cliente.Caption <> "" Then Atualiza

    txtDescricao.SetFocus

    'Zebra o listview
    If LVZebra(ListaProdutos, Picture2, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Estoques.State = 1 Then Estoques.Close
    Set Estoques = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaProdutos_DblClick()
    On Error GoTo trata_erro

    txtid_saida.text = ListaProdutos.SelectedItem.Tag

    vendasCancelarItem

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaProdutos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_saida.text = ListaProdutos.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Public Sub listaProdutosCompras()
    Lista ("")
End Sub

Private Sub Vendaorcamento()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim Tabela1 As ADODB.Recordset
    Dim mSaldo As Double
    Dim mQuantidade As Double
    Dim mControlarSaldo As String
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")
    Set Tabela1 = CreateObject("ADODB.Recordset")

    confirma = MsgBox("Confirma cancelamento do  Orçamento", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        Sql = "select saida.*"
        Sql = Sql & " from "
        Sql = Sql & " saida"
        Sql = Sql & " where "
        Sql = Sql & " saida.id_venda = '" & lblid_venda.Caption & "'"

        ' abre um Recrodset da Tabela Tabela
        If Tabela.State = 1 Then Tabela.Close
        Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Tabela.BOF = True And Tabela.EOF = True Then Exit Sub
        While Not Tabela.EOF

            mQuantidade = Tabela("quantidade")

            Sql = "select estoquesaldo.id_estoque, estoquesaldo.saldo, "
            Sql = Sql & " estoques.controlar_saldo, estoques.id_estoque"
            Sql = Sql & " from"
            Sql = Sql & " estoques"
            Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
            Sql = Sql & " where"
            Sql = Sql & " estoques.id_estoque = '" & Tabela("id_estoque") & "'"

            If Tabela1.State = 1 Then Tabela1.Close
            Tabela1.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Tabela1.RecordCount > 0 Then
                If VarType(Tabela1("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela1("Controlar_saldo")
                If VarType(Tabela1("saldo")) <> vbNull Then mSaldo = Tabela1("saldo")

                If mControlarSaldo = "S" Then
                    If VarType(Tabela1("saldo")) = vbNull Then
                        mSaldo = 0
                        mSaldo = mSaldo + mQuantidade

                        campo = "saldo"
                        Scampo = "'" & FormatValor(mSaldo, 1) & "'"

                        campo = campo & ", id_estoque"
                        Scampo = Scampo & ", '" & Tabela("id_estoque") & "'"

                        sqlIncluir "estoquesaldo", campo, Scampo, Me, "N"
                    Else
                        mSaldo = mSaldo + mQuantidade
                        campo = " saldo = '" & FormatValor(mSaldo, 1) & "'"
                        Sqlconsulta = "id_estoque = '" & Tabela("id_estoque") & "'"
                        sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"
                    End If
                End If
            End If
            Tabela.MoveNext
        Wend

        sqlDeletar "vendas", "id_venda = '" & lblid_venda.Caption & "'", Me, "N"
        sqlDeletar "saida", "id_venda = '" & lblid_venda.Caption & "'", Me, "N"

        MsgBox ("Orçamento cancelado com sucesso.."), vbInformation
        Unload Me
    End If


    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    If Tabela1.State = 1 Then Tabela1.Close
    Set Tabela1 = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub VendasCancelar()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim Tabela1 As ADODB.Recordset
    Dim mSaldo As Double
    Dim mQuantidade As Double
    Dim mControlarSaldo As String
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")
    Set Tabela1 = CreateObject("ADODB.Recordset")

    confirma = MsgBox("Confirma Cancelamento da Venda", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        Sql = "select saida.*"
        Sql = Sql & " from "
        Sql = Sql & " saida"
        Sql = Sql & " where "
        Sql = Sql & " saida.id_venda = '" & lblid_venda.Caption & "'"

        ' abre um Recrodset da Tabela Tabela
        If Tabela.State = 1 Then Tabela.Close
        Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Tabela.BOF = True And Tabela.EOF = True Then Exit Sub

        While Not Tabela.EOF

            mQuantidade = Tabela("quantidade")

            Sql = "select estoquesaldo.id_estoque, estoquesaldo.saldo, "
            Sql = Sql & " estoques.controlar_saldo, estoques.id_estoque"
            Sql = Sql & " from"
            Sql = Sql & " estoques"
            Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
            Sql = Sql & " where"
            Sql = Sql & " estoques.id_estoque = '" & Tabela("id_estoque") & "'"

            If Tabela1.State = 1 Then Tabela1.Close
            Tabela1.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Tabela1.RecordCount > 0 Then
                If VarType(Tabela1("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela1("Controlar_saldo")
                If VarType(Tabela1("saldo")) <> vbNull Then mSaldo = Tabela1("saldo")

                If mControlarSaldo = "S" Then
                    If VarType(Tabela1("saldo")) = vbNull Then
                        mSaldo = 0
                        mSaldo = mSaldo + mQuantidade

                        campo = "saldo"
                        Scampo = "'" & FormatValor(mSaldo, 1) & "'"

                        campo = campo & ", id_estoque"
                        Scampo = Scampo & ", '" & Tabela("id_estoque") & "'"

                        sqlIncluir "estoquesaldo", campo, Scampo, Me, "N"
                    Else
                        mSaldo = mSaldo + mQuantidade
                        campo = " saldo = '" & FormatValor(mSaldo, 1) & "'"
                        Sqlconsulta = "id_estoque = '" & Tabela("id_estoque") & "'"
                        sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"
                    End If
                End If
            End If
            Tabela.MoveNext
        Wend


        sqlDeletar "vendas", "id_venda = '" & lblid_venda.Caption & "'", Me, "N"
        sqlDeletar "saida", "id_venda = '" & lblid_venda.Caption & "'", Me, "N"

        MsgBox ("Venda cancelada com sucesso.."), vbInformation
        Unload Me
    End If


    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    If Tabela1.State = 1 Then Tabela1.Close
    Set Tabela1 = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub sair()
    If txtEstado_venda.text = "A" Then
        MsgBox ("Valor fazer pagamento  ou cancelar a venda...."), vbInformation
        Exit Sub
    Else
        Unload Me
        Set frmVendas = Nothing
    End If
End Sub

Private Sub VendasPagamento()
    On Error GoTo trata_erro
    Dim mAcesso As String
    ' conecta ao banco de dados

    If txtAcesso.text <> "" Then
        With frmVendasPagamento

            .lblTotalVenda.Caption = Format(lblTotalGeral.Caption, "###,##0.00")
            .txtCliente.text = lblCliente.Caption
            .txtid_cliente.text = lblid_cliente.Caption
            .txtid_venda.text = lblid_venda.Caption
            .txtid_vendedor.text = txtid_vendedor.text
            .txtDinheiro.text = Format(lblTotalGeral.Caption, "###,##0.00")
            .status.Panels.Item(1).text = "Vendedor: " & txtVendedor.text
            .txtPrazo.text = txtPrazo.text
            .txtVendedor.text = txtVendedor.text
            .txtObservacao.text = txtObservacao.text

            If lblSaldo.Caption < 0 Then .cmdAPrazo.Enabled = False Else .cmdAPrazo.Enabled = True

            .Show 1
        End With
        If txtStatus.text = "P" Then Unload Me
        Exit Sub
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub VendasClientes()
    With frmVendasClientes
        If lblid_cliente.Caption <> "" Then .txtid_cliente.text = lblid_cliente.Caption
        .txtid_venda.text = lblid_venda.Caption
        .Show 1
    End With
End Sub



Private Sub vendasCancelarItem()
    If txtid_saida.text <> "" Then
        With frmVendasCancelarItem
            .txtid_saida.text = txtid_saida.text
            .Show 1
        End With
    Else
        MsgBox ("Favor selecionar um item para cancelar..."), vbInformation
        Exit Sub
    End If
    Lista ("")
End Sub

'''--------------------------------------------------------  Vendas Compras -------------------------------------------

Private Sub VendasCompras()
    txtDescricao.SetFocus
End Sub



Private Sub cmdIncluirItem_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim mSaldoA As Double
    Dim mQuantidadeA As Double

    confirma = MsgBox("Confirma o item ", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        If lblid_venda.Caption = "" Then

            campo = "dataVenda"
            Scampo = "'" & Format(Now, "YYYYMMDD") & "'"

            campo = campo & ", maquina"
            Scampo = Scampo & ", '" & MicroBD & "'"

            campo = campo & ", status"
            Scampo = Scampo & ", 'A'"


            ' Incluir valor na tabela EntradaItens
            sqlIncluir "vendas", campo, Scampo, Me, "N"

            Buscar_id

        End If

        If txtid_estoque.text <> "" Then
            Sqlconsulta = "id_estoque = '" & txtid_estoque.text & "'"

            If mControlarSaldo = "S" Then
                mSaldoA = txtsaldo.text
                mQuantidadeA = txtQuantidade.text

                If mSaldoA < mQuantidadeA Then
                    MsgBox ("Saldo insuficiente. Saldo estoque atual: " & txtsaldo.text), vbInformation

                    txtDescricao.text = ""
                    txtid_estoque.text = ""
                    txtQuantidade.text = ""
                    lblPrecovenda.Caption = ""
                    lblTotal.Caption = ""

                    txtDescricao.SetFocus


                    Exit Sub
                End If

                mSaldo = txtsaldo.text - txtQuantidade.text
                If mSaldo < 0 Then mSaldo = 0
                '-------------- Alteara saldo
                campo = "saldo = '" & FormatValor(mSaldo, 1) & "'"
                sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"
            End If
            '-------------- Alteara dados da compra estoque
            campo = "data_venda = '" & Format(Now, "YYYYMMDD") & "'"
            campo = campo & ", quant_venda = '" & FormatValor(txtQuantidade.text, 1) & "'"
            sqlAlterar "estoques", campo, Sqlconsulta, Me, "N"

            '------------- Inclui tabela Saida
            campo = "id_estoque"
            Scampo = "'" & txtid_estoque.text & "'"

            campo = campo & ", id_venda"
            Scampo = Scampo & ", '" & lblid_venda.Caption & "'"

            campo = campo & ", quantidade"
            Scampo = Scampo & ", '" & FormatValor(txtQuantidade.text, 1) & "'"

            campo = campo & ", preco_venda"
            Scampo = Scampo & ", '" & FormatValor(lblPrecovenda.Caption, 1) & "'"

            campo = campo & ", datasaida"
            Scampo = Scampo & ", '" & Format(Now, "YYYYMMDD") & "'"

            sqlIncluir "Saida", campo, Scampo, Me, "N"

        End If

    End If

    txtDescricao.text = ""
    txtid_estoque.text = ""
    txtQuantidade.text = ""
    lblPrecovenda.Caption = ""
    lblTotal.Caption = ""

    txtDescricao.SetFocus

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(id_venda) as MaxID "
    Sql = Sql & " FROM vendas"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then lblid_venda.Caption = Tabela("maxid") Else lblid_venda.Caption = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub Consulta_estoque_6()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT estoques.id_estoque, estoques.descricao, estoques.preco_venda, estoquesaldo.saldo, "
    Sql = Sql & " estoques.controlar_saldo"
    Sql = Sql & " FROM estoques "
    Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
    Sql = Sql & " WHERE estoques.codigo_est = '" & mDescricao & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_estoque")) <> vbNull Then txtid_estoque.text = Tabela("id_estoque") Else txtid_estoque.text = ""
        If VarType(Tabela("descricao")) <> vbNull Then txtDescricao.text = Tabela("descricao") Else txtDescricao.text = ""
        If VarType(Tabela("preco_venda")) <> vbNull Then lblPrecovenda.Caption = Format(Tabela("preco_venda"), "###,##0.00") Else lblPrecovenda.Caption = ""
        If VarType(Tabela("saldo")) <> vbNull Then txtsaldo.text = Format(Tabela("saldo"), "###,##0.000") Else txtsaldo.text = "0"
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("controlar_saldo")


        'txtQuantidade.SetFocus
        ChaveM = "N"
    Else
        MsgBox ("Produto não encontrado...."), vbInformation
        Exit Sub

    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaEstoque
        .Show 1
    End With

    If IDEstoque <> "" Then
        txtid_estoque.text = IDEstoque
        IDEstoque = ""
    End If

    Sql = "SELECT estoques.id_estoque, estoques.descricao, estoques.preco_venda, estoquesaldo.saldo, "
    Sql = Sql & " estoques.controlar_saldo"
    Sql = Sql & " FROM estoques "
    Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
    Sql = Sql & " WHERE estoques.id_Estoque= '" & txtid_estoque.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_estoque")) <> vbNull Then txtid_estoque.text = Tabela("id_estoque") Else txtid_estoque.text = ""
        If VarType(Tabela("descricao")) <> vbNull Then txtDescricao.text = Tabela("descricao") Else txtDescricao.text = ""
        If VarType(Tabela("preco_venda")) <> vbNull Then lblPrecovenda.Caption = Format(Tabela("preco_venda"), "###,##0.00") Else lblPrecovenda.Caption = ""
        If VarType(Tabela("saldo")) <> vbNull Then txtsaldo.text = Format(Tabela("saldo"), "###,##0.000") Else txtsaldo.text = "0"
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("controlar_saldo")


        'txtQuantidade.SetFocus
        ChaveM = "N"

    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


'-------- Descricao
Private Sub txtDescricao_GotFocus()
    txtDescricao.BackColor = &H80FFFF
End Sub
Private Sub txtDescricao_LostFocus()
    txtDescricao.BackColor = &H80000014
End Sub
Private Sub txtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(txtDescricao.text) < 6 Then mDescricao = strzero(txtDescricao.text, 6, Vbfantes)


        If Len(mDescricao) = 6 And mDescricao <> "000000" Then
            Consulta_estoque_6
        End If

        If mDescricao = "000000" Then
            cmdConsultar_Click
        ElseIf txtid_estoque.text <> "" Then
            txtQuantidade.SetFocus
        Else
            MsgBox ("Mercadoria não cadastrada..."), vbInformation
            txtDescricao.text = ""
            txtDescricao.SetFocus
        End If

    End If
    ' If KeyAscii = vbKeyEscape Then cmdSair_Click
End Sub

'--- txtQuantidade
Private Sub txtquantidade_GotFocus()
    txtQuantidade.BackColor = &H80FFFF
End Sub
Private Sub txtquantidade_LostFocus()
    txtQuantidade.BackColor = &H80000014
    txtQuantidade.text = Format(txtQuantidade.text, "###,##0.00")
End Sub
Private Sub txtquantidade_KeyPress(KeyAscii As Integer)
    On Error GoTo trata_erro
    Dim Gtotal As Double
    If KeyAscii = vbKeyReturn Then
        If txtQuantidade.text <> "" Then
            Gtotal = txtQuantidade.text * lblPrecovenda.Caption
            lblTotal.Caption = Format(Gtotal, "###,##0.00")
            cmdIncluirItem_Click
        End If
    End If
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub VendasExtrato()

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
    Dim strVendanro As String

    Dim mDebito As Double
    Dim mCredito As Double
    Dim mAPagar As Double
    Dim mDaDos As Integer
    Dim mCabecarioDados As String
    Dim mArquivo As String

    Dim strCliente As String

    Dim bRet As Boolean


    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = " SELECT saida.*, estoques.id_estoque, estoques.unidade, estoques.descricao,estoques.codigo_est,"
    Sql = Sql & " (saida.quantidade * saida.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where saida.id_venda = '" & lblid_venda.Caption & "'"
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

    strCliente = ClienteNome
    strVendanro = lblid_venda.Caption

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
        mDaDos = 38 - Len(Tabela("descricao"))
        If mDaDos < 0 Then mDaDos = Len(Tabela("descricao")) - 38
        txtExtrato.text = txtExtrato.text & Mid(Tabela("codigo_est"), 1, 6) & Space(1)
        txtExtrato.text = txtExtrato.text & Mid(Tabela("descricao"), 1, 40) & Space(mDaDos) & Space(3)
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
        txtExtrato.text = txtExtrato.text & "                                                                   continua..." & vbCrLf
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

    txtExtrato.text = txtExtrato.text & "Papelaria " & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "..............................................................." & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "dd/mm/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "Vendedor: " & txtVendedor.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Venda...: " & strVendanro & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Codigo Descricao                                Quant.     Preco      Total  " & vbCrLf
    txtExtrato.text = txtExtrato.text & "------ ---------------------------------------- ---------- ---------- ----------" & vbCrLf

    Return

Rodape:

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    mAPagar = mDebito - mCredito


    txtExtrato.text = txtExtrato.text & Space(51) & " Total a Pagar..R$ " & Alinhar(Format(mAPagar, "###,##0.00"), 10) & vbCrLf


    Do While intLinhafinal < 19
        txtExtrato.text = txtExtrato.text & vbCrLf
        intLinhafinal = intLinhafinal + 1
    Loop



    Return

RodapeObs:

    intNroPagina = intNroPagina + 1


    txtExtrato.text = txtExtrato.text & "Papelaria" & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "............................................................3" & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "mm/dd/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "Vendedor: " & txtVendedor.text & vbCrLf & vbCrLf & vbCrLf

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    txtExtrato.text = txtExtrato.text & "Observação " & vbCrLf
    txtExtrato.text = txtExtrato.text & txtObservacao.text & vbCrLf

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

Public Function ImprimeArq(ByVal sArq As String) As Boolean
    Dim lArq As Long
    Dim sTexto As String

    If Dir$(sArq) = "" Then
        ImprimeArq = False
        Exit Function
    End If
    lArq = FreeFile()
    Open sArq For Binary Access Read As lArq
    sTexto = Space$(LOF(lArq))
    Get #lArq, , sTexto
    Close lArq
    Printer.Print sTexto
    Printer.EndDoc
    ImprimeArq = True
End Function

Private Function Alinhar(texto As String, Largura As Integer)
    Alinhar = String(Largura - Len(texto), " ") & texto
End Function

Private Sub Atualiza()
    On Error GoTo trata_erro
    Dim mCredito As Double
    Dim mDebito As Double
    Dim mPagar As Double
    Dim mSaldor As Double
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT clientes.prazo "
    Sql = Sql & " From"
    Sql = Sql & " clientes"
    Sql = Sql & " where"
    Sql = Sql & " clientes.id_cliente = '" & lblid_cliente.Caption & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("prazo")) <> vbNull Then txtPrazo.text = Tabela("prazo")
    End If


    Sql = "SELECT prazo.id_cliente,"
    Sql = Sql & " SUM(prazopagto.ValorPagto) As total"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN prazopagto ON prazo.id_prazo = prazopagto.id_prazo"
    Sql = Sql & " where"
    Sql = Sql & " prazo.id_cliente = '" & lblid_cliente.Caption & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("total")) <> vbNull Then mCredito = Tabela("total")
    Else
        mCredito = 0
    End If


    Sql = "SELECT prazo.id_cliente,"
    Sql = Sql & " SUM(prazoitem.quantidade * prazoitem.preco_venda) As total"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN prazoitem ON prazo.id_prazo = prazoitem.id_prazo"
    Sql = Sql & " where"
    Sql = Sql & " prazo.id_cliente = '" & lblid_cliente.Caption & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("total")) <> vbNull Then mDebito = Tabela("total")
    Else
        mDebito = 0
    End If

    mPagar = mDebito - mCredito
    mSaldor = lblLimite.Caption - mPagar

    '  lbltotalDebito.Caption = Format(mDebito, "###,##0.00")

    lbltotalCredito.Caption = Format(mPagar, "###,##0.00")
    lblSaldo.Caption = Format(mSaldor - lblTotalGeral.Caption, "###,##0.00")

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


