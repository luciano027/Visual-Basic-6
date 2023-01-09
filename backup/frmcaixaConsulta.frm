VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcaixaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caixa Consulta"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTotalVenda 
      Height          =   285
      Left            =   6480
      TabIndex        =   40
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frPagamento 
      Caption         =   "Dados para Pagamento"
      Height          =   7575
      Left            =   5640
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Frame frSenha 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Senha Administrador"
         Height          =   855
         Left            =   3120
         TabIndex        =   72
         Top             =   2760
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
            TabIndex        =   73
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame frameFormaPagameto 
         BackColor       =   &H0000FFFF&
         Height          =   6495
         Left            =   2280
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   3855
            Left            =   240
            TabIndex        =   56
            Top             =   1320
            Width           =   3015
            Begin VB.Label lblBoletoC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   120
               TabIndex        =   76
               Top             =   2280
               Width           =   2775
            End
            Begin VB.Label lblCartaoC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   120
               TabIndex        =   75
               Top             =   1440
               Width           =   2775
            End
            Begin VB.Label lblDinheiroC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   120
               TabIndex        =   74
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label Label21 
               Caption         =   "Boleto"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Forma de Pagamento (R$)"
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
               TabIndex        =   61
               Top             =   0
               Width           =   3255
            End
            Begin VB.Label Label17 
               Caption         =   "Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Cartão"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Desconto"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   2880
               Width           =   1095
            End
            Begin VB.Label lblDesconto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   120
               TabIndex        =   57
               Top             =   3120
               Width           =   2775
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   975
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   4575
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Total da Venda (R$)"
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
               TabIndex        =   55
               Top             =   0
               Width           =   4815
            End
            Begin VB.Label lblTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0,00"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   495
               Left            =   120
               TabIndex        =   54
               Top             =   360
               Width           =   4335
            End
         End
         Begin VB.TextBox txtCalculoTroco 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   240
            TabIndex        =   52
            Text            =   "0,00"
            Top             =   5640
            Width           =   1815
         End
         Begin Vendas.VistaButton btnPagamento 
            Height          =   615
            Left            =   3360
            TabIndex        =   47
            Top             =   3840
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
            Picture         =   "frmcaixaConsulta.frx":0000
            Pictures        =   1
            UseMaskColor    =   -1  'True
            MaskColor       =   65280
            Enabled         =   -1  'True
            NoBackground    =   0   'False
            BackColor       =   16777215
            PictureOffset   =   0
         End
         Begin Vendas.VistaButton btnSairPagamento 
            Height          =   615
            Left            =   3360
            TabIndex        =   48
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
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
            Picture         =   "frmcaixaConsulta.frx":0A72
            Pictures        =   1
            UseMaskColor    =   -1  'True
            MaskColor       =   65280
            Enabled         =   -1  'True
            NoBackground    =   0   'False
            BackColor       =   16777215
            PictureOffset   =   0
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Caption         =   "R$ Calculo Troco"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Caption         =   "Troco"
            Height          =   255
            Left            =   2880
            TabIndex        =   50
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Label lblTroco 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   49
            Top             =   5640
            Width           =   1815
         End
         Begin VB.Image Image3 
            Height          =   6180
            Left            =   120
            Picture         =   "frmcaixaConsulta.frx":0B7C
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4755
         End
      End
      Begin VB.TextBox txtDinheiroD 
         Height          =   285
         Left            =   2280
         TabIndex        =   44
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtExtrato 
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtid_venda 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView ListaSaida 
         Height          =   4215
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7435
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
         NumItems        =   0
      End
      Begin Vendas.VistaButton cmdsairPG 
         Height          =   615
         Left            =   8760
         TabIndex        =   36
         Top             =   6840
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
         Picture         =   "frmcaixaConsulta.frx":1C7D
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin Vendas.VistaButton cmdAvista 
         Height          =   615
         Left            =   5640
         TabIndex        =   37
         Top             =   6840
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
         Picture         =   "frmcaixaConsulta.frx":1D87
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
         Left            =   7200
         TabIndex        =   42
         ToolTipText     =   "Extrato da Venda"
         Top             =   6840
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
         Picture         =   "frmcaixaConsulta.frx":27F9
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin Vendas.VistaButton cmdExcluir 
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   6840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Caption         =   "Excluir Venda"
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
         Picture         =   "frmcaixaConsulta.frx":290B
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label lblTotaldaVenda 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   7680
         TabIndex        =   63
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Total Venda"
         Height          =   255
         Left            =   7680
         TabIndex        =   62
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Itens da Venda"
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
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   9735
      End
      Begin VB.Label lblCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   1920
         Width           =   5295
      End
      Begin VB.Label Label12 
         Caption         =   "Cliente         :"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblHistorico 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         Top             =   1500
         Width           =   5295
      End
      Begin VB.Label lblVendedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   31
         Top             =   1050
         Width           =   5295
      End
      Begin VB.Label lbldata_caixa 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblid_venda 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Data Venda:"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Historico      :"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Vendedor     :"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Nº da Venda:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.TextBox txtid_caixa 
      Height          =   285
      Left            =   5640
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   13680
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListaCaixa 
      Height          =   5175
      Left            =   5640
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9128
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
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7815
      Width           =   15675
      _ExtentX        =   27649
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
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   14520
      TabIndex        =   4
      Top             =   7080
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
      Picture         =   "frmcaixaConsulta.frx":2E5D
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
      Left            =   13320
      TabIndex        =   5
      Top             =   7080
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
      Picture         =   "frmcaixaConsulta.frx":2F67
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdConsultar 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Em aberto"
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
      Picture         =   "frmcaixaConsulta.frx":3079
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComCtl2.MonthView txtDataF 
      Height          =   2370
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   107872257
      CurrentDate     =   41801
   End
   Begin MSComCtl2.MonthView txtDataI 
      Height          =   2370
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   107872257
      CurrentDate     =   41801
   End
   Begin Vendas.VistaButton cmdConsultar2 
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Pagas "
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
      Picture         =   "frmcaixaConsulta.frx":3095
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdPagamento 
      Height          =   615
      Left            =   12000
      TabIndex        =   23
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "frmcaixaConsulta.frx":30B1
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label lblCan 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Canceladas (R$)"
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
      TabIndex        =   71
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lblCanceladas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   70
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label lblTotalCaixa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   69
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Total Caixa (R$)"
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
      Left            =   5640
      TabIndex        =   68
      Top             =   6960
      Width           =   4095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Boleto (R$)"
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
      Left            =   13680
      TabIndex        =   66
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblBoleto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   65
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblDescontoG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   64
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblCaixa 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5640
      TabIndex        =   41
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Desconto (R$)"
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
      Left            =   11640
      TabIndex        =   21
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblCartao 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   20
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   " Cartão (R$)"
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
      Left            =   9720
      TabIndex        =   19
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   " Dinheiro (R$)"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "A Prazo (R$)"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblDinheiro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   16
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblPrazo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblCadastro 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   12720
      TabIndex        =   13
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblConsulta 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado Consulta"
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
      Left            =   5640
      TabIndex        =   12
      Top             =   480
      Width           =   7095
   End
   Begin VB.Label lbldataF 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3780
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblDataI 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consulta"
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
      TabIndex        =   9
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7335
      Left            =   120
      Picture         =   "frmcaixaConsulta.frx":3403
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5400
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   -360
      Picture         =   "frmcaixaConsulta.frx":59BE
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   16440
   End
End
Attribute VB_Name = "frmcaixaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Scampo As String
Dim campo As String
Dim ChaveM As String
Dim Sql As String
Dim SQsort As String
Dim sqlwhere As String
Dim Sqlconsulta As String

Private Sub btnPagamento_Click()
    On Error GoTo trata_erro
    Dim mDinheiro As Double
    Dim mDesconto As Double
    Dim mCartao As Double
    Dim mBoleto As Double
    Dim mTotal As Double
    Dim vmTotal As Double
    Dim mTotalG As Double

    mDinheiro = lblDinheiroC.Caption
    mDesconto = lblDesconto.Caption
    mCartao = lblCartaoC.Caption
    mBoleto = lblBoletoC.Caption

    mTotalG = lblTotaldaVenda.Caption

    mDesconto = mTotalG - (mDinheiro + mCartao + mBoleto)

    vmTotal = (mDinheiro + mCartao + mBoleto) + mDesconto

    If mDinheiro + mCartao + mBoleto > vmTotal Then
        MsgBox ("Favor corrigir a forma de pagamento. valores incorretos"), vbInformation
        Exit Sub
    End If

    If mDinheiro + mCartao + mBoleto = 0 Then
        MsgBox ("Favor corrigir a forma de pagamento. valores incorretos"), vbInformation
        Exit Sub
    End If

    If FormatValor(vmTotal, 1) <> FormatValor(mTotalG, 1) Then
        MsgBox ("favor corrigir a forma de pagamento. valores incorretos"), vbInformation
        lblDesconto.Caption = "0,00"
        lblDinheiroC.Caption = "0,00"
        lblCartaoC.Caption = "0,00"
        lblBoletoC.Caption = "0,00"
        txtDinheiro.SetFocus
        btnSairPagamento_Click
        Exit Sub
    End If

    'If lblDinheiroC.caption = "" Then mDinheiro = 0 Else mDinheiro = lblDinheiroC.caption
    ' If mDesconto = 0 Then Else mDesconto = lblDesconto.Caption
    ' If lblCartaoC.caption = "" Then mCartao = 0 Else mCartao = lblCartaoC.caption
    ' If txtDinheiroD.text = "" Then txtDinheiroD.text = 0

    mTotal = lblTotal.Caption

    Sqlconsulta = "id_venda = '" & txtid_venda.text & "'"
    campo = "status = 'P'"
    campo = campo & ", valorcaixadinheiro ='" & FormatValor(mDinheiro, 1) & "'"
    campo = campo & ", valorcaixacartao = '" & FormatValor(mCartao, 1) & "'"
    campo = campo & ", valorcaixaDesconto = '" & FormatValor(mDesconto, 1) & "'"
    campó = campo & ", valorcaixaboleto = '" & FormatValor(mBoleto, 1) & "'"
    sqlAlterar "Caixa", campo, Sqlconsulta, Me, "N"

    btnAbilitado
    frPagamento.Visible = False
    frameFormaPagameto.Visible = False
    cmdConsultar_Click

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub btnSairPagamento_Click()
    btnAbilitado
    frPagamento.Visible = False
    frameFormaPagameto.Visible = False
    cmdConsultar.SetFocus
End Sub



Private Sub cmdExcluir_Click()

    txtSenha.text = ""
    frSenha.Visible = True
    txtSenha.SetFocus
    '  verifica_senha

End Sub

Private Sub cmdExtrato_Click()
    ClienteNome = lblCliente.Caption
    VendasExtrato
End Sub

Private Sub cmdPagamento_Click()
    frPagamento.Visible = True

    Visualizar_Pagamento
End Sub

Private Sub cmdRelatorios_Click()
    If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
        With rptCaixa
            .lblConsulta.Caption = Sqlconsulta
            .lblCliente.Caption = "Período: " & Format(txtDataI.Value, "DD/MM/YYYY") & " a " & Format(txtDataF.Value, "DD/MM/YYYY")
            .Show 1
        End With
    Else
        MsgBox ("Favor selecionar um periodo..."), vbInformation
        Exit Sub
    End If
End Sub

Private Sub cmdsairPG_Click()
    frPagamento.Visible = False
    cmdConsultar.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 15765
    Me.Height = 8625
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    txtDataI.Value = Now
    txtDataF.Value = Now

    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmcaixaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcaixaConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim caixa As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mDinheiro As Double
    Dim mPrazo As Double
    Dim mCartao As Double
    Dim mBoleto As Double
    Dim mDesconto As Double
    Dim mCanceladas As Double

    ' conecta ao banco de dados
    Set caixa = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = " Caixa.datacaixa"

    Sql = " SELECT caixa.*, vendedores.id_vendedor, vendedores.vendedor "
    Sql = Sql & " From"
    Sql = Sql & " caixa"
    Sql = Sql & " LEFT JOIN vendedores ON caixa.id_vendedor = vendedores.id_vendedor"
    Sql = Sql & " Where "
    Sql = Sql & SQconsulta
    Sql = Sql & " order by " & SQsort

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Caixa
    If caixa.State = 1 Then caixa.Close
    caixa.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaCaixa.ColumnHeaders.Clear
    ListaCaixa.ListItems.Clear

    If caixa.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation

        lblDinheiro.Caption = "0,00"
        lblPrazo.Caption = "0,00"
        lblCartao.Caption = "0,00"
        lblDescontoG.Caption = "0,00"
        lblBoleto.Caption = "0,00"
        lblTotalCaixa.Caption = "0,00"
        lblCanceladas.Caption = "0,00"

        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Caixa encontrado(s): " & caixa.RecordCount

    ListaCaixa.ColumnHeaders.Add , , "Venda", 1300
    ListaCaixa.ColumnHeaders.Add , , "Historico", 4500
    ListaCaixa.ColumnHeaders.Add , , "Data", 1700, lvwColumnCenter
    ListaCaixa.ColumnHeaders.Add , , "Dinheiro", 1500, lvwColumnRight
    ListaCaixa.ColumnHeaders.Add , , "Prazo", 1500, lvwColumnRight
    ListaCaixa.ColumnHeaders.Add , , "Cartão", 1500, lvwColumnRight
    ListaCaixa.ColumnHeaders.Add , , "Desconto", 1500, lvwColumnRight
    ListaCaixa.ColumnHeaders.Add , , "Boleto", 1500, lvwColumnRight
    ListaCaixa.ColumnHeaders.Add , , "Vendedor", 1400

    mDinheiro = 0
    mPrazo = 0
    mCanceladas = 0
    mBoleto = 0
    mDesconto = 0

    If caixa.BOF = True And caixa.EOF = True Then Exit Sub
    While Not caixa.EOF

        If VarType(caixa("id_venda")) <> vbNull Then
            Set itemx = ListaCaixa.ListItems.Add(, , caixa("id_venda"))
        Else
            Set itemx = ListaCaixa.ListItems.Add(, , "000000")
        End If
        If VarType(caixa("historico")) <> vbNull Then itemx.SubItems(1) = caixa("historico") Else itemx.SubItems(1) = ""
        If VarType(caixa("dataCaixa")) <> vbNull Then itemx.SubItems(2) = Format(caixa("dataCaixa"), "DD/MM/YYYY") Else itemx.SubItems(2) = ""

        If VarType(caixa("valorcaixadinheiro")) <> vbNull Then
            itemx.SubItems(3) = Format(caixa("valorcaixadinheiro"), "###,##0.00")
            If caixa("status") <> "C" Then
                mDinheiro = mDinheiro + caixa("valorcaixadinheiro")
            Else
                mCanceladas = mCanceladas + caixa("valorcaixadinheiro")
            End If
        Else
            itemx.SubItems(3) = ""
        End If

        If VarType(caixa("valorcaixaprazo")) <> vbNull Then
            itemx.SubItems(4) = Format(caixa("valorcaixaprazo"), "###,##0.00")
            If caixa("status") = "C" Then
                mCanceladas = mCanceladas + caixa("valorcaixaprazo")
            Else
                mPrazo = mPrazo + caixa("valorcaixaprazo")
            End If
        Else
            itemx.SubItems(4) = ""
        End If

        If VarType(caixa("valorcaixacartao")) <> vbNull Then
            itemx.SubItems(5) = Format(caixa("valorcaixacartao"), "###,##0.00")
            If caixa("status") = "C" Then
                mCanceladas = mCanceladas + caixa("valorcaixacartao")
            Else
                mCartao = mCartao + caixa("valorcaixacartao")
            End If
        Else
            itemx.SubItems(5) = ""
        End If

        If VarType(caixa("ValorCaixaDesconto")) <> vbNull Then
            If caixa("status") <> "C" Then
                itemx.SubItems(6) = Format(caixa("ValorCaixaDesconto"), "###,##0.00")
                mDesconto = mDesconto + caixa("ValorCaixaDesconto")
            End If
        Else
            itemx.SubItems(6) = ""
        End If

        If VarType(caixa("valorcaixaboleto")) <> vbNull Then
            itemx.SubItems(7) = Format(caixa("valorcaixaboleto"), "###,##0.00")
            If caixa("status") = "C" Then
                mCanceladas = mCanceladas + caixa("valorcaixaboleto")
            Else
                mBoleto = mBoleto + caixa("valorcaixaboleto")
            End If
        Else
            itemx.SubItems(7) = ""
        End If

        If VarType(caixa("vendedor")) <> vbNull Then itemx.SubItems(8) = caixa("vendedor") Else itemx.SubItems(8) = ""
        If VarType(caixa("id_caixa")) <> vbNull Then itemx.Tag = caixa("id_caixa")

        caixa.MoveNext
    Wend

    lblDinheiro.Caption = Format(mDinheiro, "###,##0.00")
    lblPrazo.Caption = Format(mPrazo, "###,##0.00")
    lblCartao.Caption = Format(mCartao, "###,##0.00")
    lblDescontoG.Caption = Format(mDesconto, "###,##0.00")
    lblBoleto.Caption = Format(mBoleto, "###,##0.00")
    lblCanceladas.Caption = Format(mCanceladas, "###,##0.00")

    lblTotalCaixa.Caption = Format((mDinheiro + mPrazo + mCartao + mBoleto) - mDesconto, "###,##0.00")




    'Zebra o listview
    If LVZebra(ListaCaixa, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If caixa.State = 1 Then caixa.Close
    Set caixa = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaCaixa_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim caixa As ADODB.Recordset
    Set caixa = CreateObject("ADODB.Recordset")

    txtid_caixa.text = ListaCaixa.SelectedItem.Tag

    If txtid_caixa.text <> "" Then
        Sql = "select caixa.status"
        Sql = Sql & " from "
        Sql = Sql & " caixa"
        Sql = Sql & " where "
        Sql = Sql & " caixa.id_caixa = '" & txtid_caixa.text & "'"
        Sql = Sql & " and caixa.status = 'A'"
        If caixa.State = 1 Then caixa.Close
        caixa.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If caixa.RecordCount > 0 Then
            cmdPagamento.Enabled = True
        Else
            cmdPagamento.Enabled = False
        End If
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaCaixa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Select Case ColumnHeader
    Case Is = "Historico"
        SQsort = "historico"
    Case Is = "Data"
        SQsort = "dataCaixa"
    Case Is = "Dinherio"
        SQsort = "valorcaixadinheiro"
    Case Is = "Prazo"
        SQsort = "valorcaixaprazo"
    Case Is = "Cartão"
        SQsort = "valorcaixacartao"
    Case Is = "Vendedor"
        SQsort = "vendedor"
    End Select

    cmdConsultar_Click

End Sub

Private Sub cmdConsultar_Click()

    lblCaixa.Caption = " Contas em Aberto no Caixa "

    lblPrazo.Caption = "0,00"
    lblDinheiro.Caption = "0,00"
    lblCartao.Caption = "0,00"
    lblDesconto.Caption = "0,00"
    lblBoleto.Caption = "0,00"

    Sqlconsulta = " caixa.status = 'A' "

    Lista (Sqlconsulta)

End Sub

Private Sub cmdConsultar2_Click()

    lblCaixa.Caption = " Contas Fechadas no Caixa "

    lblPrazo.Caption = "0,00"
    lblDinheiro.Caption = "0,00"
    lblCartao.Caption = "0,00"
    lblDesconto.Caption = "0,00"
    lblBoleto.Caption = "0,00"

    Sqlconsulta = " caixa.status = 'P' "
    Sqlconsulta = Sqlconsulta & " and caixa.datacaixa Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
    Sqlconsulta = Sqlconsulta & " or "
    Sqlconsulta = Sqlconsulta & " caixa.status = 'C' "
    Sqlconsulta = Sqlconsulta & " and caixa.datacaixa Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"

    Lista (Sqlconsulta)
End Sub


Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub


Private Sub Visualizar_Pagamento()
    On Error GoTo trata_erro
    Dim caixa As ADODB.Recordset
    Set caixa = CreateObject("ADODB.Recordset")

    If txtid_caixa.text <> "" Then
        Sql = "select caixa.*, vendedores.id_vendedor, vendedores.vendedor,"
        Sql = Sql & " prazo.id_prazo, prazo.id_cliente, prazopagto.id_prazo,"
        Sql = Sql & " clientes.id_cliente , clientes.Cliente"
        Sql = Sql & " from "
        Sql = Sql & " caixa"
        Sql = Sql & " left join vendedores on caixa.id_vendedor = vendedores.id_vendedor"
        Sql = Sql & " left join prazopagto on caixa.id_prazoPagto = prazopagto.id_prazoPagto"
        Sql = Sql & " left join prazo on prazopagto.id_prazo = prazo.id_prazo"
        Sql = Sql & " left join clientes on prazo.id_cliente = clientes.id_cliente"
        Sql = Sql & " where "
        Sql = Sql & " caixa.id_caixa = '" & txtid_caixa.text & "'"

        If caixa.State = 1 Then caixa.Close
        caixa.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If caixa.RecordCount > 0 Then
            If VarType(caixa("id_venda")) <> vbNull Then lblid_venda.Caption = caixa("id_venda") Else lblid_venda.Caption = ""
            If VarType(caixa("id_venda")) <> vbNull Then txtid_venda.text = caixa("id_venda") Else txtid_venda.text = ""
            If VarType(caixa("datacaixa")) <> vbNull Then lbldata_caixa.Caption = caixa("datacaixa") Else lbldata_caixa.Caption = ""
            If VarType(caixa("vendedor")) <> vbNull Then lblVendedor.Caption = caixa("vendedor") Else lblVendedor.Caption = ""
            If VarType(caixa("historico")) <> vbNull Then lblHistorico.Caption = caixa("historico") Else lblHistorico.Caption = ""
            If VarType(caixa("cliente")) <> vbNull Then lblCliente.Caption = caixa("cliente") Else lblCliente.Caption = ""
            If VarType(caixa("valorcaixadinheiro")) <> vbNull Then lblDinheiroC.Caption = Format(caixa("valorcaixadinheiro"), "###,##0.00") Else lblDinheiroC.Caption = "0,00"
            If VarType(caixa("valorcaixacartao")) <> vbNull Then lblCartaoC.Caption = Format(caixa("valorcaixacartao"), "###,##0.00") Else lblCartaoC.Caption = "0,00"
            If VarType(caixa("valorcaixadesconto")) <> vbNull Then lblDesconto.Caption = Format(caixa("valorcaixadesconto"), "###,##0.00") Else lblDesconto.Caption = "0,00"
            If VarType(caixa("valorcaixaboleto")) <> vbNull Then lblBoletoC.Caption = Format(caixa("valorcaixaboleto"), "###,##0.00") Else lblBoletoC.Caption = "0,00"
            txtDinheiroD.text = lblDinheiroC.Caption
            ListaItens ("")
        End If
    End If

    If caixa.State = 1 Then caixa.Close
    Set caixa = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub ListaItens(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Saida As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mTotal As Double
    Dim mDesconto As Double

    ' conecta ao banco de dados
    Set Saida = CreateObject("ADODB.Recordset")

    Sql = " select saida.id_estoque, saida.id_venda, saida.quantidade,"
    Sql = Sql & " saida.preco_venda, estoques.id_estoque, estoques.descricao,"
    Sql = Sql & " saida.quantidade * saida.preco_venda as totalIten"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " left join estoques on saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where "
    Sql = Sql & " saida.id_venda = '" & txtid_venda.text & "'"

    ' abre um Recrodset da Tabela Saida
    If Saida.State = 1 Then Saida.Close
    Saida.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaSaida.ColumnHeaders.Clear
    ListaSaida.ListItems.Clear

    If Saida.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If

    ListaSaida.ColumnHeaders.Add , , "Descrição", 5000
    ListaSaida.ColumnHeaders.Add , , "Quantidade", 1500, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Preço", 1500, lvwColumnRight
    ListaSaida.ColumnHeaders.Add , , "Total", 1500, lvwColumnRight

    mTotal = 0

    If Saida.BOF = True And Saida.EOF = True Then Exit Sub
    While Not Saida.EOF
        If VarType(Saida("descricao")) <> vbNull Then Set itemx = ListaSaida.ListItems.Add(, , Saida("descricao"))
        If VarType(Saida("quantidade")) <> vbNull Then itemx.SubItems(1) = Format(Saida("quantidade"), "##,##0.000") Else itemx.SubItems(1) = ""
        If VarType(Saida("preco_venda")) <> vbNull Then itemx.SubItems(2) = Format(Saida("preco_venda"), "##,##0.00") Else itemx.SubItems(2) = ""
        If VarType(Saida("totalIten")) <> vbNull Then itemx.SubItems(3) = Format(Saida("totalIten"), "##,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Saida("id_venda")) <> vbNull Then itemx.Tag = Saida("id_venda")

        mTotal = mTotal + Saida("totaliten")

        Saida.MoveNext
    Wend

    If lblDesconto.Caption <> "" Then mDesconto = lblDesconto.Caption Else mDesconto = 0

    txtTotalVenda.text = mTotal

    mTotal = mTotal - mDesconto

    lblTotal.Caption = Format(txtTotalVenda.text, "##,##0.00")
    lblTotaldaVenda.Caption = Format(txtTotalVenda.text, "##,##0.00")


    ' txtDinheiro.SetFocus

    'Zebra o listview
    If LVZebra(ListaSaida, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Saida.State = 1 Then Saida.Close
    Set Saida = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


'-------- txtDinheiro
Private Sub txtDinheiro_GotFocus()
'    txtDinheiro.BackColor = &H80FFFF
End Sub
Private Sub txtDinheiro_LostFocus()
    On Error GoTo trata_erro
    Dim mDinheiro As Double
    Dim mDesconto As Double
    Dim mCartao As Double
    Dim mTotal As Double
    Dim mtroco As Double

    If IsEmpty(lblDinheiroC.Caption) Then lblDinheiroC.Caption = "0,00"

    txtDinheiro.BackColor = &H80000014
   ' lblDinheiroC.Caption = Format(lblDinheiroC.Caption, "###,##0.00")
    lblDinheiroC.Caption = Format(txtDinheiro.text, "###,##0.00")

    mTotal = lblTotaldaVenda.Caption

    lblDesconto.Caption = "0,00"

    mDinheiro = lblDinheiroC.Caption
    mCartao = lblCartaoC.Caption

    If mDinheiro + mCartao < mTotal Then
        mDesconto = mTotal - (mDinheiro + mCartao)
        lblDesconto.Caption = Format(mDesconto, "###,##0.00")
    End If

    If mDinheiro = 0 Then lblDinheiroC.Caption = "0,00"

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub txtDinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCartao.SetFocus
End Sub

'-------- txtCartao
Private Sub txtCartao_GotFocus()
    txtCartao.BackColor = &H80FFFF
End Sub
Private Sub txtCartao_LostFocus()
    On Error GoTo trata_erro
    Dim mDinheiro As Double
    Dim mDesconto As Double
    Dim mCartao As Double
    Dim mTotal As Double
    Dim mtroco As Double


    txtCartao.BackColor = &H80000014
    lblCartaoC.Caption = Format(lblCartaoC.Caption, "###,##0.00")

    mTotal = lblTotaldaVenda.Caption

    mDinheiro = lblDinheiroC.Caption
    mCartao = lblCartaoC.Caption

    If mCartao = 0 Then lblCartaoC.Caption = "0,00"

    If mDinheiro + mCartao < mTotal Then
        mDesconto = mTotal - (mDinheiro + mCartao)
        lblDesconto.Caption = Format(mDesconto, "###,##0.00")
    End If

    If mDinheiro + mCartao = mTotal Then
        mDesconto = 0
        lblDesconto.Caption = Format(mDesconto, "###,##0.00")
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub txtCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtBoleto.SetFocus
End Sub


'-------- txtboleto
Private Sub txtboleto_GotFocus()
    txtBoleto.BackColor = &H80FFFF
End Sub
Private Sub txtboleto_LostFocus()
    On Error GoTo trata_erro
    Dim mDinheiro As Double
    Dim mCartao As Double
    Dim mDesconto As Double
    Dim mBoleto As Double
    Dim mTotal As Double
    Dim mtroco As Double


    txtBoleto.BackColor = &H80000014
    lblBoletoC.Caption = Format(lblBoletoC.Caption, "###,##0.00")

    mTotal = lblTotaldaVenda.Caption

    mDinheiro = lblDinheiroC.Caption
    mCartao = lblCartaoC.Caption
    mBoleto = lblBoletoC.Caption

    If mBoleto = 0 Then lblBoletoC.Caption = "0,00"

    If mDinheiro + mBoleto + mCartao < mTotal Then
        mDesconto = mTotal - (mDinheiro + mBoleto + mCartao)
        lblDesconto.Caption = Format(mDesconto, "###,##0.00")
    End If

    If mDinheiro + mBoleto + mCartao = mTotal Then
        mDesconto = 0
        lblDesconto.Caption = Format(mDesconto, "###,##0.00")
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub txtboleto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then btnPagamento.SetFocus
End Sub


Private Sub cmdAvista_Click()
    frameFormaPagameto.Visible = True
    btnDesabilitado
   ' txtDinheiro_GotFocus
   ' txtDinheiro.SetFocus
    ' cmdConsultar.SetFocus
    cmdPagamento.SetFocus
End Sub


Private Sub btnAbilitado()
    cmdAvista.Enabled = True
    cmdExtrato.Enabled = True
    cmdExcluir.Enabled = True
    cmdsairPG.Enabled = True
End Sub

Private Sub btnDesabilitado()
    cmdAvista.Enabled = False
    cmdExtrato.Enabled = False
    cmdExcluir.Enabled = False
    cmdsairPG.Enabled = False
End Sub


'-------- txtCalculoTroco
Private Sub txtCalculoTroco_GotFocus()
    txtCalculoTroco.BackColor = &H80FFFF
End Sub
Private Sub txtCalculoTroco_LostFocus()
    On Error GoTo trata_erro
    Dim mDinheiro As Double
    Dim mDesconto As Double
    Dim mCartao As Double
    Dim mTotal As Double
    Dim mtroco As Double

    txtCalculoTroco.BackColor = &H80000014
    txtCalculoTroco.text = Format(txtCalculoTroco.text, "###,##0.00")

    If lblDesconto.Caption = "" Then mDesconto = 0 Else mDesconto = lblDesconto.Caption
    If txtCalculoTroco.text = "" Then mDinheiro = 0 Else mDinheiro = txtCalculoTroco.text
    If mDesconto = 0 Then lblDesconto.Caption = "0,00"
    If mDinheiro = 0 Then txtCalculoTroco.text = "0,00"

    mTotal = lblTotal.Caption

    mtroco = Format(mTotal - mDinheiro, "###,##0.00")
    If mtroco < 0 Then mtroco = Format((mDinheiro) - mTotal, "###,##0.00")
    If (mDesconto + mDinheiro) <= mTotal Then mtroco = 0


    lblTroco.Caption = Format(mtroco, "###,##0.00")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub txtCalculoTroco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then btnPagamento.SetFocus
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
    Dim strCliente As String


    Dim mDebito As Double
    Dim mCredito As Double
    Dim mAPagar As Double
    Dim mDaDos As Integer
    Dim mCabecarioDados As String
    Dim mArquivo As String

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

    txtExtrato.text = txtExtrato.text & "BOTELHO MATERIAL DE CONSTRUCAO " & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "AV MURIAE, 376 - NOVA CARAPINA - TEL. 3341-2005 / 99511-3123" & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "dd/mm/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "" & vbCrLf
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


    txtExtrato.text = txtExtrato.text & "BOTELHO MATERIAL DE CONSTRUCAO " & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "AV MURIAE, 376 - NOVA CARAPINA - TEL. 3341-2005 / 99511-3123" & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "mm/dd/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "........: " & vbCrLf & vbCrLf & vbCrLf

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf

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

Private Sub VendasCancelar()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim Tabela1 As ADODB.Recordset

    Dim mSaldo As Double
    Dim mQuantidade As Double
    Dim mControlarSaldo As String
    Dim mDesconto As Double
    Dim mDinheiro As Double
    Dim mCartao As Double
    Dim mBoleto As Double

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

        If txtTotalVenda.text = "" Then txtTotalVenda.text = 0
        If txtDinheiroD.text = "" Then txtDinheiroD.text = 0
        If lblCartaoC.Caption = "" Then lblCartaoC.Caption = 0
        If lblBoletoC.Caption = "" Then lblBoletoC.Caption = 0
        If lblDesconto.Caption = "" Then lblDesconto.Caption = 0

        mDinheiro = lblDinheiroC.Caption
        mCartao = lblCartaoC.Caption
        mBoleto = lblBoletoC.Caption

        mDesconto = 0

        mDesconto = mDinheiro + mCartao + mBoleto

        Sqlconsulta = "id_venda = '" & lblid_venda.Caption & "'"
        campo = "status = 'C'"
        campo = campo & ", historico = 'Venda Cancelada'"
        '  If txtDinheiroD.text <> 0 Then campo = campo & ", valorcaixadinheiro ='" & FormatValor(txtDinheiroD.text, 1) & "'"
        '  If lblCartaoC.caption <> 0 Then campo = campo & ", valorcaixacartao = '" & FormatValor(lblCartaoC.caption, 1) & "'"
        '  If lblBoletoC.caption <> 0 Then campo = campo & ", valorcaixaboleto = '" & FormatValor(lblBoletoC.caption, 1) & "'"
        '  campo = campo & ", valorcaixaDesconto = '" & FormatValor(mDesconto, 1) & "'"
        sqlAlterar "Caixa", campo, Sqlconsulta, Me, "N"


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

    Sql = "Select vendedores.tipo_acesso "
    Sql = Sql & " from "
    Sql = Sql & " vendedores"
    Sql = Sql & " where vendedores.acesso = '" & txtSenha.text & "'"

    ' abre um Recrodset da Tabela Tabela
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("tipo_acesso")) <> vbNull Then mTipoAcesso = Tabela("tipo_acesso") Else mTipoAcesso = ""
    End If

    If mTipoAcesso = "A" Then
        VendasCancelar
    Else
        SemAcesso
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub SemAcesso()
    MsgBox ("Você não tem autorização para este tipo de acesso.."), vbInformation
End Sub


