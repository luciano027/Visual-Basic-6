VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaixaDevolucaoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa Devolução Consulta"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6240
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_entrada 
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_fornecedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.ListView ListaEntradas 
      Height          =   6615
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11668
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
      TabIndex        =   5
      Top             =   8310
      Width           =   13065
      _ExtentX        =   23045
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
      Left            =   13680
      TabIndex        =   6
      Top             =   8160
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
      Picture         =   "frmCaixaDevolucaoConsulta.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdAlterar 
      Height          =   615
      Left            =   12480
      TabIndex        =   7
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Alterar"
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
      Picture         =   "frmCaixaDevolucaoConsulta.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdIncluir 
      Height          =   615
      Left            =   11280
      TabIndex        =   8
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Incluir"
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
      Picture         =   "frmCaixaDevolucaoConsulta.frx":045C
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdConsultar 
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      Caption         =   "Consultar"
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
      Picture         =   "frmCaixaDevolucaoConsulta.frx":07AE
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
      Left            =   2760
      TabIndex        =   13
      Top             =   720
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108789761
      CurrentDate     =   41801
   End
   Begin MSComCtl2.MonthView txtDataI 
      Height          =   2370
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108789761
      CurrentDate     =   41801
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   480
      Width           =   1935
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
      Top             =   120
      Width           =   6015
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
      Left            =   11640
      TabIndex        =   11
      Top             =   120
      Width           =   3135
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
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7425
      Left            =   120
      Picture         =   "frmCaixaDevolucaoConsulta.frx":07CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmCaixaDevolucaoConsulta.frx":25BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15120
   End
End
Attribute VB_Name = "frmCaixaDevolucaoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
