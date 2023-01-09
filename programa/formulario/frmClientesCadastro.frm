VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientesCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Frame3"
      Height          =   2250
      Left            =   10920
      TabIndex        =   47
      Top             =   1320
      Width           =   1935
      Begin VB.TextBox txtLimite 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   50
         Text            =   "0,00"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkSim 
         Caption         =   "Sim"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Limite"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Venda a Prazo "
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
         TabIndex        =   48
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   120
      TabIndex        =   44
      Top             =   3720
      Width           =   9135
      Begin RichTextLib.RichTextBox txtObservacao 
         Height          =   975
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   1720
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmClientesCadastro.frx":0000
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observação"
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
         TabIndex        =   45
         Top             =   0
         Width           =   9135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1050
      Left            =   10920
      TabIndex        =   40
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativo"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
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
         TabIndex        =   43
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.TextBox txtChave 
      Height          =   285
      Left            =   8160
      TabIndex        =   26
      Text            =   "0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   285
      Left            =   4080
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_login 
      Height          =   285
      Left            =   5760
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   4920
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6840
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   9360
      TabIndex        =   14
      Top             =   3720
      Width           =   3495
      Begin VB.TextBox txtCnpj 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtContato 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtInscricao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cart.Ident."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Documento Identificação"
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
         TabIndex        =   19
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C.P.F."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10695
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   1080
         Width           =   7815
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtuf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endereço"
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
         TabIndex        =   7
         Top             =   0
         Width           =   10695
      End
      Begin VB.Image cmdConsultaCEP 
         Height          =   240
         Left            =   1320
         Picture         =   "frmClientesCadastro.frx":0082
         ToolTipText     =   "Verificar CEP Digitado"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image cmdConsulCEP 
         Height          =   360
         Left            =   1560
         Picture         =   "frmClientesCadastro.frx":03C4
         Stretch         =   -1  'True
         ToolTipText     =   "Consulta CEP "
         Top             =   480
         Width           =   360
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3201
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Identificação"
      TabPicture(0)   =   "frmClientesCadastro.frx":06CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtTel2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFax"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtemail"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCliente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Top             =   720
         Width           =   8415
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6840
         TabIndex        =   30
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         TabIndex        =   29
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtTel2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   1080
         Left            =   240
         Picture         =   "frmClientesCadastro.frx":06EA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   34
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone - FAX"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   1080
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   6165
      Width           =   13065
      _ExtentX        =   23045
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
      Left            =   11400
      TabIndex        =   37
      Top             =   5400
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
      Picture         =   "frmClientesCadastro.frx":0CD3
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
      Left            =   9720
      TabIndex        =   38
      Top             =   5400
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
      Picture         =   "frmClientesCadastro.frx":0DDD
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
      Left            =   8160
      TabIndex        =   39
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Gravar"
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
      Picture         =   "frmClientesCadastro.frx":132F
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmClientesCadastro.frx":1881
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   13035
   End
End
Attribute VB_Name = "frmClientesCadastro"
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
    txtCliente.SetFocus
    If txtTipo.text = "A" Or txtTipo.text = "E" Or txtTipo.text = "L" Then AutalizaCadastro

End Sub

Private Sub Form_Load()
    Me.Width = 13155
    Me.Height = 6915
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmClientesCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmClientesCadastro = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro




    Unload Me
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Dim table As String

    ' Rotina de gravacao de inclusao dos dados
    If txtTipo.text = "I" Then

        campo = " data_cadastro"
        Scampo = "'" & Format(Date$, "YYYYMMDD") & "'"

        If txtCliente.text <> "" Then
            campo = campo & ", cliente"
            Scampo = Scampo & ",'" & txtCliente.text & "'"
        Else
            MsgBox ("Nome do clientes não pode ficar em branco..")
            txtCliente.SetFocus
            Exit Sub
        End If
        If txtTel2.text <> "" Then
            campo = campo & ", tel2"
            Scampo = Scampo & ", '" & txtTel2.text & "'"
        End If
        If txtFax.text <> "" Then
            campo = campo & ", fax"
            Scampo = Scampo & ", '" & txtFax.text & "'"
        End If
        If txtemail.text <> "" Then
            campo = campo & ", email"
            Scampo = Scampo & ", '" & txtemail.text & "'"
        End If

        If txtContato.text <> "" Then
            campo = campo & ", Contato"
            Scampo = Scampo & ", '" & txtContato.text & "'"
        End If

        If txtInscricao.text <> "" Then
            campo = campo & ", Inscricao"
            Scampo = Scampo & ", '" & txtInscricao.text & "'"
        End If

        If txtCnpj.text <> "" Then
            campo = campo & ", cnpj"
            Scampo = Scampo & ", '" & txtCnpj.text & "'"
        End If

        If txtCep.text <> "" Then
            campo = campo & ", cep"
            Scampo = Scampo & ", '" & txtCep.text & "'"
        End If
        If txtuf.text <> "" Then
            campo = campo & ", uf"
            Scampo = Scampo & ", '" & txtuf.text & "'"
        End If

        If txtNumero.text <> "" Then
            campo = campo & ", numero"
            Scampo = Scampo & ", '" & txtNumero.text & "'"
        End If

        If txtBairro.text <> "" Then
            campo = campo & ", bairro"
            Scampo = Scampo & ", '" & txtBairro.text & "'"
        End If
        If txtCidade.text <> "" Then
            campo = campo & ", cidade"
            Scampo = Scampo & ", '" & txtCidade.text & "'"
        End If


        If txtRua.text <> "" Then
            campo = campo & ", rua"
            Scampo = Scampo & ", '" & txtRua.text & "'"
        End If

        If optAtivo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'A'"
        End If

        If optInativo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'I'"
        End If

        If txtObservacao.text <> "" Then
            campo = campo & ", observacao"
            Scampo = Scampo & ", '" & txtObservacao.text & "'"
        End If

        If chkSim.Value = 1 Then
            campo = campo & ", prazo"
            Scampo = Scampo & ", 'S'"
        Else
            campo = campo & ", prazo"
            Scampo = Scampo & ", 'N'"
        End If

        If txtLimite.text <> "" Then
            campo = campo & ", limite"
            Scampo = Scampo & ", '" & FormatValor(txtLimite.text, 1) & "'"
        End If


        ' Incluir valor na tabela clientes
        sqlIncluir "clientes", campo, Scampo, Me, "S"

    End If
    ' rotina de gravacao de alteracao dos dados
    If txtTipo.text = "A" Then

        ' Consulta os dados da tabela clientes
        Sqlconsulta = "id_cliente = '" & txtid_cliente.text & "'"

        If txtCliente.text <> "" Then campo = " cliente = '" & UCase(txtCliente.text) & "'" Else txtCliente.SetFocus
        If txtTel2.text <> "" Then campo = campo & ", tel2 = '" & txtTel2.text & "'"
        If txtFax.text <> "" Then campo = campo & ", fax = '" & txtFax.text & "'"
        If txtemail.text <> "" Then campo = campo & ", email = '" & txtemail.text & "'"
        If txtCep.text <> "" Then campo = campo & ", cep = '" & txtCep.text & "'"
        If txtuf.text <> "" Then campo = campo & ", uf = '" & txtuf.text & "'"
        If txtNumero.text <> "" Then campo = campo & ", numero = '" & txtNumero.text & "'"
        If txtBairro.text <> "" Then campo = campo & ", bairro = '" & txtBairro.text & "'"
        If txtCidade.text <> "" Then campo = campo & ", cidade = '" & txtCidade.text & "'"
        If txtRua.text <> "" Then campo = campo & ", Rua = '" & txtRua.text & "'"
        If txtContato.text <> "" Then campo = campo & ", Contato = '" & txtContato.text & "'"
        If txtInscricao.text <> "" Then campo = campo & ", Inscricao = '" & txtInscricao.text & "'"
        If txtCnpj.text <> "" Then campo = campo & ", cnpj = '" & txtCnpj.text & "'"
        If optAtivo.Value = True Then campo = campo & ", status = 'A'" Else campo = campo & ", status = 'I'"
        If txtObservacao.text <> "" Then campo = campo & ", observacao = '" & txtObservacao.text & "'"
        If chkSim.Value = 1 Then campo = campo & ", prazo = 'S'" Else campo = campo & ", prazo = 'N' "

        ' Aletar dos dados da tabela clientes
        sqlAlterar "clientes", campo, Sqlconsulta, Me, "S"

    End If

    If txtTipo.text = "L" Then
        Sqlconsulta = "id_cliente = '" & txtid_cliente.text & "'"
        If txtLimite.text <> "" Then campo = " limite = '" & FormatValor(txtLimite.text, 1) & "'"
        sqlAlterar "clientes", campo, Sqlconsulta, Me, "S"
    End If

    cmdGravar.Enabled = False

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub AutalizaCadastro()
    On Error GoTo trata_erro

    If txtTipo.text = "A" Or txtTipo.text = "E" Or txtTipo.text = "L" Then
        If txtChave.text = "0" Then
            Dim clientes As ADODB.Recordset
            ' conecta ao banco de dados

            Set clientes = CreateObject("ADODB.Recordset")    '''

            ' abre um Recrodset da Tabela clientes
            Sql = " select "
            Sql = Sql & " clientes.*"
            Sql = Sql & " from  "
            Sql = Sql & " clientes "
            Sql = Sql & " where "
            Sql = Sql & " id_cliente = '" & txtid_cliente.text & "'"

            If clientes.State = 1 Then clientes.Close
            clientes.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If clientes.RecordCount > 0 Then

                If VarType(clientes("cliente")) <> vbNull Then txtCliente.text = clientes("cliente") Else txtCliente.text = ""
                If VarType(clientes("tel2")) <> vbNull Then txtTel2.text = clientes("tel2") Else txtTel2.text = ""
                If VarType(clientes("fax")) <> vbNull Then txtFax.text = clientes("fax") Else txtFax.text = ""
                If VarType(clientes("email")) <> vbNull Then txtemail.text = clientes("email") Else txtemail.text = ""
                If VarType(clientes("Rua")) <> vbNull Then txtRua.text = clientes("Rua") Else txtRua.text = ""
                If VarType(clientes("bairro")) <> vbNull Then txtBairro.text = clientes("bairro") Else txtBairro.text = ""
                If VarType(clientes("cidade")) <> vbNull Then txtCidade.text = clientes("cidade") Else txtCidade.text = ""
                If VarType(clientes("cep")) <> vbNull Then txtCep.text = clientes("cep") Else txtCep.text = ""
                If VarType(clientes("uf")) <> vbNull Then txtuf.text = clientes("uf") Else txtuf.text = ""
                If VarType(clientes("numero")) <> vbNull Then txtNumero.text = clientes("numero") Else txtNumero.text = ""
                If VarType(clientes("Contato")) <> vbNull Then txtContato.text = clientes("Contato") Else txtContato.text = ""
                If VarType(clientes("cnpj")) <> vbNull Then txtCnpj.text = clientes("cnpj") Else txtCnpj.text = ""
                If VarType(clientes("Inscricao")) <> vbNull Then txtInscricao.text = clientes("Inscricao") Else txtInscricao.text = ""
                If VarType(clientes("observacao")) <> vbNull Then txtObservacao.text = clientes("observacao") Else txtObservacao.text = ""
                If VarType(clientes("status")) <> vbNull Then
                    If clientes("status") = "A" Then optAtivo.Value = True
                    If clientes("status") = "I" Then optInativo.Value = True
                End If
                If clientes("prazo") = "S" Then chkSim.Value = 1 Else chkSim.Value = 0
                If VarType(clientes("limite")) <> vbNull Then txtLimite.text = Format(clientes("limite"), "###,##0.00") Else txtLimite.text = "0,00"

            End If

            If txtLimite.text = "L" Then txtLimite.SetFocus

            If clientes.State = 1 Then clientes.Close
            Set clientes = Nothing

            If txtTipo.text = "E" Then cmdGravar.Enabled = False
            If txtTipo.text = "A" Then cmdExcluir.Enabled = False
            txtChave.text = "1"

        End If
    End If

    txtCliente.SetFocus
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Image4_Click()

End Sub

'---------------------------------------------------------------
'----------------------- campos do formulario-------------------------------

'------------ nome
Private Sub txtcliente_GotFocus()
    txtCliente.BackColor = &H80FFFF
End Sub
Private Sub txtcliente_LostFocus()
    txtCliente.BackColor = &H80000014
    If Len(txtCliente.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtCliente.text)
        txtCliente.SetFocus
    End If
End Sub
Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtTel2.SetFocus
End Sub

'--- tel2
Private Sub txttel2_GotFocus()
    txtTel2.BackColor = &H80FFFF
End Sub
Private Sub txttel2_LostFocus()
    txtTel2.BackColor = &H80000014
    txtTel2.text = SoNumero(txtTel2.text)
    txtTel2.text = FormataTelefone(txtTel2.text)
    If Len(txtTel2.text) > 16 Then
        MsgBox "Comprimento do campo e de 16 digitos, voce digitou " & Len(txtTel2.text)
        txtTel2.SetFocus
    End If
End Sub
Private Sub txttel2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFax.SetFocus
End Sub

'---fax
Private Sub txtfax_GotFocus()
    txtFax.BackColor = &H80FFFF
End Sub
Private Sub txtfax_LostFocus()
    txtFax.BackColor = &H80000014
    txtFax.text = SoNumero(txtFax.text)
    txtFax.text = FormataTelefone(txtFax.text)
    If Len(txtFax.text) > 16 Then
        MsgBox "Comprimento do campo e de 16 digitos, voce digitou " & Len(txtFax.text)
        txtFax.SetFocus
    End If
End Sub
Private Sub txtfax_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtemail.SetFocus
End Sub


'-------- email
Private Sub txtemail_GotFocus()
    txtemail.BackColor = &H80FFFF
End Sub
Private Sub txtemail_LostFocus()
    txtemail.BackColor = &H80000014
    If Len(txtemail.text) > 150 Then
        MsgBox "Comprimento do campo e de 150 digitos, voce digitou " & Len(txtemail.text)
        txtemail.SetFocus
    End If
End Sub
Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCep.SetFocus
End Sub


'-------- cep
Private Sub txtcep_GotFocus()
    txtCep.BackColor = &H80FFFF
    If CEPVer <> "" Then
        txtCep.text = CEPVer
        CEPVer = ""
    End If
End Sub
Private Sub txtcep_LostFocus()
    txtCep.BackColor = &H80000014
    txtCep.text = SoNumero(txtCep.text)
    If Len(txtCep.text) > 10 Then
        MsgBox "Comprimento do campo e de 10 digitos, voce digitou " & Len(txtCep.text)
        txtCep.SetFocus
    End If
End Sub
Private Sub txtcep_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtuf.SetFocus
End Sub
Private Sub cmdConsultaCEP_Click()
    On Error GoTo trata_erro
    Dim CEP As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim CEPs As String
    ' conecta ao banco de dados
    Set CEP = CreateObject("ADODB.Recordset")

    If txtCep.text <> "" Then
        CEPs = SoNumero(txtCep.text)
        CEPs = Mid(txtCep.text, 1, 5) & "-" & Mid(txtCep.text, 6, 3)
        Aguarde_Process Me, True
        Sql = "select cep_endereco.id_endereco, cep_endereco.id_bairro, cep_endereco.endereco_completo, cep_endereco.cep, " _
            & " cep_cidade.id_cidade, cep_cidade.id_estado, cep_cidade.cidade, " _
            & " cep_bairro.id_bairro, cep_bairro.id_cidade, cep_bairro.bairro, " _
            & " cep_estados.id_estado , cep_estados.uf " _
            & " From " _
            & " cep_endereco " _
            & " left join cep_cidade on cep_endereco.id_cidade = cep_cidade.id_cidade " _
            & " left join cep_bairro on cep_endereco.id_bairro = cep_bairro.id_bairro " _
            & " left join cep_estados on cep_cidade.id_estado = cep_estados.id_estado " _
            & " where cep = '" & CEPs & "'"

        ' abre um Recrodset da Tabela CEP
        If CEP.State = 1 Then CEP.Close
        CEP.Open Sql, banco, adOpenKeyset, adLockOptimistic
        Aguarde_Process Me, False
        If CEP.RecordCount > 0 Then
            If VarType(CEP("uf")) <> vbNull Then txtuf.text = CEP("uf") Else txtuf.text = ""
            If VarType(CEP("bairro")) <> vbNull Then txtBairro.text = CEP("bairro") Else txtBairro.text = ""
            If VarType(CEP("cidade")) <> vbNull Then txtCidade.text = CEP("cidade") Else txtCidade.text = ""
            If VarType(CEP("endereco_completo")) <> vbNull Then txtRua.text = CEP("endereco_completo") Else txtRua.text = ""
            txtCep.text = CEPs
        Else
            MsgBox ("CEP Não encontrado..."), vbExclamation
        End If
    End If

    If CEP.State = 1 Then CEP.Close
    Set CEP = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
Private Sub cmdConsulCEP_Click()
    With frmCepConsulta
        .Show 1
    End With
End Sub


'-------- uf
Private Sub txtUF_GotFocus()
    txtuf.BackColor = &H80FFFF
End Sub
Private Sub txtuf_LostFocus()
    txtuf.BackColor = &H80000014
    If Len(txtuf.text) > 2 Then
        MsgBox "Comprimento do campo e de 2 digitos, voce digitou " & Len(txtuf.text)
        txtuf.SetFocus
    End If
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNumero.SetFocus
End Sub

'-------- Numero
Private Sub txtNumero_GotFocus()
    txtNumero.BackColor = &H80FFFF
End Sub
Private Sub txtNumero_LostFocus()
    txtNumero.BackColor = &H80000014
    If Len(txtNumero.text) > 5 Then
        MsgBox "Comprimento do campo e de 5 digitos, voce digitou " & Len(txtNumero.text)
        txtNumero.SetFocus
    End If
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtBairro.SetFocus
End Sub
'-------- bairro
Private Sub txtbairro_GotFocus()
    txtBairro.BackColor = &H80FFFF
End Sub
Private Sub txtbairro_LostFocus()
    txtBairro.BackColor = &H80000014
    If Len(txtBairro.text) > 70 Then
        MsgBox "Comprimento do campo e de 70 digitos, voce digitou " & Len(txtBairro.text)
        txtBairro.SetFocus
    End If
End Sub
Private Sub txtbairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCidade.SetFocus
End Sub
'-------- cidade
Private Sub txtcidade_GotFocus()
    txtCidade.BackColor = &H80FFFF
End Sub
Private Sub txtcidade_LostFocus()
    txtCidade.BackColor = &H80000014
    If Len(txtCidade.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtCidade.text)
        txtCidade.SetFocus
    End If
End Sub
Private Sub txtcidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtRua.SetFocus
End Sub


'-------- Rua
Private Sub txtRua_GotFocus()
    txtRua.BackColor = &H80FFFF
End Sub
Private Sub txtRua_LostFocus()
    txtRua.BackColor = &H80000014
    If Len(txtRua.text) > 100 Then
        MsgBox "Comprimento do campo e de 100 digitos, voce digitou " & Len(txtRua.text)
        txtRua.SetFocus
    End If
End Sub
Private Sub txtRua_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtContato.SetFocus
End Sub

'-------- Contato
Private Sub txtContato_GotFocus()
    txtContato.BackColor = &H80FFFF
End Sub
Private Sub txtContato_LostFocus()
    txtContato.BackColor = &H80000014
    If Len(txtContato.text) > 30 Then
        MsgBox "Comprimento do campo e de 30 digitos, voce digitou " & Len(txtContato.text)
        txtContato.SetFocus
    End If
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtInscricao.SetFocus
End Sub

'-------- txtInscricao
Private Sub txtInscricao_GotFocus()
    txtInscricao.BackColor = &H80FFFF
End Sub
Private Sub txtInscricao_LostFocus()
    txtInscricao.BackColor = &H80000014
    If Len(txtInscricao.text) > 30 Then
        MsgBox "Comprimento do campo e de 30 digitos, voce digitou " & Len(txtInscricao.text)
        txtInscricao.SetFocus
    End If
End Sub
Private Sub txtInscricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCnpj.SetFocus
End Sub


'-------- cnpj
Private Sub txtcnpj_GotFocus()
    txtCnpj.BackColor = &H80FFFF
End Sub
Private Sub txtcnpj_LostFocus()
    txtCnpj.BackColor = &H80000014
    If Len(txtCnpj.text) > 30 Then
        MsgBox "Comprimento do campo e de 30 digitos, voce digitou " & Len(txtCnpj.text)
        txtCnpj.SetFocus
    End If
End Sub
Private Sub txtcnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub



