VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagarConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas a Pagar"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   1575
      Begin VB.OptionButton optemAberto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "em Aberto"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optPagas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pagas"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optGeral 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Geral"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contas"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1920
      TabIndex        =   22
      Top             =   4080
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   15000
      ScaleHeight     =   240
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   15000
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_ContasPagarItem 
      Height          =   285
      Left            =   15000
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_fornecedor 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFornecedor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   4815
   End
   Begin VB.TextBox txtNF 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListaPagar 
      Height          =   6375
      Left            =   5640
      TabIndex        =   6
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11245
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
      TabIndex        =   7
      Top             =   7875
      Width           =   15150
      _ExtentX        =   26723
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
      Left            =   13920
      TabIndex        =   8
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
      Picture         =   "frmPagarConsulta.frx":0000
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
      Left            =   11520
      TabIndex        =   9
      Top             =   7080
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
      Picture         =   "frmPagarConsulta.frx":010A
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
      Left            =   12720
      TabIndex        =   10
      Top             =   7080
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
      Picture         =   "frmPagarConsulta.frx":045C
      Pictures        =   1
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
      TabIndex        =   11
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108658689
      CurrentDate     =   41801
   End
   Begin MSComCtl2.MonthView txtDataI 
      Height          =   2370
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   108658689
      CurrentDate     =   41801
   End
   Begin Vendas.VistaButton cmdConsultar 
      Height          =   615
      Left            =   3600
      TabIndex        =   13
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      Picture         =   "frmPagarConsulta.frx":07AE
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdExcluir 
      Height          =   615
      Left            =   10320
      TabIndex        =   33
      Top             =   7080
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmPagarConsulta.frx":07CA
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdPagar 
      Height          =   615
      Left            =   9120
      TabIndex        =   34
      Top             =   7080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
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
      Picture         =   "frmPagarConsulta.frx":0D1C
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Total Pagas (R$)"
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
      Left            =   7440
      TabIndex        =   32
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblPagas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      Height          =   550
      Left            =   7440
      TabIndex        =   31
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Total a Pagar (R$)"
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
      TabIndex        =   30
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label lblDinheiro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      Height          =   550
      Left            =   5640
      TabIndex        =   29
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Historico"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   3840
      Width           =   3495
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
      TabIndex        =   21
      Top             =   240
      Width           =   6255
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
      Left            =   11880
      TabIndex        =   20
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   600
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   600
      Width           =   1095
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
      TabIndex        =   16
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   5145
   End
   Begin VB.Image cmdConsultaFornecedor 
      Height          =   255
      Left            =   5100
      Picture         =   "frmPagarConsulta.frx":126E
      Stretch         =   -1  'True
      Top             =   3520
      Width           =   345
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Documento (NF)"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7380
      Left            =   120
      Picture         =   "frmPagarConsulta.frx":1578
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   0
      Picture         =   "frmPagarConsulta.frx":53C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16680
   End
End
Attribute VB_Name = "frmPagarConsulta"
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



Private Sub cmdExcluir_Click()
    If txtid_ContasPagarItem.text <> "" Then
        With frmPagarAlterar
            .txtid_ContasPagarItem.text = txtid_ContasPagarItem.text
            .txtTipo.text = "E"
            .Show 1
        End With
    Else
        MsgBox ("Favor selecionar uma NF..."), vbInformation
        Exit Sub
    End If
End Sub

Private Sub cmdPagar_Click()
    If txtid_ContasPagarItem.text <> "" Then
        With frmPagarPagamento
            .txtid_ContasPagarItem.text = txtid_ContasPagarItem.text
            .Show 1
        End With
    Else
        MsgBox ("Favor selecionar uma NF..."), vbInformation
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 15240
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
    Set frmPagarConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPagarConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdConsultaFornecedor_Click()
    On Error GoTo trata_erro
    Dim fornecedor As ADODB.Recordset
    Dim vControle As Integer

    vControle = frmBuscaSimples.getKey("fornecedores", "fornecedor")

    Set fornecedor = CreateObject("ADODB.Recordset")

    If Not vControle = -1 Then
        Sql = "SELECT id_fornecedor,  fornecedor FROM fornecedores WHERE id_fornecedor = " & vControle
        If fornecedor.State = 1 Then fornecedor.Close
        fornecedor.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If fornecedor.RecordCount > 0 Then
            If VarType(fornecedor("id_fornecedor")) <> vbNull Then txtid_fornecedor.text = fornecedor("id_fornecedor") Else txtid_fornecedor.text = ""
            If VarType(fornecedor("fornecedor")) <> vbNull Then txtFornecedor.text = fornecedor("fornecedor") Else txtFornecedor.text = ""
        End If
    End If

    If fornecedor.State = 1 Then fornecedor.Close
    Set fornecedor = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub cmdIncluir_Click()
    With frmPagarCadastro
        .Show 1
    End With
End Sub


Private Sub cmdAlterar_Click()
    If txtid_ContasPagarItem.text <> "" Then
        With frmPagarAlterar
            .txtid_ContasPagarItem.text = txtid_ContasPagarItem.text
            .txtTipo.text = "A"
            .Show 1
        End With
    Else
        MsgBox ("Favor selecionar uma NF..."), vbInformation
        Exit Sub
    End If
End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Entradas As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mPagar As Double
    Dim mPagas As Double
    ' conecta ao banco de dados
    Set Entradas = CreateObject("ADODB.Recordset")

    Sql = " SELECT contaspagar.*, contaspagaritem.vencimento, contaspagaritem.valorpagar,"
    Sql = Sql & " contaspagaritem.datapagto,contaspagaritem.valorpago,contaspagaritem.id_contasPagarItem,"
    Sql = Sql & " Fornecedores.id_fornecedor , Fornecedores.fornecedor"
    Sql = Sql & " From"
    Sql = Sql & " contaspagar"
    Sql = Sql & " LEFT JOIN contaspagaritem ON contaspagar.id_contasPagar = contaspagaritem.id_contasPagar"
    Sql = Sql & " LEFT JOIN fornecedores ON contaspagar.id_fornecedor = fornecedores.id_fornecedor"
    Sql = Sql & " Where"
    Sql = Sql & SQconsulta
    Sql = Sql & " order by contaspagaritem.vencimento"
    Sql = Sql & " limit 100"

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Entradas
    If Entradas.State = 1 Then Entradas.Close
    Entradas.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaPagar.ColumnHeaders.Clear
    ListaPagar.ListItems.Clear

    If Entradas.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Entradas encontrado(s): " & Entradas.RecordCount

    ListaPagar.ColumnHeaders.Add , , "Fornecedor", 3200
    ListaPagar.ColumnHeaders.Add , , "Vencimento", 1500, lvwColumnCenter
    ListaPagar.ColumnHeaders.Add , , "A Pagar(R$)", 1500, lvwColumnRight
    ListaPagar.ColumnHeaders.Add , , "Pagas (R$)", 1500, lvwColumnCenter
    ListaPagar.ColumnHeaders.Add , , "NF", 1500, lvwColumnRight

    mPagar = 0
    mPagas = 0

    If Entradas.BOF = True And Entradas.EOF = True Then Exit Sub
    While Not Entradas.EOF
        If VarType(Entradas("fornecedor")) <> vbNull Then Set itemx = ListaPagar.ListItems.Add(, , Entradas("fornecedor"))
        If VarType(Entradas("vencimento")) <> vbNull Then itemx.SubItems(1) = Format(Entradas("vencimento"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Entradas("valorpagar")) <> vbNull Then itemx.SubItems(2) = Format(Entradas("valorpagar"), "###,##0.00") Else itemx.SubItems(2) = ""
        If VarType(Entradas("valorpago")) <> vbNull Then itemx.SubItems(3) = Format(Entradas("valorpago"), "###,##0.00") Else itemx.SubItems(3) = ""
        If VarType(Entradas("documento")) <> vbNull Then itemx.SubItems(4) = Entradas("documento") Else itemx.SubItems(4) = ""
        If VarType(Entradas("id_contasPagarItem")) <> vbNull Then itemx.Tag = Entradas("id_contasPagarItem")

        If VarType(Entradas("valorpago")) <> vbNull Then
            mPagas = mPagas + Entradas("valorpago")
        Else
            mPagar = mPagar + Entradas("valorpagar")
        End If

        Entradas.MoveNext
    Wend

    lblDinheiro.Caption = Format(mPagar, "###,##0.00")
    lblPagas.Caption = Format(mPagas, "###,##0.00")

    'Zebra o listview
    If LVZebra(ListaPagar, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Entradas.State = 1 Then Entradas.Close
    Set Entradas = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaPagar_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_ContasPagarItem.text = ListaPagar.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()



    If optGeral.Value = True Then
        Sqlconsulta = " 1=1 "
    End If

    If optemAberto.Value = True Then
        Sqlconsulta = " contaspagaritem.datapagto is null "
    End If

    If optPagas.Value = True Then
        Sqlconsulta = " contaspagaritem.vencimento Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "'"
        Sqlconsulta = Sqlconsulta & " And '" & Format(lbldataF.Caption, "YYYYMMDD") & "' and contaspagaritem.datapagto is not null"
    End If

    If txtid_fornecedor.text <> "" Then Sqlconsulta = Sqlconsulta & " and contaspagar.id_fornecedor = '" & txtid_fornecedor.text & "'"
    If txtNF.text <> "" Then Sqlconsulta = Sqlconsulta & " and contaspagar.documento like '%" & txtNF.text & "%'"

    Lista (Sqlconsulta)

End Sub


Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub

Private Sub cmdLimpar_Click()
    txtDataI.Value = Now
    txtDataF.Value = Now
    txtid_fornecedor.text = ""
    txtFornecedor.text = ""
    txtNF.text = ""
    cmdConsultar_Click
End Sub
