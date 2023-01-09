VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntradaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNF 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtid_fornecedor 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_entrada 
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6240
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListaEntradas 
      Height          =   6615
      Left            =   5640
      TabIndex        =   2
      Top             =   480
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
      TabIndex        =   3
      Top             =   7935
      Width           =   14985
      _ExtentX        =   26432
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
      Top             =   7200
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
      Picture         =   "frmentradaConsulta.frx":0000
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
      Top             =   7200
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
      Picture         =   "frmentradaConsulta.frx":010A
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
      Top             =   7200
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
      Picture         =   "frmentradaConsulta.frx":045C
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
      TabIndex        =   10
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
   Begin Vendas.VistaButton cmdConsultar 
      Height          =   615
      Left            =   3720
      TabIndex        =   21
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
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
      Picture         =   "frmentradaConsulta.frx":07AE
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdLimpar 
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
      Caption         =   "Limpar Consulta"
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
      Picture         =   "frmentradaConsulta.frx":07CA
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.TextBox txtFornecedor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Documento (NF)"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Image cmdConsultaFornecedor 
      Height          =   315
      Left            =   5040
      Picture         =   "frmentradaConsulta.frx":07E6
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   360
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   5175
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
      TabIndex        =   15
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7425
      Left            =   120
      Picture         =   "frmentradaConsulta.frx":0AF0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5415
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
      TabIndex        =   5
      Top             =   240
      Width           =   3135
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
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmentradaConsulta.frx":28E0
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   15120
   End
End
Attribute VB_Name = "frmEntradaConsulta"
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

Private Sub cmdLimpar_Click()
    txtDataI.Value = Now
    txtDataF.Value = Now
    txtid_fornecedor.text = ""
    txtFornecedor.text = ""
    txtNF.text = ""
    cmdConsultar_Click
End Sub

Private Sub Form_Activate()
' If txtTipo.text = "A" Or txtTipo.text = "E" Then AutalizaCadastro
End Sub

Private Sub Form_Load()
    Me.Width = 15075
    Me.Height = 8685
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
    Set frmEntradaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEntradaConsulta = Nothing
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
    With frmEntradaCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
End Sub


Private Sub cmdAlterar_Click()
    If txtid_entrada.text <> "" Then
        With frmEntradaCadastro
            .txtid_entrada.text = txtid_entrada.text
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
    ' conecta ao banco de dados
    Set Entradas = CreateObject("ADODB.Recordset")

    Sql = " SELECT entrada.id_fornecedor, entrada.nfData, entrada.Nfnro, entrada.id_entrada,"
    Sql = Sql & " Fornecedores.id_fornecedor , Fornecedores.fornecedor"
    Sql = Sql & " From"
    Sql = Sql & " Entrada"
    Sql = Sql & " LEFT JOIN fornecedores ON entrada.id_fornecedor = fornecedores.id_fornecedor"
    Sql = Sql & " Where"
    Sql = Sql & SQconsulta
    Sql = Sql & " order by fornecedores.fornecedor"
    Sql = Sql & " limit 100"

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Entradas
    If Entradas.State = 1 Then Entradas.Close
    Entradas.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaEntradas.ColumnHeaders.Clear
    ListaEntradas.ListItems.Clear

    If Entradas.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Entradas encontrado(s): " & Entradas.RecordCount

    ListaEntradas.ColumnHeaders.Add , , "Fornecedor", 5900
    ListaEntradas.ColumnHeaders.Add , , "Data Entrega", 1500, lvwColumnCenter
    ListaEntradas.ColumnHeaders.Add , , "NF", 1500, lvwColumnRight

    If Entradas.BOF = True And Entradas.EOF = True Then Exit Sub
    While Not Entradas.EOF
        If VarType(Entradas("fornecedor")) <> vbNull Then Set itemx = ListaEntradas.ListItems.Add(, , Entradas("fornecedor"))
        If VarType(Entradas("nfdata")) <> vbNull Then itemx.SubItems(1) = Format(Entradas("nfdata"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Entradas("nfNro")) <> vbNull Then itemx.SubItems(2) = Entradas("nfNro") Else itemx.SubItems(2) = ""
        If VarType(Entradas("id_entrada")) <> vbNull Then itemx.Tag = Entradas("id_entrada")
        Entradas.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaEntradas, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Entradas.State = 1 Then Entradas.Close
    Set Entradas = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaEntradas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_entrada.text = ListaEntradas.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()

    Sqlconsulta = " 1=1 "

    If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
        Sqlconsulta = Sqlconsulta & " and entrada.nfdata Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
    End If

    If txtid_fornecedor.text <> "" Then Sqlconsulta = Sqlconsulta & " and entrada.id_fornecedor = '" & txtid_fornecedor.text & "'"
    If txtNF.text <> "" Then Sqlconsulta = Sqlconsulta & " and entrada.nfnro like '%" & txtNF.text & "%'"

    Lista (Sqlconsulta)

End Sub


Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub

































