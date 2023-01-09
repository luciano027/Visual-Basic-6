VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaidaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saída Consulta"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHistorico 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   5155
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6960
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_saidaacerto 
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView ListaSaidas 
      Height          =   6615
      Left            =   5640
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   7905
      Width           =   14925
      _ExtentX        =   26326
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
      TabIndex        =   5
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
      Picture         =   "frmSaidaConsulta.frx":0000
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
      TabIndex        =   6
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
      Picture         =   "frmSaidaConsulta.frx":010A
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
      TabIndex        =   7
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
      Picture         =   "frmSaidaConsulta.frx":045C
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
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1720
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
      Picture         =   "frmSaidaConsulta.frx":07AE
      Pictures        =   2
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Historico"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   5160
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   600
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
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
      TabIndex        =   10
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7260
      Left            =   120
      Picture         =   "frmSaidaConsulta.frx":07CA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5415
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
      TabIndex        =   9
      Top             =   240
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
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   16680
   End
   Begin VB.Image Image2 
      Height          =   7920
      Left            =   0
      Picture         =   "frmSaidaConsulta.frx":2540
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15435
   End
End
Attribute VB_Name = "frmSaidaConsulta"
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

Private Sub Form_Activate()
' If txtTipo.text = "A" Or txtTipo.text = "E" Then AutalizaCadastro
End Sub

Private Sub Form_Load()
    Me.Width = 15015
    Me.Height = 8655
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
    Set frmSaidaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSaidaConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdIncluir_Click()
    With frmSaidaCadastro
        .txtTipo.text = "I"
        .Show 1
    End With
End Sub


Private Sub cmdAlterar_Click()
    If txtid_saidaAcerto.text <> "" Then
        With frmSaidaCadastro
            .txtid_saidaAcerto.text = txtid_saidaAcerto.text
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
    Dim Saidas As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Saidas = CreateObject("ADODB.Recordset")

    Sql = " SELECT saidaacerto.*"
    Sql = Sql & " From"
    Sql = Sql & " saidaacerto"
    Sql = Sql & " Where"
    Sql = Sql & SQconsulta
    Sql = Sql & " order by saidaacerto.dataAcerto"
    Sql = Sql & " limit 100"

    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Saidas
    If Saidas.State = 1 Then Saidas.Close
    Saidas.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaSaidas.ColumnHeaders.Clear
    ListaSaidas.ListItems.Clear

    If Saidas.RecordCount = 0 Then
        ' muda o curso para o normal
        MsgBox ("Informação não encontrada"), vbInformation
        Exit Sub
    End If
    lblConsulta.Caption = " Resultado Consulta"
    lblCadastro.Caption = " Saidas encontrado(s): " & Saidas.RecordCount

    ListaSaidas.ColumnHeaders.Add , , "Historico", 7500
    ListaSaidas.ColumnHeaders.Add , , "Data Acerto", 1500, lvwColumnCenter

    If Saidas.BOF = True And Saidas.EOF = True Then Exit Sub
    While Not Saidas.EOF
        If VarType(Saidas("historico")) <> vbNull Then Set itemx = ListaSaidas.ListItems.Add(, , Saidas("historico"))
        If VarType(Saidas("dataacerto")) <> vbNull Then itemx.SubItems(1) = Format(Saidas("dataacerto"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Saidas("id_SaidaAcerto")) <> vbNull Then itemx.Tag = Saidas("id_SaidaAcerto")
        Saidas.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaSaidas, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Saidas.State = 1 Then Saidas.Close
    Set Saidas = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaSaidas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_saidaAcerto.text = ListaSaidas.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()

    Sqlconsulta = " 1=1 "

    If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
        Sqlconsulta = Sqlconsulta & " and Saidaacerto.dataacerto Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
    End If

    If txtHistorico.text <> "" Then Sqlconsulta = Sqlconsulta & " and Saidaacerto.historico like '%" & txtHistorico.text & "%'"

    Lista (Sqlconsulta)

End Sub


Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub

