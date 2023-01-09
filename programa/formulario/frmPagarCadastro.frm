VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagarCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas a Pagar"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frIncliurItem 
      Caption         =   "Frame3"
      Height          =   1695
      Left            =   3840
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtVencimento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtValorPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin Vendas.VistaButton cmdSairFrame 
         Height          =   615
         Left            =   1560
         TabIndex        =   23
         Top             =   960
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
         Picture         =   "frmPagarCadastro.frx":0000
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
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
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
         Picture         =   "frmPagarCadastro.frx":010A
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label Label9 
         Caption         =   "Valor (R$)"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Incluir Item"
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
         TabIndex        =   25
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.TextBox txtid_ContasPagar 
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtparcela 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Text            =   "1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5520
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtid_contasPagarItem 
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   10215
      Begin MSComctlLib.ListView ListaPagar 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3625
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
      Begin Vendas.VistaButton cmdIncluirItem 
         Height          =   375
         Left            =   9720
         TabIndex        =   14
         Top             =   360
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmPagarCadastro.frx":065C
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin Vendas.VistaButton cmdExcluirItem 
         Height          =   375
         Left            =   9720
         TabIndex        =   15
         Top             =   720
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmPagarCadastro.frx":09AE
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parcela(s)"
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
         TabIndex        =   12
         Top             =   0
         Width           =   10215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   10215
      Begin VB.TextBox txtid_fornecedor 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtFornecedor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   9615
      End
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dados da Conta"
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
         TabIndex        =   10
         Top             =   0
         Width           =   10215
      End
      Begin VB.Image cmdConsultaFornecedor 
         Height          =   315
         Left            =   9720
         Picture         =   "frmPagarCadastro.frx":0D00
         Stretch         =   -1  'True
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Documento Nº"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Historico"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5610
      Width           =   10485
      _ExtentX        =   18494
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
      Left            =   9240
      TabIndex        =   1
      Top             =   4920
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
      Picture         =   "frmPagarCadastro.frx":100A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   8280
      Left            =   0
      Picture         =   "frmPagarCadastro.frx":1114
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10875
   End
End
Attribute VB_Name = "frmPagarCadastro"
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
Dim chave As String
Dim Sql As String
Dim SQsort As String

Private Sub cmdExcluir_Click()

End Sub

Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Me.Width = 10575
    Me.Height = 6360
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmPagarCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPagarCadastro = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdIncluirItem_Click()
    frIncliurItem.Visible = True
    txtVencimento.text = ""
    txtValorPagar.text = ""
    txtVencimento.SetFocus
End Sub

Private Sub cmdSairFrame_Click()
    frIncliurItem.Visible = False
End Sub



Private Sub cmdConsultaFornecedor_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaFornecedor
        .Show 1
    End With

    If IDFornecedor <> "" Then
        txtid_fornecedor.text = IDFornecedor
        IDFornecedor = ""
    End If

    Sql = "SELECT id_fornecedor, fornecedor FROM fornecedores WHERE id_fornecedor = '" & txtid_fornecedor.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_fornecedor")) <> vbNull Then txtid_fornecedor.text = Tabela("id_fornecedor") Else txtid_fornecedor.text = ""
        If VarType(Tabela("fornecedor")) <> vbNull Then txtFornecedor.text = Tabela("fornecedor") Else txtFornecedor.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


'----------------------------------------------------------




Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    If txtid_contasPagar.text = "" Then
        campo = " data_cadastro"
        Scampo = "'" & Format(Date$, "YYYYMMDD") & "'"

        If txtDocumento.text <> "" Then
            campo = campo & ", documento"
            Scampo = Scampo & ", '" & txtDocumento.text & "'"
        End If

        If txtHistorico.text <> "" Then
            campo = campo & ", historico"
            Scampo = Scampo & ", '" & txtHistorico.text & "'"
        End If

        If txtid_fornecedor.text <> "" Then
            campo = campo & ", id_fornecedor"
            Scampo = Scampo & ", '" & txtid_fornecedor.text & "'"
        End If

        sqlIncluir "ContasPagar", campo, Scampo, Me, "N"

        Buscar_id

    End If

    If txtid_contasPagar.text <> "" Then

        campo = "id_contaspagar"
        Scampo = "'" & txtid_contasPagar.text & "'"

        campo = campo & ", parcela"
        Scampo = Scampo & ", '" & txtparcela.text & "'"

        If txtVencimento.text <> "" Then
            campo = campo & ", vencimento"
            Scampo = Scampo & ", '" & Format(txtVencimento.text, "YYYYMMDD") & "'"
        Else
            MsgBox ("Valor do Vencimento em branco.."), vbInformation
            txtVencimento.SetFocus
            Exit Sub
        End If

        If txtValorPagar.text <> "" Then
            campo = campo & ", valorpagar"
            Scampo = Scampo & ", '" & FormatValor(txtValorPagar.text, 1) & "'"
        Else
            MsgBox ("Valor não pode ficar em branco..."), vbInformation
            txtValorPagar.SetFocus
            Exit Sub
        End If


        sqlIncluir "ContasPagaritem", campo, Scampo, Me, "S"

        txtparcela.text = txtparcela.text + 1

        frIncliurItem.Visible = False

        txtDocumento.SetFocus

        Lista ("")

    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(id_contasPagar) as MaxID "
    Sql = Sql & " FROM ContasPagar"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then txtid_contasPagar.text = Tabela("maxid") Else txtid_contasPagar.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

'--------------------------- define dados da lista grid Consulta
Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Pagar As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim mPagas As Double
    Dim mPagar As Double
    ' conecta ao banco de dados
    Set Pagar = CreateObject("ADODB.Recordset")

    If SQsort = "" Then SQsort = "contaspagaritem.vencimento"

    Sql = " SELECT contaspagarItem.*"
    Sql = Sql & " From"
    Sql = Sql & " contaspagarItem"
    Sql = Sql & " Where"
    Sql = Sql & " id_contasPagar = '" & txtid_contasPagar.text & "'"
    Sql = Sql & " order by Parcela"
    '  Sql = Sql & " limit 300"


    Aguarde_Process Me, True
    ' abre um Recrodset da Tabela Pagar
    If Pagar.State = 1 Then Pagar.Close
    Pagar.Open Sql, banco, adOpenKeyset, adLockOptimistic
    Aguarde_Process Me, False

    ListaPagar.ColumnHeaders.Clear
    ListaPagar.ListItems.Clear

    If Pagar.RecordCount = 0 Then
        ' muda o curso para o normal
        Exit Sub
    End If

    ListaPagar.ColumnHeaders.Add , , "Parcela", 1800
    ListaPagar.ColumnHeaders.Add , , "Vencimento", 2500, lvwColumnCenter
    ListaPagar.ColumnHeaders.Add , , "Valor", 2500, lvwColumnRight
    ListaPagar.ColumnHeaders.Add , , "Data Pagamento", 2500, lvwColumnCenter

    If Pagar.BOF = True And Pagar.EOF = True Then Exit Sub

    mPagar = 0
    mPagas = 0

    While Not Pagar.EOF
        If VarType(Pagar("parcela")) <> vbNull Then Set itemx = ListaPagar.ListItems.Add(, , Pagar("parcela"))
        If VarType(Pagar("vencimento")) <> vbNull Then itemx.SubItems(1) = Format(Pagar("vencimento"), "DD/MM/YYYY") Else itemx.SubItems(1) = ""
        If VarType(Pagar("valorpagar")) <> vbNull Then itemx.SubItems(2) = Format(Pagar("valorpagar"), "###,##0.00") Else itemx.SubItems(2) = ""
        If VarType(Pagar("datapagto")) <> vbNull Then itemx.SubItems(3) = Format(Pagar("datapagto"), "DD/MM/YYYY") Else itemx.SubItems(3) = ""
        If VarType(Pagar("id_contaspagarItem")) <> vbNull Then itemx.Tag = Pagar("id_contaspagarItem")
        Pagar.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaPagar, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Pagar.State = 1 Then Pagar.Close
    Set Pagar = Nothing

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


Private Sub cmdExcluirItem_Click()
    On Error GoTo trata_erro
    Dim Excluir As Boolean


    Sqlconsulta = "id_contasPagarItem = '" & txtid_ContasPagarItem.text & "'"
    confirma = MsgBox("Confirma Exclusão da parcela contas a pagar", vbQuestion + vbYesNo, "Excluir")
    If confirma = vbYes Then
        sqlDeletar "ContasPagarItem", Sqlconsulta, Me, "S"
    End If

    Lista ("")

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub




'----------------------------------------------------------

'--- txtdocumento
Private Sub txtdocumento_GotFocus()
    txtDocumento.BackColor = &H80FFFF
End Sub
Private Sub txtdocumento_LostFocus()
    txtDocumento.BackColor = &H80000014
    If Len(txtDocumento.text) > 10 Then
        MsgBox "Comprimento do campo e de 10 digitos, voce digitou " & Len(txtDocumento.text)
        txtDocumento.SetFocus
    End If
End Sub
Private Sub txtdocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtHistorico.SetFocus
End Sub

'--- txthistorico
Private Sub txthistorico_GotFocus()
    txtHistorico.BackColor = &H80FFFF
End Sub
Private Sub txthistorico_LostFocus()
    txtHistorico.BackColor = &H80000014
    If Len(txtHistorico.text) > 40 Then
        MsgBox "Comprimento do campo e de 40 digitos, voce digitou " & Len(txtHistorico.text)
        txtHistorico.SetFocus
    End If
End Sub
Private Sub txthistorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdIncluirItem_Click
End Sub



'--- txtVencimento
Private Sub txtvencimento_GotFocus()
    txtVencimento.BackColor = &H80FFFF
End Sub
Private Sub txtvencimento_LostFocus()
    txtVencimento.BackColor = &H80000014
    If txtVencimento.text = "h" Or txtVencimento.text = "H" Then txtVencimento.text = Format(Now, "dd/mm/yyyy")
    txtVencimento.text = SoNumero(txtVencimento.text)
    If txtVencimento.text <> "" Then
        txtVencimento.text = Mid(txtVencimento.text, 1, 2) & "/" & Mid(txtVencimento.text, 3, 2) & "/" & Mid(txtVencimento.text, 5, 4)
    End If
End Sub
Private Sub txtvencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtValorPagar.SetFocus
End Sub

'--- txtvalorPagar
Private Sub txtvalorPagar_GotFocus()
    txtValorPagar.BackColor = &H80FFFF
End Sub
Private Sub txtvalorPagar_LostFocus()
    txtValorPagar.BackColor = &H80000014
    txtValorPagar.text = Format(txtValorPagar.text, "###,##0.00")
End Sub
Private Sub txtvalorPagar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub

