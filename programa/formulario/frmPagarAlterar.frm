VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagarAlterar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas a Pagar"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   8175
      Begin VB.TextBox txtFornecedor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txtVencimento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtValorPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6720
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contas a Pagar"
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
         TabIndex        =   18
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label Label3 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Documento"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Historico"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Valor (R$)"
         Height          =   255
         Left            =   6720
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtid_contasPagar 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtid_contasPagarItem 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3060
      Width           =   8535
      _ExtentX        =   15055
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
      Left            =   7200
      TabIndex        =   13
      Top             =   2280
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
      Picture         =   "frmPagarAlterar.frx":0000
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
      Left            =   6000
      TabIndex        =   14
      Top             =   2280
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
      Picture         =   "frmPagarAlterar.frx":010A
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
      Left            =   4800
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmPagarAlterar.frx":065C
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   0
      Picture         =   "frmPagarAlterar.frx":0BAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8595
   End
End
Attribute VB_Name = "frmPagarAlterar"
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



Private Sub Form_Load()
    Me.Width = 8625
    Me.Height = 3810
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub Form_Activate()
    txtDocumento.SetFocus
    If txtTipo.text = "A" Or txtTipo.text = "E" Then AutalizaCadastro

End Sub
Private Sub cmdSair_Click()
    Unload Me
    Set frmPagarAlterar = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPagarAlterar = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub AutalizaCadastro()
    On Error GoTo trata_erro


    Dim ContasPagar As ADODB.Recordset
    ' conecta ao banco de dados

    Set ContasPagar = CreateObject("ADODB.Recordset")    '''

    ' abre um Recrodset da Tabela ContasPagar
    Sql = " SELECT contaspagar.*, contaspagaritem.vencimento, contaspagaritem.valorpagar,"
    Sql = Sql & " contaspagaritem.datapagto,contaspagaritem.valorpago,"
    Sql = Sql & " Fornecedores.id_fornecedor , Fornecedores.fornecedor"
    Sql = Sql & " From"
    Sql = Sql & " contaspagar"
    Sql = Sql & " LEFT JOIN contaspagaritem ON contaspagar.id_contasPagar = contaspagaritem.id_contasPagar"
    Sql = Sql & " LEFT JOIN fornecedores ON contaspagar.id_fornecedor = fornecedores.id_fornecedor"
    Sql = Sql & " where "
    Sql = Sql & " contaspagaritem.id_contaspagaritem = '" & txtid_ContasPagarItem.text & "'"

    If ContasPagar.State = 1 Then ContasPagar.Close
    ContasPagar.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If ContasPagar.RecordCount > 0 Then

        If VarType(ContasPagar("id_contaspagar")) <> vbNull Then txtid_contasPagar.text = ContasPagar("id_contaspagar") Else txtid_contasPagar.text = ""
        If VarType(ContasPagar("fornecedor")) <> vbNull Then txtFornecedor.text = ContasPagar("fornecedor") Else txtFornecedor.text = ""
        If VarType(ContasPagar("documento")) <> vbNull Then txtDocumento.text = ContasPagar("documento") Else txtDocumento.text = ""
        If VarType(ContasPagar("historico")) <> vbNull Then txtHistorico.text = ContasPagar("historico") Else txtHistorico.text = ""
        If VarType(ContasPagar("vencimento")) <> vbNull Then txtVencimento.text = Format(ContasPagar("vencimento"), "DD/MM/yyyy") Else txtVencimento.text = ""
        If VarType(ContasPagar("valorpagar")) <> vbNull Then txtValorPagar.text = Format(ContasPagar("valorpagar"), "###,##0.00") Else txtValorPagar.text = ""

    End If
    If ContasPagar.State = 1 Then ContasPagar.Close
    Set ContasPagar = Nothing

    If txtTipo.text = "E" Then cmdGravar.Enabled = False
    If txtTipo.text = "A" Then cmdExcluir.Enabled = False

    txtDocumento.SetFocus
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "id_contasPagar = '" & txtid_contasPagar.text & "'"
    If txtDocumento.text <> "" Then
        campo = " documento = '" & Mid(txtDocumento.text, 1, 10) & "'"
    Else
        campo = " documento = '0'"
    End If
    If txtHistorico.text <> "" Then campo = campo & ", historico = '" & Mid(txtHistorico.text, 1, 40) & "'"
    sqlAlterar "Contaspagar", campo, Sqlconsulta, Me, "N"


    Sqlconsulta = "id_contaspagaritem = '" & txtid_ContasPagarItem.text & "'"
    If txtVencimento.text <> "" Then campo = "vencimento = '" & Format(txtVencimento.text, "YYYYMMDD") & "'"
    If txtValorPagar.text <> "" Then campo = campo & ", valorpagar = '" & FormatValor(txtValorPagar.text, 1) & "'"
    sqlAlterar "ContaspagarItem", campo, Sqlconsulta, Me, "N"

    MsgBox ("Dados Alterados com sucesso.."), vbInformation

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro
    Dim Excluir As Boolean

    Sqlconsulta = "id_contasPagarItem = '" & txtid_ContasPagarItem.text & "'"
    confirma = MsgBox("Confirma Exclusão da parcela contas a pagar", vbQuestion + vbYesNo, "Excluir")
    If confirma = vbYes Then
        sqlDeletar "ContasPagarItem", Sqlconsulta, Me, "S"

        Unload Me

    End If
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
    If KeyAscii = vbKeyReturn Then txtVencimento.SetFocus
End Sub

'--- txtVencimento
Private Sub txtvencimento_GotFocus()
    txtVencimento.BackColor = &H80FFFF
End Sub
Private Sub txtvencimento_LostFocus()
    txtVencimento.BackColor = &H80000014
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





