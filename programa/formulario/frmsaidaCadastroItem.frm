VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsaidaCadastroItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saida"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdataAcerto 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtid_saidaAcerto 
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frIncliurItem 
      Caption         =   "Frame3"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   10815
      Begin VB.TextBox txtid_estoque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtQuantidadeI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPrecoCustoI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9360
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDescricaoI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lblUnidadeI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7440
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Image cmdConsulInsumos 
         Height          =   360
         Left            =   1440
         Picture         =   "frmsaidaCadastroItem.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Consulta CEP "
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Quantidade"
         Height          =   255
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   975
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
         TabIndex        =   6
         Top             =   0
         Width           =   10815
      End
      Begin VB.Label Label12 
         Caption         =   "Preco Custo (R$)"
         Height          =   255
         Left            =   9360
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Descrição"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Unidade"
         Height          =   255
         Left            =   7440
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox txtid_saida 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   11160
      _ExtentX        =   19685
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
      Left            =   9840
      TabIndex        =   16
      Top             =   1680
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
      Picture         =   "frmsaidaCadastroItem.frx":030A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdGravarI 
      Height          =   615
      Left            =   8640
      TabIndex        =   17
      Top             =   1680
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
      Picture         =   "frmsaidaCadastroItem.frx":0414
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   0
      Picture         =   "frmsaidaCadastroItem.frx":0966
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11115
   End
End
Attribute VB_Name = "frmsaidaCadastroItem"
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
Dim SaldoAn As Double


Private Sub Form_Activate()
    txtid_estoque.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 11145
    Me.Height = 3285
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu


End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmsaidaCadastroItem = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmsaidaCadastroItem = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------- campos do formulario--------------------------------------------------


'---------------------------------------------------------------------------------------------------------------------

Private Sub cmdIncluirItem_Click()
    txtid_estoque.text = ""
    txtid_estoque.text = ""
    lblDescricaoI.Caption = ""
    lblUnidadeI.Caption = ""
    txtQuantidadeI.text = ""
    txtPrecoCustoI.text = ""
    frIncliurItem.Visible = True
    txtid_estoque.SetFocus
End Sub

Private Sub cmdConsulInsumos_Click()
    On Error GoTo trata_erro
    Dim id_estoque As ADODB.Recordset
    ' conecta ao banco de dados

    Set id_estoque = CreateObject("ADODB.Recordset")    '''


    With frmConsultaEstoque
        .Show 1
    End With

    If IDEstoque <> "" Then
        txtid_estoque.text = IDEstoque

        ' abre um Recrodset da Tabela id_estoque
        Sql = " SELECT estoques.*"
        Sql = Sql & " From"
        Sql = Sql & " estoques"
        Sql = Sql & " where "
        Sql = Sql & " id_estoque = '" & IDEstoque & "'"

        If id_estoque.State = 1 Then id_estoque.Close
        id_estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If id_estoque.RecordCount > 0 Then

            If VarType(id_estoque("id_estoque")) <> vbNull Then txtid_estoque.text = id_estoque("id_estoque") Else txtid_estoque.text = ""
            If VarType(id_estoque("descricao")) <> vbNull Then lblDescricaoI.Caption = id_estoque("descricao") Else lblDescricaoI.Caption = ""
            If VarType(id_estoque("unidade")) <> vbNull Then lblUnidadeI.Caption = id_estoque("unidade") Else lblUnidadeI.Caption = ""
            If VarType(id_estoque("preco_compra")) <> vbNull Then txtPrecoCustoI.text = Format(id_estoque("preco_compra"), "###,##0.00") Else txtPrecoCustoI.text = ""

        End If
        If id_estoque.State = 1 Then id_estoque.Close
        Set id_estoque = Nothing

        IDEstoque = ""

        txtQuantidadeI.SetFocus
    End If


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'------------ Quantiade
Private Sub txtquantidadei_GotFocus()
    txtQuantidadeI.BackColor = &H80FFFF
End Sub
Private Sub txtquantidadei_LostFocus()
    txtQuantidadeI.BackColor = &H80000014
    txtQuantidadeI.text = Format(txtQuantidadeI.text, "###,##0.000")
End Sub
Private Sub txtquantidadei_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtPrecoCustoI.SetFocus
End Sub

'------------ Quantiade
Private Sub txtPrecoCustoi_GotFocus()
    txtPrecoCustoI.BackColor = &H80FFFF
End Sub
Private Sub txtPrecoCustoi_LostFocus()
    txtPrecoCustoI.BackColor = &H80000014
    txtPrecoCustoI.text = Format(txtPrecoCustoI.text, "###,##0.00")
End Sub
Private Sub txtPrecoCustoi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravarI.SetFocus
End Sub


'------------ id_estoque
Private Sub txtid_estoque_GotFocus()
    txtid_estoque.BackColor = &H80FFFF
End Sub
Private Sub txtid_estoque_LostFocus()
    txtid_estoque.BackColor = &H80000014
    If Len(txtid_estoque.text) > 6 Then
        MsgBox "Comprimento do campo e de  digitos, voce digitou " & Len(txtid_estoque.text)
        txtid_estoque.SetFocus
    End If
    txtid_estoque.text = SoNumero(txtid_estoque.text)
End Sub
Private Sub txtid_estoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then procura_id_estoque
End Sub

Private Sub procura_id_estoque()
    On Error GoTo trata_erro
    Dim id_estoque As ADODB.Recordset
    ' conecta ao banco de dados

    Set id_estoque = CreateObject("ADODB.Recordset")    '''

    ' abre um Recrodset da Tabela id_estoque
    Sql = " SELECT estoques.*"
    Sql = Sql & " From"
    Sql = Sql & " estoques"
    Sql = Sql & " where "
    Sql = Sql & " id_estoque = '" & txtid_estoque.text & "'"

    If id_estoque.State = 1 Then id_estoque.Close
    id_estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If id_estoque.RecordCount > 0 Then

        If VarType(id_estoque("id_estoque")) <> vbNull Then txtid_estoque.text = id_estoque("id_estoque") Else txtid_estoque.text = ""

    End If
    If id_estoque.State = 1 Then id_estoque.Close
    Set id_estoque = Nothing

    txtQuantidadeI.SetFocus

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub txtid_estoque_Change()
    On Error GoTo trata_erro
    Dim Estoque As ADODB.Recordset
    ' conecta ao banco de dados

    Set Estoque = CreateObject("ADODB.Recordset")    '''

    ' abre um Recrodset da Tabela Estoque
    Sql = " SELECT estoques.*"
    Sql = Sql & " From"
    Sql = Sql & " estoques"
    Sql = Sql & " where "
    Sql = Sql & " id_Estoque = '" & txtid_estoque.text & "'"

    If Estoque.State = 1 Then Estoque.Close
    Estoque.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Estoque.RecordCount > 0 Then

        If VarType(Estoque("id_Estoque")) <> vbNull Then txtid_estoque.text = Estoque("id_Estoque") Else txtid_estoque.text = ""
        If VarType(Estoque("descricao")) <> vbNull Then lblDescricaoI.Caption = Estoque("descricao") Else lblDescricaoI.Caption = ""
        If VarType(Estoque("unidade")) <> vbNull Then lblUnidadeI.Caption = Estoque("unidade") Else lblUnidadeI.Caption = ""
        If VarType(Estoque("preco_compra")) <> vbNull Then txtPrecoCustoI.text = Format(Estoque("preco_compra"), "###,##0.00") Else txtPrecoCustoI.text = ""

    End If
    If Estoque.State = 1 Then Estoque.Close
    Set Estoque = Nothing

    txtQuantidadeI.SetFocus

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdGravarI_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double

    campo = "id_saidaacerto"
    Scampo = "'" & txtid_saidaAcerto.text & "'"

    campo = campo & ", dataSaida"
    Scampo = Scampo & ", '" & Format(txtdataAcerto.text, "YYYYMMDD") & "'"

    If txtid_estoque.text <> "" Then
        campo = campo & ", id_estoque"
        Scampo = Scampo & ", '" & txtid_estoque.text & "'"
    End If

    If txtQuantidadeI.text <> "" Then
        campo = campo & ", quantidade"
        Scampo = Scampo & ", '" & FormatValor(txtQuantidadeI.text, 1) & "'"
    End If

    If txtPrecoCustoI.text <> "" Then
        campo = campo & ", preco_custo"
        Scampo = Scampo & ", '" & FormatValor(txtPrecoCustoI.text, 1) & "'"
    End If

    ' Incluir valor na tabela EntradaItens
    sqlIncluir "saida", campo, Scampo, Me, "S"

    IncluirSaldo

    txtid_estoque.text = ""
    txtid_estoque.text = ""
    lblDescricaoI.Caption = ""
    lblUnidadeI.Caption = ""
    txtQuantidadeI.text = ""
    txtPrecoCustoI.text = ""

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub IncluirSaldo()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    ' Altera saldo na tabela Grupo
    Sqlconsulta = " id_estoque = '" & txtid_estoque.text & "'"

    Sql = "SELECT estoquesaldo.* "
    Sql = Sql & " FROM estoquesaldo"
    Sql = Sql & " where"
    Sql = Sql & Sqlconsulta

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("saldo")) <> vbNull Then mSaldo = Tabela("saldo") Else mSaldo = "0"

        mSaldo = mSaldo - txtQuantidadeI.text
        campo = "saldo ='" & FormatValor(mSaldo, 1) & "'"
        sqlAlterar "Estoquesaldo", campo, Sqlconsulta, Me, "N"


    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub









