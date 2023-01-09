VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmVendasCompras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendas- incluir item"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSaldo 
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7695
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2415
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
         TabIndex        =   3
         Top             =   320
         Width           =   7095
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
         Left            =   5400
         TabIndex        =   13
         Top             =   960
         Width           =   2175
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
         Left            =   2760
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Preço Venda"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Quantidade"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image cmdConsultar 
         Height          =   375
         Left            =   7200
         Picture         =   "frmVendasCompras.frx":0000
         Stretch         =   -1  'True
         Top             =   320
         Width           =   360
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
         TabIndex        =   4
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.TextBox txtid_venda 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   6720
      TabIndex        =   5
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
      Picture         =   "frmVendasCompras.frx":030A
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
      Left            =   5520
      TabIndex        =   6
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
      Picture         =   "frmVendasCompras.frx":0414
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2355
      Width           =   7965
      _ExtentX        =   14049
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
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   0
      Picture         =   "frmVendasCompras.frx":0966
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7995
   End
End
Attribute VB_Name = "frmVendasCompras"
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




Private Sub Form_Activate()
'
'    txtDescricao.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 8085
    Me.Height = 3105
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmVendasCompras = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendasCompras = Nothing
    MenuPrincipal.AbilidataMenu
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
        If VarType(Tabela("saldo")) <> vbNull Then txtSaldo.text = Format(Tabela("saldo"), "###,##0.000") Else txtSaldo.text = "0"
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("controlar_saldo")

        txtQuantidade.SetFocus

    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double

    confirma = MsgBox("Confirma o item ", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        If txtid_venda.text = "" Then

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

                mSaldo = txtSaldo.text - txtQuantidade.text
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
            Scampo = Scampo & ", '" & txtid_venda.text & "'"

            campo = campo & ", quantidade"
            Scampo = Scampo & ", '" & FormatValor(txtQuantidade.text, 1) & "'"

            campo = campo & ", preco_venda"
            Scampo = Scampo & ", '" & FormatValor(lblPrecovenda.Caption, 1) & "'"

            campo = campo & ", datasaida"
            Scampo = Scampo & ", '" & Format(Now, "YYYYMMDD") & "'"

            sqlIncluir "Saida", campo, Scampo, Me, "N"

        End If

        With frmVendas
            .lblid_venda.Caption = txtid_venda.text
            .lbldataVenda.Caption = Format(Now, "DD/MM/YYYY")
            .listaProdutosCompras
        End With

    End If

    txtDescricao.text = ""
    txtid_estoque.text = ""
    txtQuantidade.text = ""
    lblPrecovenda.Caption = ""
    lblTotal.Caption = ""

    txtDescricao.SetFocus

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
        If VarType(Tabela("maxid")) <> vbNull Then txtid_venda.text = Tabela("maxid") Else txtid_venda.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdGravar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar_Click
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
        If txtDescricao.text = "0" Then
            cmdConsultar_Click
        Else
            txtQuantidade.SetFocus
        End If
    End If
    If KeyAscii = vbKeyEscape Then cmdSair_Click
End Sub

'--- txtQuantidade
Private Sub txtQuantidade_GotFocus()
    txtQuantidade.BackColor = &H80FFFF
End Sub
Private Sub txtQuantidade_LostFocus()
    txtQuantidade.BackColor = &H80000014
    txtQuantidade.text = Format(txtQuantidade.text, "###,##0.00")
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    Dim Gtotal As Double
    If KeyAscii = vbKeyReturn Then
        If txtQuantidade.text <> "" Then
            Gtotal = txtQuantidade.text * lblPrecovenda.Caption
            lblTotal.Caption = Format(Gtotal, "###,##0.00")
            cmdGravar.SetFocus
        End If
    End If

End Sub

