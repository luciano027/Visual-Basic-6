VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendasCancelarItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar Item Vendas"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSaldo 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_venda 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
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
         Width           =   7455
      End
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
         TabIndex        =   2
         Top             =   960
         Width           =   2415
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
         Left            =   -1560
         TabIndex        =   9
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label1 
         Caption         =   "Quantidade"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
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
         TabIndex        =   5
         Top             =   960
         Width           =   2415
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
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Preço Venda"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox txtid_saida 
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
      TabIndex        =   12
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
      Picture         =   "frmVendasCancelarItem.frx":0000
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
      TabIndex        =   13
      Top             =   2430
      Width           =   7980
      _ExtentX        =   14076
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
   Begin Vendas.VistaButton cmdExcluir 
      Height          =   615
      Left            =   5520
      TabIndex        =   14
      Top             =   1680
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
      Picture         =   "frmVendasCancelarItem.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   0
      Picture         =   "frmVendasCancelarItem.frx":065C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7995
   End
End
Attribute VB_Name = "frmVendasCancelarItem"
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
    AtualizaCadastro
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
    Set frmVendasCancelarItem = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendasCancelarItem = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub AtualizaCadastro()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = " SELECT saida.id_venda, saida.id_estoque, saida.quantidade, saida.preco_venda,"
    Sql = Sql & " estoques.id_estoque, estoques.descricao, estoques.unidade, estoquesaldo.saldo,"
    Sql = Sql & " (saida.quantidade*saida.preco_venda) AS total,"
    Sql = Sql & " estoques.controlar_saldo"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " left join estoquesaldo on saida.id_estoque = estoquesaldo.id_estoque"
    Sql = Sql & " where "
    Sql = Sql & " saida.id_saida = '" & txtid_saida.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("descricao")) <> vbNull Then txtDescricao.text = Tabela("descricao") Else txtDescricao.text = ""
        If VarType(Tabela("quantidade")) <> vbNull Then txtQuantidade.text = Format(Tabela("quantidade"), "###,##0.000") Else txtQuantidade.text = ""
        If VarType(Tabela("saldo")) <> vbNull Then txtsaldo.text = Format(Tabela("saldo"), "###,##0.000") Else txtsaldo.text = ""
        If VarType(Tabela("preco_venda")) <> vbNull Then lblPrecovenda.Caption = Format(Tabela("preco_venda"), "###,##0.00") Else lblPrecovenda.Caption = ""
        If VarType(Tabela("id_estoque")) <> vbNull Then txtid_estoque.text = Tabela("id_estoque") Else txtid_estoque.text = ""
        If VarType(Tabela("total")) <> vbNull Then lblTotal.Caption = Format(Tabela("total"), "###,##0.00") Else lblTotal.Caption = ""
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("controlar_saldo")
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim mQuantidade As Double

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")



    confirma = MsgBox("Confirma a exclusão do item ", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        Sqlconsulta = "id_saida = '" & txtid_saida.text & "'"
        sqlDeletar "saida", Sqlconsulta, Me, "N"

        If mControlarSaldo = "S" Then

            Sql = " select estoquesaldo.saldo, estoquesaldo.id_estoque"
            Sql = Sql & " From"
            Sql = Sql & " estoquesaldo"
            Sql = Sql & " where "
            Sql = Sql & " estoquesaldo.id_estoque = '" & txtid_estoque.text & "'"

            If Tabela.State = 1 Then Tabela.Close
            Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Tabela.RecordCount > 0 Then

                Sqlconsulta = "id_estoque = '" & txtid_estoque.text & "'"

                mSaldo = txtsaldo.text
                mQuantidade = txtQuantidade.text

                mSaldo = mSaldo + mQuantidade
                '-------------- Alteara saldo
                campo = "saldo = '" & FormatValor(mSaldo, 1) & "'"
                sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"

            Else

                campo = "id_estoque"
                Scampo = "'" & txtid_estoque.text & "'"

                campo = campo & ", saldo"
                Scampo = Scampo & ", '" & FormatValor(txtQuantidade.text, 1) & "'"

                sqlIncluir "Estoquesaldo", campo, Scampo, Me, "N"

            End If
        End If

        Unload Me

    End If


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub






