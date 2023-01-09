VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaAlterarItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Cliente"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   7335
      Begin VB.TextBox txtid_estoque 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   150
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
         Left            =   240
         TabIndex        =   14
         Top             =   435
         Width           =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item do Estoque"
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
         TabIndex        =   16
         Top             =   0
         Width           =   7335
      End
      Begin VB.Image cmdConsultar 
         Height          =   375
         Left            =   6840
         Picture         =   "frmFichaAlterarItem.frx":0000
         Stretch         =   -1  'True
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label10 
         Caption         =   "Descrição"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtValorCompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Width           =   1935
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
         TabIndex        =   8
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor (R$)"
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
         Width           =   2175
      End
   End
   Begin VB.TextBox txtid_prazoitem 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.Label lblProduto 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label4 
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
         TabIndex        =   3
         Top             =   -1560
         Width           =   7695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descrição"
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
         TabIndex        =   2
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label lblCliente 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   6360
      TabIndex        =   9
      Top             =   2640
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
      Picture         =   "frmFichaAlterarItem.frx":030A
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
      TabIndex        =   10
      Top             =   3480
      Width           =   7665
      _ExtentX        =   13520
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
   Begin Vendas.VistaButton cmdGravar 
      Height          =   615
      Left            =   5160
      TabIndex        =   11
      Top             =   2640
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
      Picture         =   "frmFichaAlterarItem.frx":0414
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   0
      Picture         =   "frmFichaAlterarItem.frx":0966
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7755
   End
End
Attribute VB_Name = "frmFichaAlterarItem"
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
Dim mID_prazo As String
Dim mDinheiro As Double
Dim mCartao As Double


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
        If VarType(Tabela("preco_venda")) <> vbNull Then txtValorCompra.text = Format(Tabela("preco_venda"), "###,##0.00") Else txtValorCompra.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)


End Sub

Private Sub Form_Activate()
    txtValorCompra.SetFocus
End Sub

Private Sub Form_Load()
    Me.Width = 7755
    Me.Height = 4230
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaAlterarItem = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFichaAlterarItem = Nothing
End Sub


Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    If txtValorCompra.text <> "" Then

        campo = "preco_venda = '" & FormatValor(txtValorCompra.text, 1) & "'"
        campo = campo & ", id_estoque = '" & txtid_estoque.text & "'"

        Sqlconsulta = " prazoitem.id_prazoitem = '" & txtid_prazoitem.text & "'"

        sqlAlterar "prazoitem", campo, Sqlconsulta, Me, "N"

        MsgBox ("Valor alterar com sucesso..."), vbInformation

        Unload Me

    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


'--- txtvalorcompra
Private Sub txtvalorcompra_GotFocus()
    txtValorCompra.BackColor = &H80FFFF
End Sub
Private Sub txtvalorcompra_LostFocus()
    txtValorCompra.BackColor = &H80000014
    txtValorCompra.text = Format(txtValorCompra.text, "###,##0.00")
End Sub
Private Sub txtvalorcompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub


