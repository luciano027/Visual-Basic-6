VERSION 5.00
Begin VB.Form frmCaixaDevolucao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa Devolução"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsaldo 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtpreco_venda 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtChave 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6255
      Begin VB.TextBox txtHistorico 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         TabIndex        =   2
         Text            =   "1,00"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Image cmdConsultar 
         Height          =   375
         Left            =   5760
         Picture         =   "frmCaixaDevolucao.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Devolução"
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
         Width           =   6255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1500
      End
   End
   Begin VB.PictureBox status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6540
      TabIndex        =   0
      Top             =   3315
      Width           =   6600
   End
   Begin Vendas.VistaButton cmdsair 
      Height          =   615
      Left            =   5400
      TabIndex        =   11
      Top             =   2520
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
      Picture         =   "frmCaixaDevolucao.frx":030A
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
      Left            =   3720
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Devolver"
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
      Picture         =   "frmCaixaDevolucao.frx":0414
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor a Devolver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblDevolver 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   0
      Picture         =   "frmCaixaDevolucao.frx":0E86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "frmCaixaDevolucao"
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
        If VarType(Tabela("descricao")) <> vbNull Then txtHistorico.text = Tabela("descricao") Else txtHistorico.text = ""
        If VarType(Tabela("preco_venda")) <> vbNull Then txtpreco_venda.text = Format(Tabela("preco_venda"), "###,##0.00") Else txtpreco_venda.text = ""
        If VarType(Tabela("saldo")) <> vbNull Then txtsaldo.text = Format(Tabela("saldo"), "###,##0.000") Else txtsaldo.text = "0"
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("controlar_saldo")


        lblDevolver.Caption = Format(txtpreco_venda.text * txtQuantidade.text, "###,##0.00")

        'txtQuantidade.SetFocus
        ChaveM = "N"

    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Form_Activate()
    txtHistorico.SetFocus
    txtData.text = Format(Date, "DD/MM/YYYY")
End Sub

Private Sub Form_Load()
    Me.Width = 6720
    Me.Height = 4080

    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmCaixaDevolucao = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCaixaDevolucao = Nothing
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Dim mHistorico As String
    Dim table As String

    mHistorico = Mid("Dev:" & txtHistorico.text, 1, 100)

    If txtHistorico.text <> "" Then
        campo = " historico"
        Scampo = "'" & mHistorico & "'"
    Else
        MsgBox ("Historico não pode ficar em branco..")
        txtHistorico.SetFocus
        Exit Sub
    End If

    If lblDevolver.Caption <> "0,00" Then
        campo = campo & ", valorcaixadinheiro"
        Scampo = Scampo & ", '-" & FormatValor(lblDevolver.Caption, 1) & "'"
    Else
        MsgBox ("Valor não pode ficar em branco..")
        txtHistorico.SetFocus
        Exit Sub
    End If

    campo = campo & ", id_vendedor"
    Scampo = Scampo & ", ' " & txtid_vendedor.text & "'"

    campo = campo & ", datacaixa"
    Scampo = Scampo & ", '" & Format(txtData.text, "YYYYMMDD") & "'"

    campo = campo & ", status"
    Scampo = Scampo & ", 'P'"

    ' Incluir valor na tabela desc_gru
    sqlIncluir "caixa", campo, Scampo, Me, "S"

    Devolucao

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Devolucao()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim mSaldo As Double
    Dim mQuantidade As Double
    Dim mControlarSaldo As String
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "select estoquesaldo.id_estoque, estoquesaldo.saldo, "
    Sql = Sql & " estoques.controlar_saldo, estoques.id_estoque"
    Sql = Sql & " from"
    Sql = Sql & " estoques"
    Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
    Sql = Sql & " where"
    Sql = Sql & " estoques.id_estoque = '" & txtid_estoque.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("controlar_saldo")) <> vbNull Then mControlarSaldo = Tabela("Controlar_saldo")
        If VarType(Tabela("saldo")) <> vbNull Then mSaldo = Tabela("saldo")

        If mControlarSaldo = "S" Then
            If VarType(Tabela("saldo")) = vbNull Then
                mSaldo = 0
                mSaldo = mSaldo + txtQuantidade.text

                campo = "saldo"
                Scampo = "'" & FormatValor(mSaldo, 1) & "'"

                campo = campo & ", id_estoque"
                Scampo = Scampo & ", '" & Tabela("id_estoque") & "'"

                sqlIncluir "estoquesaldo", campo, Scampo, Me, "N"
            Else
                mSaldo = mSaldo + txtQuantidade.text
                campo = " saldo = '" & FormatValor(mSaldo, 1) & "'"
                Sqlconsulta = "id_estoque = '" & Tabela("id_estoque") & "'"
                sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"
            End If
        End If
    End If


    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'---------------------------------------------------------------
'----------------------- campos do formulario-------------------------------

'------------ nome
Private Sub txthistorico_GotFocus()
    txtHistorico.BackColor = &H80FFFF
End Sub
Private Sub txthistorico_LostFocus()
    txtHistorico.BackColor = &H80000014
    If Len(txtHistorico.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtHistorico.text)
        txtHistorico.SetFocus
    End If
End Sub
Private Sub txthistorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtQuantidade.SetFocus
End Sub


'--- txtquantidade
Private Sub txtquantidade_GotFocus()
    txtQuantidade.BackColor = &H80FFFF
End Sub
Private Sub txtquantidade_LostFocus()
    txtQuantidade.BackColor = &H80000014
    txtQuantidade.text = Format(txtQuantidade.text, "###,##0.00")
    lblDevolver.Caption = Format(txtpreco_venda.text * txtQuantidade.text, "###,##0.00")
End Sub
Private Sub txtquantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub














