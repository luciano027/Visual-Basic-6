VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoqueCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estoque"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1050
      Left            =   9480
      TabIndex        =   22
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativo"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
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
         Width           =   1455
      End
   End
   Begin VB.TextBox txtChave 
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_login 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   9135
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Preço de Compra"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblPreco_compra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblQuant_venda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   37
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade Vendida"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   36
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblData_venda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data da Ultima Venda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblSaldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Atual"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblquant_compra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade Comprada"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbldata_compra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data da Ultima Compra"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Outras informações do produto"
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
         TabIndex        =   13
         Top             =   0
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtcodigo_est 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         TabIndex        =   40
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkControlarSaldo 
         Caption         =   "Controlar Saldo"
         Height          =   255
         Left            =   7560
         TabIndex        =   26
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtSaldo_minimo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtid_estoque 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtPreco_venda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtUnidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8400
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Refêrencia"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Minimo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Estoque"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Venda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endereço"
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
         Width           =   9135
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   3645
      Width           =   11130
      _ExtentX        =   19632
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
      Left            =   9480
      TabIndex        =   19
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "frmEstoqueCadastro.frx":0000
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
      Left            =   9480
      TabIndex        =   20
      Top             =   2040
      Width           =   1455
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
      Picture         =   "frmEstoqueCadastro.frx":010A
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
      Left            =   9480
      TabIndex        =   21
      Top             =   1320
      Width           =   1455
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
      Picture         =   "frmEstoqueCadastro.frx":065C
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   0
      Picture         =   "frmEstoqueCadastro.frx":0BAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14955
   End
End
Attribute VB_Name = "frmEstoqueCadastro"
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


Private Sub Form_Activate()
    txtDescricao.SetFocus
    If txtTipo.text = "A" Or txtTipo.text = "E" Then AutalizaCadastro

End Sub

Private Sub Form_Load()
    Me.Width = 11220
    Me.Height = 4395
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmEstoqueCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEstoqueCadastro = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro




    Unload Me
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    Dim table As String

    ' Rotina de gravacao de inclusao dos dados
    If txtTipo.text = "I" Then

        campo = " data_cadastro"
        Scampo = "'" & Format(Date$, "YYYYMMDD") & "'"

        If txtDescricao.text <> "" Then
            campo = campo & ", descricao"
            Scampo = Scampo & ",'" & txtDescricao.text & "'"
        Else
            MsgBox ("Descrição não pode ficar em branco..")
            txtDescricao.SetFocus
            Exit Sub
        End If
        If txtUnidade.text <> "" Then
            campo = campo & ", Unidade"
            Scampo = Scampo & ", '" & txtUnidade.text & "'"
        End If
        If txtSaldo_minimo.text <> "" Then
            campo = campo & ", saldo_minimo"
            Scampo = Scampo & ", '" & FormatValor(txtSaldo_minimo.text, 1) & "'"
        End If
        If txtpreco_venda.text <> "" Then
            campo = campo & ", preco_venda"
            Scampo = Scampo & ", '" & FormatValor(txtpreco_venda.text, 1) & "'"
        End If
        If chkControlarSaldo.Value = 1 Then
            campo = campo & ", controlar_saldo"
            Scampo = Scampo & ", 'S'"
        Else
            campo = campo & ", controlar_saldo"
            Scampo = Scampo & ", 'N'"
        End If

        If optAtivo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'A'"
        End If

        If optInativo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'I'"
        End If

        If txtcodigo_est.text <> "" Then
            campo = campo & ", codigo_est"
            Scampo = Scampo & ", '" & txtcodigo_est.text & "'"
        Else
            Buscar_id
            campo = campo & ", codigo_est"
            Scampo = Scampo & ", '" & txtcodigo_est.text & "'"
        End If

        ' Incluir valor na tabela Estoques
        sqlIncluir "Estoques", campo, Scampo, Me, "S"

    End If
    ' rotina de gravacao de alteracao dos dados
    If txtTipo.text = "A" Then

        ' Consulta os dados da tabela Estoques
        Sqlconsulta = "id_estoque = '" & txtid_estoque.text & "'"

        If txtDescricao.text <> "" Then campo = " Descricao = '" & UCase(txtDescricao.text) & "'" Else txtDescricao.SetFocus
        If txtUnidade.text <> "" Then campo = campo & ", Unidade = '" & txtUnidade.text & "'"
        If txtSaldo_minimo.text <> "" Then campo = campo & ", saldo_minimo = '" & FormatValor(txtSaldo_minimo.text, 1) & "'"
        If txtpreco_venda.text <> "" Then campo = campo & ", preco_venda = '" & FormatValor(txtpreco_venda.text, 1) & "'"
        If chkControlarSaldo.Value = 1 Then campo = campo & ", controlar_saldo = 'S'" Else campo = campo & ", controlar_saldo = 'N'"
        If optAtivo.Value = True Then campo = campo & ", status = 'A'" Else campo = campo & ", status = 'I'"
        If txtcodigo_est <> "" Then campo = campo & ", codigo_est = '" & txtcodigo_est.text & "'"

        ' Aletar dos dados da tabela Estoques
        sqlAlterar "Estoques", campo, Sqlconsulta, Me, "S"

    End If

    cmdGravar.Enabled = False

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim mCodigo_est As Integer

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(codigo_est) as MaxID "
    Sql = Sql & " FROM estoques"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then mCodigo_est = Tabela("maxid") Else mCodigo_est = 0
    End If

    txtcodigo_est.text = strzero(STR(mCodigo_est + 1), 6, Vbfantes)

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub AutalizaCadastro()
    On Error GoTo trata_erro

    If txtTipo.text = "A" Or txtTipo.text = "E" Then
        If txtChave.text = "0" Then
            Dim Estoques As ADODB.Recordset
            ' conecta ao banco de dados

            Set Estoques = CreateObject("ADODB.Recordset")    '''

            ' abre um Recrodset da Tabela Estoques
            Sql = " SELECT estoques.*, fornecedores.id_fornecedor, fornecedores.fornecedor,"
            Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo"
            Sql = Sql & " From"
            Sql = Sql & " estoques"
            Sql = Sql & " LEFT JOIN fornecedores ON estoques.id_fornecedor = fornecedores.id_fornecedor"
            Sql = Sql & " left join estoquesaldo on estoques.id_estoque = estoquesaldo.id_estoque"
            Sql = Sql & " where "
            Sql = Sql & " estoques.id_estoque = '" & txtid_estoque.text & "'"

            If Estoques.State = 1 Then Estoques.Close
            Estoques.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Estoques.RecordCount > 0 Then

                If VarType(Estoques("Descricao")) <> vbNull Then txtDescricao.text = Estoques("Descricao") Else txtDescricao.text = ""
                If VarType(Estoques("Unidade")) <> vbNull Then txtUnidade.text = Estoques("Unidade") Else txtUnidade.text = ""
                If VarType(Estoques("saldo_minimo")) <> vbNull Then txtSaldo_minimo.text = Format(Estoques("saldo_minimo"), "###,##0.000") Else txtSaldo_minimo.text = ""
                If VarType(Estoques("preco_venda")) <> vbNull Then txtpreco_venda.text = Format(Estoques("preco_venda"), "###,##0.00") Else txtpreco_venda.text = ""

                If VarType(Estoques("controlar_saldo")) <> vbNull Then
                    If Estoques("controlar_saldo") = "S" Then chkControlarSaldo.Value = 1
                    If Estoques("controlar_saldo") = "N" Then chkControlarSaldo.Value = 0
                End If

                If VarType(Estoques("status")) <> vbNull Then
                    If Estoques("status") = "A" Then optAtivo.Value = True
                    If Estoques("status") = "I" Then optInativo.Value = True
                End If

                If VarType(Estoques("fornecedor")) <> vbNull Then lblFornecedor.Caption = Estoques("fornecedor") Else lblFornecedor.Caption = ""
                If VarType(Estoques("data_compra")) <> vbNull Then lbldata_compra.Caption = Format(Estoques("data_compra"), "dd/mm/yyyy") Else lbldata_compra.Caption = ""
                If VarType(Estoques("quant_compra")) <> vbNull Then lblquant_compra.Caption = Format(Estoques("quant_compra"), "###,##0.000") Else lblquant_compra.Caption = ""
                If VarType(Estoques("data_venda")) <> vbNull Then lblData_venda.Caption = Format(Estoques("data_venda"), "dd/mm/yyyy") Else lblData_venda.Caption = ""
                If VarType(Estoques("quant_venda")) <> vbNull Then lblQuant_venda.Caption = Format(Estoques("quant_venda"), "###,##0.00") Else lblQuant_venda.Caption = ""
                If VarType(Estoques("preco_compra")) <> vbNull Then lblPreco_compra.Caption = Format(Estoques("preco_compra"), "###,##0.00") Else lblPreco_compra.Caption = ""
                If VarType(Estoques("saldo")) <> vbNull Then lblSaldo.Caption = Format(Estoques("saldo"), "###,##0.000") Else lblSaldo.Caption = ""
                If VarType(Estoques("codigo_est")) <> vbNull Then txtcodigo_est.text = Estoques("codigo_est") Else txtcodigo_est.text = ""

            End If
            If Estoques.State = 1 Then Estoques.Close
            Set Estoques = Nothing

            If txtTipo.text = "E" Then cmdGravar.Enabled = False
            If txtTipo.text = "A" Then cmdExcluir.Enabled = False
            txtChave.text = "1"

        End If
    End If

    txtDescricao.SetFocus
    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'---------------------------------------------------------------
'----------------------- campos do formulario-------------------------------

'------------ nome
Private Sub txtDescricao_GotFocus()
    txtDescricao.BackColor = &H80FFFF
End Sub
Private Sub txtDescricao_LostFocus()
    txtDescricao.BackColor = &H80000014
    If Len(txtDescricao.text) > 100 Then
        MsgBox "Comprimento do campo e de 100 digitos, voce digitou " & Len(txtDescricao.text)
        txtDescricao.SetFocus
    End If
End Sub
Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtUnidade.SetFocus
End Sub

'--- Unidade
Private Sub txtUnidade_GotFocus()
    txtUnidade.BackColor = &H80FFFF
End Sub
Private Sub txtUnidade_LostFocus()
    txtUnidade.BackColor = &H80000014
    If Len(txtUnidade.text) > 5 Then
        MsgBox "Comprimento do campo e de 5 digitos, voce digitou " & Len(txtUnidade.text)
        txtUnidade.SetFocus
    End If
End Sub
Private Sub txtUnidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtSaldo_minimo.SetFocus
End Sub

'---saldo_minimo
Private Sub txtsaldo_minimo_GotFocus()
    txtSaldo_minimo.BackColor = &H80FFFF
End Sub
Private Sub txtsaldo_minimo_LostFocus()
    txtSaldo_minimo.BackColor = &H80000014
    txtSaldo_minimo.text = Format(txtSaldo_minimo.text, "###,##0.00")
End Sub
Private Sub txtsaldo_minimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtpreco_venda.SetFocus
End Sub


'-------- preco_venda
Private Sub txtpreco_venda_GotFocus()
    txtpreco_venda.BackColor = &H80FFFF
End Sub
Private Sub txtpreco_venda_LostFocus()
    txtpreco_venda.BackColor = &H80000014
    txtpreco_venda.text = Format(txtpreco_venda.text, "###,##0.00")
End Sub
Private Sub txtpreco_venda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub




