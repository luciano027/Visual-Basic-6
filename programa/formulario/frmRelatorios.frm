VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelatorios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRelatorio 
      Height          =   285
      Left            =   5280
      TabIndex        =   43
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_relatorio 
      Height          =   285
      Left            =   5040
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3720
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   41
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   480
      TabIndex        =   29
      Top             =   600
      Width           =   7935
      Begin VB.OptionButton optEstoque 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Estoque"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optCaixa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Caixa"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1420
         Width           =   1455
      End
      Begin VB.OptionButton optContasPagar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contas a Pagar"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1150
         Width           =   1455
      End
      Begin VB.OptionButton optFornecedor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptVendedor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   900
         Width           =   1455
      End
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   660
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListaRelatorios 
         Height          =   1695
         Left            =   1800
         TabIndex        =   36
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2990
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblObs 
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
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   7695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Observação"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   2160
         Width           =   7935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tipo de relatórios"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   20
      Top             =   3360
      Width           =   5175
      Begin VB.TextBox txtNF 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   3600
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optGeral 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Geral"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optPagas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pagas"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optemAberto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "em Aberto"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Documento (NF)"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Historico"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contas a Pagar"
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.TextBox txtid_estoque 
      Height          =   285
      Left            =   9000
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtestoque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   17
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   3600
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtVendedor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3600
      TabIndex        =   14
      Top             =   5760
      Width           =   4935
   End
   Begin VB.TextBox txtid_Cliente 
      Height          =   285
      Left            =   9000
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_fornecedor 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7650
      Width           =   14220
      _ExtentX        =   25083
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
      Left            =   11520
      TabIndex        =   1
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "               Sair"
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
      Picture         =   "frmRelatorios.frx":0000
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
      Left            =   11160
      TabIndex        =   2
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
      Left            =   8520
      TabIndex        =   3
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
   Begin VB.TextBox txtFornecedor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   5040
      Width           =   4935
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   5040
      Width           =   4335
   End
   Begin Vendas.VistaButton cmdImprimir 
      Height          =   615
      Left            =   9240
      TabIndex        =   39
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "             Imprimir"
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
      Picture         =   "frmRelatorios.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Consultaestoque 
      Height          =   315
      Left            =   13320
      Picture         =   "frmRelatorios.frx":021C
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   360
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Produto do Estoque"
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   5520
      Width           =   4665
   End
   Begin VB.Image ConsultaVendedor 
      Height          =   315
      Left            =   8520
      Picture         =   "frmRelatorios.frx":0526
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   360
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   5520
      Width           =   5265
   End
   Begin VB.Image ConsultaCliente 
      Height          =   315
      Left            =   13320
      Picture         =   "frmRelatorios.frx":0830
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   9000
      TabIndex        =   13
      Top             =   4800
      Width           =   4665
   End
   Begin VB.Image cmdConsultaFornecedor 
      Height          =   315
      Left            =   8520
      Picture         =   "frmRelatorios.frx":0B3A
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   4800
      Width           =   5265
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
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   13215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Período"
      Height          =   255
      Left            =   10440
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbldataF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   11520
      TabIndex        =   5
      Top             =   600
      Width           =   2130
   End
   Begin VB.Label lblDataI 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   7170
      Left            =   240
      Picture         =   "frmRelatorios.frx":0E44
      Stretch         =   -1  'True
      Top             =   240
      Width           =   13605
   End
   Begin VB.Image Image1 
      Height          =   7680
      Left            =   0
      Picture         =   "frmRelatorios.frx":4080
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14280
   End
End
Attribute VB_Name = "frmRelatorios"
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
Dim Mitem As Integer


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

Private Sub cmdImprimir_Click()
    Dim mRelatorio As String

    If txtRelatorio.text = "rptCaixa" Then
        Sqlconsulta = " caixa.datacaixa Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
        If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
            With rptCaixa
                .lblConsulta.Caption = Sqlconsulta
                .lblCliente.Caption = "Período: " & Format(txtDataI.Value, "DD/MM/YYYY") & " a " & Format(txtDataF.Value, "DD/MM/YYYY")
                .Show 1
            End With
        Else
            MsgBox ("Favor selecionar um periodo..."), vbInformation
            Exit Sub
        End If
    End If

    If txtRelatorio.text = "rptCaixaResumo" Then
        Sqlconsulta = " caixa.datacaixa Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
        If lblDataI.Caption <> "" And lbldataF.Caption <> "" Then
            With rptCaixaResumo
                .lblConsulta.Caption = Sqlconsulta
                .lblCliente.Caption = "Período: " & Format(txtDataI.Value, "DD/MM/YYYY") & " a " & Format(txtDataF.Value, "DD/MM/YYYY")
                .Show 1
            End With
        Else
            MsgBox ("Favor selecionar um periodo..."), vbInformation
            Exit Sub
        End If
    End If

    If txtRelatorio.text = "rptInventario" Then

        Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_compra,"
        Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo,"
        Sql = Sql & " estoques.preco_compra * estoquesaldo.saldo AS Total"
        Sql = Sql & " From"
        Sql = Sql & " estoquesaldo"
        Sql = Sql & " LEFT JOIN estoques ON estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " where "
        Sql = Sql & " estoquesaldo.saldo > 0 "
        Sql = Sql & " order by estoques.descricao"

        With rptInventario
            .lblConsulta.Caption = Sql
            .lblUnidade.Caption = "unidade"
            .txtUnidade.DataField = "unidade"
            .Show 1
        End With
    End If

    If txtRelatorio.text = "rptEstoque" Then

        Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_compra,"
        Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo,"
        Sql = Sql & " estoques.preco_compra * estoquesaldo.saldo AS Total"
        Sql = Sql & " From"
        Sql = Sql & " estoques"
        Sql = Sql & " LEFT JOIN estoquesaldo ON estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " where estoques.status = 'A'"
        Sql = Sql & " order by estoques.descricao"

        With rptInventario
            .lblTitulo.Caption = "Produtos Cadastrados"
            .lblPreco.Caption = "Preço Compra"
            .txtPreco.DataField = "Preco_compra"
            .lblUnidade.Caption = "unidade"
            .txtUnidade.DataField = "unidade"
            .lblConsulta.Caption = Sql
            .Show 1
        End With
    End If

    If txtRelatorio.text = "rptTabelaVenda" Then

        Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_venda,"
        Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo,"
        Sql = Sql & " estoques.preco_venda * estoquesaldo.saldo AS Total"
        Sql = Sql & " From"
        Sql = Sql & " estoques"
        Sql = Sql & " LEFT JOIN estoquesaldo ON estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " where estoques.status = 'A'"
        Sql = Sql & " order by estoques.descricao"

        With rptInventario
            .lblTitulo.Caption = "Tabela de Venda"
            .txtPreco.DataField = "Preco_venda"
            .lblPreco.Caption = "Preço Venda"
            .lblUnidade.Caption = "unidade"
            .txtUnidade.DataField = "unidade"
            .lblConsulta.Caption = Sql
            .Show 1
        End With
    End If

    If txtRelatorio.text = "rptTabelaCompra" Then

        Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_compra,"
        Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo,"
        Sql = Sql & " estoques.preco_compra * estoquesaldo.saldo AS Total"
        Sql = Sql & " From"
        Sql = Sql & " estoques"
        Sql = Sql & " LEFT JOIN estoquesaldo ON estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " where estoques.status = 'A'"
        Sql = Sql & " order by estoques.descricao"

        With rptInventario
            .lblTitulo.Caption = "Tabela de Compra"
            .txtPreco.DataField = "preco_compra"
            .lblPreco.Caption = "Preço Compra"
            .lblUnidade.Caption = "unidade"
            .txtUnidade.DataField = "unidade"
            .lblConsulta.Caption = Sql
            .Show 1
        End With
    End If

    If txtRelatorio.text = "rptSaldoMinimo" Then

        Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_compra,"
        Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo, estoques.controlar_saldo, estoques.status,"
        Sql = Sql & " estoques.preco_compra * estoquesaldo.saldo AS Total, estoques.saldo_minimo"
        Sql = Sql & " From"
        Sql = Sql & " Estoques"
        Sql = Sql & " LEFT JOIN estoquesaldo ON estoques.id_estoque = estoquesaldo.id_estoque"
        Sql = Sql & " Where Estoques.saldo_minimo >= estoquesaldo.saldo"
        Sql = Sql & " and estoques.controlar_saldo = 'S'"
        Sql = Sql & " and estoques.status = 'A'"
        Sql = Sql & " order by estoques.descricao"

        With rptMinimo
            .lblTitulo.Caption = "Saldo Minimo"
            .txtPreco.DataField = "preco_compra"
            .lblUnidade.Caption = "Saldo Minimo"
            .txtUnidade.DataField = "saldo_minimo"
            .lblPreco.Caption = "Preço Compra"
            .lblConsulta.Caption = Sql
            .Show 1
        End With
    End If

    If txtRelatorio.text = "rptExtratoFornecedor" Then
        If lblDataI.Caption <> "" And lbldataF.Caption <> "" And txtid_fornecedor.text <> "" Then

            Sqlconsulta = " entrada.nfdata Between '" & Format(lblDataI.Caption, "YYYYMMDD") & "' And '" & Format(lbldataF.Caption, "YYYYMMDD") & "'"
            Sqlconsulta = Sqlconsulta & " and entrada.id_fornecedor = '" & txtid_fornecedor.text & "'"

            With rptExtratoFornecedor
                .lblCliente.Caption = "Fornecedor: " & txtFornecedor.text
                .lblPeriodo.Caption = "Período   : " & Format(txtDataI.Value, "DD/MM/YYYY") & " a " & Format(txtDataF.Value, "DD/MM/YYYY")
                .lblConsulta.Caption = Sqlconsulta
                .Show 1
            End With
        Else
            MsgBox ("Favor selecionar periodo ou vendedor..."), vbInformation
            Exit Sub
        End If
    End If

    If txtRelatorio.text = "rptSaldoMinimoCompra" Then

        If txtid_fornecedor.text <> "" Then

            Sql = " SELECT estoques.id_estoque, estoques.descricao, estoques.unidade, estoques.preco_compra, "
            Sql = Sql & " estoques.id_fornecedor, estoques.saldo_minimo,"
            Sql = Sql & " estoquesaldo.id_estoque, estoquesaldo.saldo, estoques.controlar_saldo, estoques.status,"
            Sql = Sql & " estoques.preco_compra * estoquesaldo.saldo AS Total"
            Sql = Sql & " From"
            Sql = Sql & " Estoques"
            Sql = Sql & " LEFT JOIN estoquesaldo ON estoques.id_estoque = estoquesaldo.id_estoque"
            Sql = Sql & " Where "
            Sql = Sql & " Estoques.saldo_minimo >= estoquesaldo.saldo"
            Sql = Sql & " and estoques.controlar_saldo = 'S'"
            Sql = Sql & " and estoques.status = 'A'"
            Sql = Sql & " and estoques.id_fornecedor = '" & txtid_fornecedor.text & "'"
            Sql = Sql & " order by estoques.descricao"

            With rptInventario
                .lblTitulo.Caption = "Saldo Minimo por fornecedor"
                .txtPreco.DataField = "preco_compra"
                .lblUnidade.Caption = "Saldo Minimo"
                .txtUnidade.DataField = "saldo_minimo"
                .lblPreco.Caption = "Preço Compra"
                .lblCliente.Caption = "Fornecedor: " & txtFornecedor.text
                .lblConsulta.Caption = Sql
                .Show 1
            End With
        Else
            MsgBox ("Favor selecionar um fornecedor..."), vbInformation
            Exit Sub
        End If
    End If



End Sub

Private Sub Consultaestoque_Click()
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

    Sql = "SELECT id_Estoque, descricao FROM Estoques WHERE id_Estoque = '" & txtid_estoque.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Estoque")) <> vbNull Then txtid_estoque.text = Tabela("id_Estoque") Else txtid_estoque.text = ""
        If VarType(Tabela("descricao")) <> vbNull Then txtestoque.text = Tabela("descricao") Else txtestoque.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub ConsultaVendedor_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaVendedor
        .Show 1
    End With

    If IDVendedor <> "" Then
        txtid_vendedor.text = IDVendedor
        IDVendedor = ""
    End If

    Sql = "SELECT id_Vendedor, Vendedor FROM Vendedores WHERE id_Vendedor = '" & txtid_vendedor.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Vendedor")) <> vbNull Then txtid_vendedor.text = Tabela("id_Vendedor") Else txtid_vendedor.text = ""
        If VarType(Tabela("Vendedor")) <> vbNull Then txtVendedor.text = Tabela("Vendedor") Else txtVendedor.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Me.Width = 14310
    Me.Height = 8400
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu

    txtDataI.Value = Now
    txtDataF.Value = Now

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmRelatorios = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRelatorios = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub ConsultaCliente_Click()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    With frmConsultaClientes
        .Show 1
    End With

    If IDCliente <> "" Then
        txtid_cliente.text = IDCliente
        IDCliente = ""
    End If

    Sql = "SELECT id_Cliente, Cliente FROM Clientes WHERE id_Cliente = '" & txtid_cliente.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Cliente")) <> vbNull Then txtid_cliente.text = Tabela("id_Cliente") Else txtid_cliente.text = ""
        If VarType(Tabela("Cliente")) <> vbNull Then txtCliente.text = Tabela("Cliente") Else txtCliente.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Image4_Click()

End Sub

Private Sub optCaixa_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Caixa'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub optCliente_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Cliente'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub optContasPagar_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Contas a Pagar'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub optEstoque_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Estoque'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub optFornecedor_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Fornecedor'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub




Private Sub OptVendedor_Click()
    On Error GoTo trata_erro

    Sqlconsulta = "relatorios.tipo = 'Vendedor'"

    Lista (Sqlconsulta)

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub txtDataF_DateClick(ByVal DateClicked As Date)
    lbldataF.Caption = Format(txtDataF.Value, "DD/MM/YYYY")
End Sub

Private Sub txtDataI_DateClick(ByVal DateClicked As Date)
    lblDataI.Caption = Format(txtDataI.Value, "DD/MM/YYYY")
End Sub


Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Relatorios As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Relatorios = CreateObject("ADODB.Recordset")

    Sql = " SELECT relatorios.*"
    Sql = Sql & " From"
    Sql = Sql & " relatorios"
    Sql = Sql & " Where "
    Sql = Sql & SQconsulta

    ' abre um Recrodset da Tabela Relatorios
    If Relatorios.State = 1 Then Relatorios.Close
    Relatorios.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaRelatorios.ColumnHeaders.Clear
    ListaRelatorios.ListItems.Clear

    ListaRelatorios.ColumnHeaders.Add , , "Descrição", 6000

    If Relatorios.BOF = True And Relatorios.EOF = True Then Exit Sub
    While Not Relatorios.EOF
        If VarType(Relatorios("Descricao")) <> vbNull Then Set itemx = ListaRelatorios.ListItems.Add(, , Relatorios("Descricao"))
        If VarType(Relatorios("id_relatorio")) <> vbNull Then itemx.Tag = Relatorios("id_relatorio")
        Relatorios.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaRelatorios, Picture1, vbWhite, vb3DLight, Me) = False Then Exit Sub

    If Relatorios.State = 1 Then Relatorios.Close
    Set Relatorios = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaRelatorios_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro
    Dim Relatorios As ADODB.Recordset
    ' conecta ao banco de dados
    Set Relatorios = CreateObject("ADODB.Recordset")

    txtid_relatorio.text = ListaRelatorios.SelectedItem.Tag

    Sql = " SELECT relatorios.*"
    Sql = Sql & " From"
    Sql = Sql & " relatorios"
    Sql = Sql & " Where "
    Sql = Sql & " relatorios.id_relatorio = '" & txtid_relatorio.text & "'"

    ' abre um Recrodset da Tabela Relatorios
    If Relatorios.State = 1 Then Relatorios.Close
    Relatorios.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Relatorios.RecordCount > 0 Then
        If VarType(Relatorios("observacao")) <> vbNull Then lblObs.Caption = Relatorios("observacao")
        If VarType(Relatorios("relatorio")) <> vbNull Then txtRelatorio.text = Relatorios("relatorio")
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

