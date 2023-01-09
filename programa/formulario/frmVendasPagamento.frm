VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendasPagamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendas Pagamento"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   3135
      Begin VB.TextBox txtCaixaBoleto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtCartao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtDinheiro 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Boleto"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblDesconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cartão"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma de Pagamento (R$)"
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
         TabIndex        =   15
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox txtConfirmaPag 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Text            =   "N"
      Top             =   2160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtCliente 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPrazo 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_prazo 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      Begin VB.Label lblTotalVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total da Venda (R$)"
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
         Width           =   4815
      End
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   4560
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
      Picture         =   "frmVendasPagamento.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdAPrazo 
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "A Prazo"
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
      Picture         =   "frmVendasPagamento.frx":010A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   4920
      _ExtentX        =   8678
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
   Begin Vendas.VistaButton cmdAvista 
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Pagamento"
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
      Picture         =   "frmVendasPagamento.frx":0A58
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtid_venda 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin Vendas.VistaButton cmdExtrato 
      Height          =   615
      Left            =   3360
      TabIndex        =   22
      ToolTipText     =   "Extrato da Venda"
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Extrato"
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
      Picture         =   "frmVendasPagamento.frx":14CA
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.TextBox txtExtrato 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtVendedor 
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtObservacao 
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   0
      Picture         =   "frmVendasPagamento.frx":15DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5115
   End
End
Attribute VB_Name = "frmVendasPagamento"
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
Dim mDinheiro As Double
Dim mCartao As Double
Dim mDesconto As Double
Dim mBoleto As Double
Dim mTotal As Double
Dim mTotalG As Double
Dim mTotalD As Double


Private Sub Buscar_id()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT max(id_prazo) as MaxID "
    Sql = Sql & " FROM prazo"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("maxid")) <> vbNull Then txtid_prazo.text = Tabela("maxid") Else txtid_prazo.text = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


Private Sub cmdAPrazo_Click()
    On Error GoTo trata_erro

    If txtid_cliente.text <> "" Then
        With frmVendasPagPrazo
            .txtid_cliente.text = txtid_cliente.text
            .txtid_vendedor.text = txtid_vendedor.text
            .txtid_venda.text = txtid_venda.text
            .txtApagar.text = lblTotalVenda.Caption
            .txtCliente.text = txtCliente.text
            .Show 1
        End With

    Else
        MsgBox ("Favor selecionar um cliente..."), vbInformation
        Unload Me
    End If

    If txtConfirmaPag.text = "S" Then Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdExtrato_Click()
    ClienteNome = txtCliente.text
    VendasExtrato
End Sub

Private Sub Form_Activate()
'If txtPrazo.text = "S" Then cmdAPrazo.Enabled = True Else cmdAPrazo.Enabled = False
    txtDinheiro.SetFocus

End Sub

Private Sub Form_Load()
    Me.Width = 5010
    Me.Height = 6090
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmVendasPagamento = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendasPagamento = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdAvista_Click()
    On Error GoTo trata_erro

    If txtDinheiro.text <> "" Then mDinheiro = txtDinheiro.text Else mDinheiro = 0
    If txtCartao.text <> "" Then mCartao = txtCartao.text Else mCartao = 0
    If txtCaixaBoleto.text <> "" Then mBoleto = txtCaixaBoleto.text Else mBoleto = 0

    If lblDesconto.Caption <> "" Then mDesconto = lblDesconto.Caption Else mDesconto = 0
    mTotal = (mDinheiro + mCartao)
    mTotalG = lblTotalVenda.Caption
    mTotalD = mTotal - mDesconto

    If mTotal > mTotalG Or mDesconto = mTotalG Then
        MsgBox ("Valores não estão corretos..."), vbInformation
        txtDinheiro.text = lblTotalVenda.Caption
        txtCartao.text = ""
        lblDesconto.Caption = ""
        Exit Sub
    End If

    confirma = MsgBox("Confirma Pagamento ", vbQuestion + vbYesNo, "Incluir")
    If confirma = vbYes Then

        Sqlconsulta = "id_venda = '" & txtid_venda.text & "'"

        '--------------------------- Vendas
        campo = "status = 'P'"
        If mDinheiro > 0 Then
            campo = campo & ", dinheiro = '" & FormatValor(mDinheiro, 1) & "'"
        End If

        If mCartao > 0 Then
            campo = campo & ", cartao = '" & FormatValor(mCartao, 1) & "'"
        End If

        If mDesconto > 0 Then
            campo = campo & ", Desconto = '" & FormatValor(mDesconto, 1) & "'"
        End If

        If mBoleto > 0 Then
            campo = campo & ", boleto = '" & FormatValor(mBoleto, 1) & "'"
        End If

        campo = campo & ", id_vendedor = '" & txtid_vendedor.text & "'"

        sqlAlterar "Vendas", campo, Sqlconsulta, Me, "N"

        '----------------------------- Saida

        campo = "id_vendedor = '" & txtid_vendedor.text & "'"
        sqlAlterar "saida", campo, Sqlconsulta, Me, "N"

        '----------------------------- Caixa
        campo = "id_venda"
        Scampo = "'" & txtid_venda.text & "'"

        campo = campo & ", id_vendedor"
        Scampo = Scampo & ", '" & txtid_vendedor.text & "'"

        campo = campo & ", historico"
        Scampo = Scampo & ", 'Venda a Vista'"

        campo = campo & ", datacaixa"
        Scampo = Scampo & ", '" & Format(Now, "YYYYMMDD") & "'"

        If mDinheiro > 0 Then
            campo = campo & ", valorcaixadinheiro"
            Scampo = Scampo & ", '" & FormatValor((mDinheiro), 1) & "'"
        End If

        If mCartao > 0 Then
            campo = campo & ", valorcaixacartao"
            Scampo = Scampo & ", '" & FormatValor(mCartao, 1) & "'"
        End If

        If mDesconto > 0 Then
            campo = campo & ", valorcaixadesconto"
            Scampo = Scampo & ", '" & FormatValor(mDesconto, 1) & "'"
        End If

        If mBoleto > 0 Then
            campo = campo & ", valorcaixaboleto"
            Scampo = Scampo & ", '" & FormatValor(mBoleto, 1) & "'"
        End If

        campo = campo & ", status"
        Scampo = Scampo & ", 'A'"

        sqlIncluir "caixa", campo, Scampo, Me, "N"

        frmVendas.txtStatus.text = "P"

        ' MsgBox ("Pagamento efetuado com sucesso.."), vbInformation

        Unload Me

    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub VendasExtrato()

    On Error GoTo trata_erro
    Dim intLinhaInicial As Integer
    Dim intLinhafinal As Integer
    Dim intNroPagina As Integer
    Dim intX As Integer
    Dim strCustFileName As String
    Dim strBackSlash As String
    Dim intCustFileNbr As Integer

    Dim strFirstName As String
    Dim strLastName As String
    Dim strAddr As String
    Dim strCity As String
    Dim strState As String
    Dim strZip As String
    Dim strVendanro As String

    Dim mDebito As Double
    Dim mCredito As Double
    Dim mAPagar As Double
    Dim mDaDos As Integer
    Dim mCabecarioDados As String
    Dim mArquivo As String

    Dim strCliente As String

    Dim bRet As Boolean


    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = " SELECT saida.*, estoques.id_estoque, estoques.unidade, estoques.descricao,estoques.codigo_est,"
    Sql = Sql & " (saida.quantidade * saida.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " saida"
    Sql = Sql & " LEFT JOIN estoques ON saida.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where saida.id_venda = '" & txtid_venda.text & "'"
    Sql = Sql & " order by estoques.descricao"

    ' ------------------- Verificar a existencia do arquivo ------------------------
    mArquivo = Dir(strgExtrato)
    If mArquivo = "maq" & MicroBD & ".txt" Then
        Kill (strgExtrato)
    End If
    ' ---------------------------------------------------------------------------------

    intLinhaInicial = 1
    intLinhafinal = 1  '-----> 19
    intNroPagina = 0

    strCliente = ClienteNome
    strVendanro = txtid_venda.text

    txtExtrato.text = ""

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not Tabela.EOF
        ' read and print all the records in the input file
        If intLinhaInicial = 1 Then
            GoSub Cabecario
            intLinhafinal = 1
            intLinhaInicial = 1
        End If
        If intLinhafinal = 16 Then
            GoSub Cabecario
            intLinhafinal = 1
            intLinhaInicial = 1
        End If
        ' print a line of data
        mDaDos = 38 - Len(Tabela("descricao"))
        If mDaDos < 0 Then mDaDos = Len(Tabela("descricao")) - 38
        txtExtrato.text = txtExtrato.text & Mid(Tabela("codigo_est"), 1, 6) & Space(1)
        txtExtrato.text = txtExtrato.text & Mid(Tabela("descricao"), 1, 40) & Space(mDaDos) & Space(3)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("quantidade"), "###,##0.00"), 10) & Space(1)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("Preco_venda"), "###,##0.00"), 10) & Space(1)
        txtExtrato.text = txtExtrato.text & Alinhar(Format(Tabela("total"), "###,##0.00"), 10) & vbCrLf

        mDebito = mDebito + Tabela("total")

        intLinhaInicial = intLinhaInicial + 1
        intLinhafinal = intLinhafinal + 1

        If intLinhaInicial = 15 Then intLinhaInicial = 1

        Tabela.MoveNext
    Wend


    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    GoSub Rodape


    GeraTXT (strgExtrato)

    Call Shell(ArqImprime, vbNormalFocus)


    Exit Sub


    ' internal subroutine to print report headings
    '------------
Cabecario:
    If intNroPagina > 0 Then
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & String(80, "-") & vbCrLf
        txtExtrato.text = txtExtrato.text & "                                                                   continua..." & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf

    Else
        txtExtrato.text = txtExtrato.text & Chr(27) & Chr(120) & Chr(48) & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf
        txtExtrato.text = txtExtrato.text & vbCrLf

    End If
    ' increment the page counter
    intNroPagina = intNroPagina + 1

    ' Print 4 blank lines, which provides a for top margin. These four lines do NOT
    ' count toward the limit of 60 lines.

    txtExtrato.text = txtExtrato.text & "Papelaria" & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & ".........................................................." & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "dd/mm/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "Vendedor: " & txtVendedor.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Venda...: " & strVendanro & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Codigo Descricao                                Quant.     Preco      Total  " & vbCrLf
    txtExtrato.text = txtExtrato.text & "------ ---------------------------------------- ---------- ---------- ----------" & vbCrLf

    Return

Rodape:

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    mAPagar = mDebito - mCredito


    txtExtrato.text = txtExtrato.text & Space(51) & " Total a Pagar..R$ " & Alinhar(Format(mAPagar, "###,##0.00"), 10) & vbCrLf


    Do While intLinhafinal < 19
        txtExtrato.text = txtExtrato.text & vbCrLf
        intLinhafinal = intLinhafinal + 1
    Loop



    Return

RodapeObs:

    intNroPagina = intNroPagina + 1


    txtExtrato.text = txtExtrato.text & "Papelaria " & Space(39) & "Pagina:" & Format$(intNroPagina, "@@@") & vbCrLf
    txtExtrato.text = txtExtrato.text & "........................................." & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "mm/dd/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & strCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & "Vendedor: " & txtVendedor.text & vbCrLf & vbCrLf & vbCrLf

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    txtExtrato.text = txtExtrato.text & "Observação " & vbCrLf
    txtExtrato.text = txtExtrato.text & txtObservacao.text & vbCrLf

    Return



    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub


Private Sub GeraTXT(ByVal strFile As String)
    Dim FSO As New FileSystemObject
    Dim iARQ As TextStream

    On Error GoTo Erro_GeraLog

    Set FSO = New FileSystemObject

    If Dir$(strgExtrato) <> vbNullString Then
        Set iARQ = FSO.OpenTextFile(strFile, ForAppending)
    Else
        Set iARQ = FSO.CreateTextFile(strFile, False)
    End If

    iARQ.WriteLine txtExtrato.text

    Set FSO = Nothing
    Set iARQ = Nothing

    Exit Sub

Erro_GeraLog:
    MsgBox "Erro ao gerar !", vbCritical, "ERRO!"
    ' FinalizaAtualizador
End Sub

Public Function ImprimeArq(ByVal sArq As String) As Boolean
    Dim lArq As Long
    Dim sTexto As String

    If Dir$(sArq) = "" Then
        ImprimeArq = False
        Exit Function
    End If
    lArq = FreeFile()
    Open sArq For Binary Access Read As lArq
    sTexto = Space$(LOF(lArq))
    Get #lArq, , sTexto
    Close lArq
    Printer.Print sTexto
    Printer.EndDoc
    ImprimeArq = True
End Function

Private Function Alinhar(texto As String, Largura As Integer)
    Alinhar = String(Largura - Len(texto), " ") & texto
End Function


'-------- Dinheiro
Private Sub txtDinheiro_GotFocus()
    txtDinheiro.BackColor = &H80FFFF
    If txtid_cliente.text <> "" Then cmdAvista.Enabled = False
    If txtid_cliente.text = "" Then cmdAPrazo.Enabled = True
End Sub
Private Sub txtDinheiro_LostFocus()
    txtDinheiro.BackColor = &H80000014
    txtDinheiro.text = Format(txtDinheiro.text, "###,##0.00")
End Sub
Private Sub txtDinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtDinheiro.text <> "" Then mDinheiro = txtDinheiro.text Else mDinheiro = 0
        If txtCartao.text <> "" Then mCartao = txtCartao.text Else mCartao = 0
        mTotalG = lblTotalVenda.Caption
        lblDesconto.Caption = Format(mTotalG - (mDinheiro + mCartao), "###,##0.00")
        txtCartao.SetFocus
    End If
End Sub

'-------- Cartao
Private Sub txtCartao_GotFocus()
    txtCartao.text = ""
    txtCartao.BackColor = &H80FFFF
End Sub
Private Sub txtCartao_LostFocus()
    txtCartao.BackColor = &H80000014
    txtCartao.text = Format(txtCartao.text, "###,##0.00")
End Sub
Private Sub txtCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtDinheiro.text <> "" Then mDinheiro = txtDinheiro.text Else mDinheiro = 0
        If txtCartao.text <> "" Then mCartao = txtCartao.text Else mCartao = 0
        mTotalG = lblTotalVenda.Caption
        lblDesconto.Caption = Format(mTotalG - (mDinheiro + mCartao), "###,##0.00")
        txtCaixaBoleto.SetFocus
    End If
End Sub

'-------- Boleto
Private Sub txtcaixaboleto_GotFocus()
    txtCaixaBoleto.text = ""
    txtCaixaBoleto.BackColor = &H80FFFF
End Sub
Private Sub txtcaixaboleto_LostFocus()
    txtCaixaBoleto.BackColor = &H80000014
    If txtCaixaBoleto.text = "" Then
        mBoleto = 0
        txtCaixaBoleto.text = "0,00"
    Else
        txtCaixaBoleto.text = Format(txtCaixaBoleto.text, "###,##0.00")
        mBoleto = txtCaixaBoleto.text
    End If
End Sub
Private Sub txtcaixaboleto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mBoleto = txtCaixaBoleto.text
        If txtDinheiro.text <> "" Then mDinheiro = txtDinheiro.text Else mDinheiro = 0
        If txtCartao.text <> "" Then mCartao = txtCartao.text Else mCartao = 0
        mTotalG = lblTotalVenda.Caption
        lblDesconto.Caption = Format(mTotalG - (mDinheiro + mCartao + mBoleto), "###,##0.00")
        cmdAvista_Click
    End If
End Sub



