VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha do Cliente"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtExtrato 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   7080
      Width           =   8535
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   7935
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
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8535
      End
      Begin VB.Image cmdConsultar 
         Height          =   315
         Left            =   8040
         Picture         =   "frmFichaConsulta.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.TextBox txtid_prazo 
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTipoc 
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   8160
      ScaleHeight     =   240
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListaFicha 
      Height          =   6615
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11668
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8370
      Width           =   12765
      _ExtentX        =   22516
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
      Left            =   11280
      TabIndex        =   8
      Top             =   7440
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
      Picture         =   "frmFichaConsulta.frx":030A
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdFicha 
      Height          =   615
      Left            =   10080
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Ficha"
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
      Picture         =   "frmFichaConsulta.frx":0414
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdExtrato 
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   7440
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmFichaConsulta.frx":0766
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Label lblCadastro 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9240
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblConsulta 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado Consulta"
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
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "frmFichaConsulta.frx":0878
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "frmFichaConsulta"
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
Dim mCliente As String
Dim strID_cliente As String

Private Sub cmdExtrato_Click()
    On Error GoTo trata_erro
    Dim PrazoPagtos As ADODB.Recordset
    Dim Prazo As ADODB.Recordset
    Dim mTotalCredito As Double
    Dim mTotalDebito As Double
    Dim mTotalPagar As Double

    If txtid_prazo.text <> "" Then
        ' conecta ao banco de dados
        Set PrazoPagtos = CreateObject("ADODB.Recordset")
        Set Prazo = CreateObject("ADODB.Recordset")
        Sql = " SELECT prazo.id_cliente, prazopagto.id_prazo, prazo.id_prazo,"
        Sql = Sql & " SUM(prazopagto.ValorPagto) AS totalcredito,"
        Sql = Sql & " clientes.id_cliente , clientes.Cliente, clientes.tel2"
        Sql = Sql & " From"
        Sql = Sql & " prazopagto"
        Sql = Sql & " LEFT JOIN prazo ON prazopagto.id_prazo = prazo.id_prazo"
        Sql = Sql & " LEFT JOIN clientes ON prazo.id_cliente  = clientes.id_cliente"
        Sql = Sql & " Where prazopagto.id_prazo = '" & txtid_prazo.text & "'"

        ' abre um Recrodset da Tabela PrazoPagtos
        If PrazoPagtos.State = 1 Then PrazoPagtos.Close
        PrazoPagtos.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If PrazoPagtos.RecordCount > 0 Then
            If VarType(PrazoPagtos("totalcredito")) <> vbNull Then mTotalCredito = PrazoPagtos("totalcredito")
            If VarType(PrazoPagtos("cliente")) <> vbNull Then mCliente = PrazoPagtos("cliente") Else mCliente = ""
            If VarType(PrazoPagtos("id_cliente")) <> vbNull Then strID_cliente = PrazoPagtos("id_cliente")
            ClienteNome = mCliente
        Else
            mTotalCredito = 0
        End If

        If mCliente = "" Then

            Sql = " SELECT prazo.id_cliente,"
            Sql = Sql & " clientes.id_cliente , clientes.Cliente, clientes.tel2"
            Sql = Sql & " From"
            Sql = Sql & " prazo"
            Sql = Sql & " LEFT JOIN clientes ON prazo.id_cliente  = clientes.id_cliente"
            Sql = Sql & " Where prazo.id_prazo = '" & txtid_prazo.text & "'"

            ' abre um Recrodset da Tabela PrazoPagtos
            If Prazo.State = 1 Then Prazo.Close
            Prazo.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Prazo.RecordCount > 0 Then
                If VarType(Prazo("id_cliente")) <> vbNull Then mCliente = Prazo("cliente") Else mCliente = ""
            End If
        End If

        If Prazo.State = 1 Then Prazo.Close
        Set Prazo = Nothing


        If PrazoPagtos.State = 1 Then PrazoPagtos.Close
        Set PrazoPagtos = Nothing


        extrato

        '  If Prazo.State = 1 Then Prazo.Close
        '  Set Prazo = Nothing
    Else
        MsgBox ("Favor selecionar um cliente..."), vbInformation
    End If

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdFicha_Click()

    With frmVendedorCodigo
        .txtTipo.text = "F"
        .txtAcesso.text = ""
        .txtid_prazo.text = txtid_prazo.text
        .Show 1
    End With

    cmdConsultar_Click

End Sub

Private Sub Form_Activate()
    frmFichaConsulta.ZOrder (0)
End Sub

Private Sub Form_Load()
    Me.Width = 12675
    Me.Height = 9090
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu

End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmFichaConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFichaConsulta = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub Lista(SQconsulta As String)
    On Error GoTo trata_erro
    Dim Ficha As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    ' conecta ao banco de dados
    Set Ficha = CreateObject("ADODB.Recordset")

    Sql = "SELECT prazo.*, clientes.id_cliente, clientes.cliente"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN clientes ON clientes.id_cliente = prazo.id_cliente"
    If SQconsulta <> "" Then Sql = Sql & " Where " & SQconsulta
    Sql = Sql & " order by clientes.cliente"

    ' abre um Recrodset da Tabela Ficha
    If Ficha.State = 1 Then Ficha.Close
    Ficha.Open Sql, banco, adOpenKeyset, adLockOptimistic

    ListaFicha.ColumnHeaders.Clear
    ListaFicha.ListItems.Clear

    If Ficha.RecordCount = 0 Then
        ' muda o curso para o normal
        lblCadastro.Caption = "Ficha(s) encontrado(s): 0"
        Exit Sub
    End If

    lblCadastro.Caption = "Ficha(s) encontrado(s): " & Ficha.RecordCount

    ListaFicha.ColumnHeaders.Add , , "Cliente ", 10250
    ListaFicha.ColumnHeaders.Add , , "Data Debito", 1550

    If Ficha.BOF = True And Ficha.EOF = True Then Exit Sub
    While Not Ficha.EOF

        If VarType(Ficha("cliente")) <> vbNull Then Set itemx = ListaFicha.ListItems.Add(, , Ficha("cliente"))
        If VarType(Ficha("data_venda")) <> vbNull Then itemx.SubItems(1) = Format(Ficha("data_venda"), "dd/mm/yyyy") Else itemx.SubItems(1) = ""
        If VarType(Ficha("id_prazo")) <> vbNull Then itemx.Tag = Ficha("id_prazo")
        Ficha.MoveNext
    Wend

    'Zebra o listview
    If LVZebra(ListaFicha, Picture1, vbWhite, &HC0FFC0, Me) = False Then Exit Sub

    If Ficha.State = 1 Then Ficha.Close
    Set Ficha = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub ListaFicha_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trata_erro

    txtid_prazo.text = ListaFicha.SelectedItem.Tag

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdConsultar_Click()
    Aguarde_Process Me, True
    Call consultarNome_Cliente
    Aguarde_Process Me, False
End Sub

Private Sub consultarNome_Cliente()

    Sqlconsulta = " clientes.status = 'A'"

    If txtConsulta.text <> "" Then
        Sqlconsulta = " clientes.cliente like '%" & txtConsulta.text & "%'"
    End If

    Lista (Sqlconsulta)

End Sub


Private Sub txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub




Private Sub extrato()
'---------------------------------------  Vendas Extrato -----------------------------------
    On Error GoTo trata_erro
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

    Dim mDebito As Double
    Dim mCredito As Double
    Dim mAPagar As Double
    Dim mDaDos As Integer
    Dim mCabecarioDados As String
    Dim mArquivo As String
    Dim mDescricao As String



    Dim bRet As Boolean


    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    If txtid_prazo.text <> "" Then
        Sql = "SELECT SUM(valorpagto) AS totalcredito, prazo.id_prazo, prazo.id_prazo,"
        Sql = Sql & " clientes.id_cliente, clientes.cliente"
        Sql = Sql & " From"
        Sql = Sql & " prazopagto"
        Sql = Sql & " left join prazo ON  prazopagto.id_prazo = prazo.id_prazo"
        Sql = Sql & " left join clientes on prazo.id_cliente = clientes.id_cliente"
        Sql = Sql & " Where"
        Sql = Sql & " Prazo.id_prazo = '" & txtid_prazo.text & "'"

        If Tabela.State = 1 Then Tabela.Close
        Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Tabela.RecordCount > 0 Then
            If VarType(Tabela("totalcredito")) <> vbNull Then mCredito = Tabela("totalcredito") Else mCredito = 0
            If VarType(Tabela("cliente")) <> vbNull Then mCliente = Tabela("Cliente")
        Else
            mCredito = 0
        End If
    Else
        mCredito = 0
    End If



    Sql = " SELECT prazoitem.*, estoques.id_estoque, estoques.unidade, estoques.descricao, estoques.codigo_est,"
    Sql = Sql & " (prazoitem.quantidade * prazoitem.preco_venda) as total"
    Sql = Sql & " From"
    Sql = Sql & " prazoitem"
    Sql = Sql & " LEFT JOIN estoques ON prazoitem.id_estoque = estoques.id_estoque"
    Sql = Sql & " Where prazoitem.id_prazo = '" & txtid_prazo.text & "'"
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
        ' mDaDos = 38 - Len(Tabela("descricao") & "-" & Format(Tabela("datacompra"), "dd/mm/yy"))
        mDaDos = 30 - Len(Mid(Tabela("descricao"), 1, 30))
        If mDaDos < 0 Then mDaDos = Len(Tabela("descricao")) - 30


        If mDaDos = 0 Then
            mDescricao = Mid(Tabela("descricao"), 1, 30) & "-" & Format(Tabela("dataCompra"), "DD/MM/YY")
        Else
            mDescricao = Mid(Tabela("descricao") & Space(mDaDos), 1, 30) & "-" & Format(Tabela("dataCompra"), "DD/MM/YY")
        End If

        txtExtrato.text = txtExtrato.text & Mid(Tabela("codigo_est"), 1, 6) & Space(1)
        txtExtrato.text = txtExtrato.text & mDescricao & Space(3)
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
        txtExtrato.text = txtExtrato.text & "                                                                   continua....." & vbCrLf
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
    txtExtrato.text = txtExtrato.text & "............................................." & vbCrLf
    txtExtrato.text = txtExtrato.text & "Data....: " & Format$(Date, "dd/mm/yy") & Space(10) & "Extrato do Orcamento" & Space(14) & "Hora....: " & Format$(Time, "hh:nn:ss") & vbCrLf
    txtExtrato.text = txtExtrato.text & "Cliente.: " & mCliente & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & vbCrLf
    txtExtrato.text = txtExtrato.text & "Codigo Descricao                        Data    Quant.     Preco      Total  " & vbCrLf
    txtExtrato.text = txtExtrato.text & "------ ---------------------------------------- ---------- ---------- ----------" & vbCrLf

    Return

Rodape:

    txtExtrato.text = txtExtrato.text & "--------------------------------------------------------------------------------" & vbCrLf
    mAPagar = mDebito - mCredito


    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Debito..R$ " & Alinhar(Format(mDebito, "###,##0.00"), 10) & vbCrLf
    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Credito.R$ " & Alinhar(Format(mCredito, "###,##0.00"), 10) & vbCrLf
    txtExtrato.text = txtExtrato.text & Space(51) & "Total a Pagar...R$ " & Alinhar(Format(mAPagar, "###,##0.00"), 10) & vbCrLf


    Do While intLinhafinal < 19
        txtExtrato.text = txtExtrato.text & vbCrLf
        intLinhafinal = intLinhafinal + 1
    Loop



    Return


    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description) & txtExtrato.text

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


Private Function Alinhar(texto As String, Largura As Integer)
    Alinhar = String(Largura - Len(texto), " ") & texto
End Function

