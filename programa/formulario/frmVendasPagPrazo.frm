VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendasPagPrazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venda a Prazo"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtApagar 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Text            =   "N"
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_venda 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_prazo 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtPrazo 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Text            =   "N"
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   240
         TabIndex        =   1
         Top             =   435
         Width           =   5655
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente - Situação Financeira (R$)"
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
         TabIndex        =   10
         Top             =   0
         Width           =   6255
      End
      Begin VB.Label lblPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   4440
         TabIndex        =   9
         Top             =   1575
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A Pagar (R$)"
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
         Left            =   4440
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lbltotalDebito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label lbltotalCredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2280
         TabIndex        =   6
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Debito (R$)"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Credito (R$)"
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
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   1440
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   6555
      _ExtentX        =   11562
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
   Begin Vendas.VistaButton cmdConfirmaPrazo 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      Caption         =   "Confirma venda a Prazo"
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
      Picture         =   "frmVendasPagPrazo.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   4920
      TabIndex        =   13
      Top             =   2640
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
      Picture         =   "frmVendasPagPrazo.frx":001C
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
      Picture         =   "frmVendasPagPrazo.frx":0126
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "frmVendasPagPrazo"
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
    Atualiza
    If txtPrazo.text = "S" Then cmdConfirmaPrazo.Enabled = True Else cmdConfirmaPrazo.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Width = 6645
    Me.Height = 4230
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmVendasPagPrazo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendasPagPrazo = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdConfirmaPrazo_Click()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "select prazo.id_cliente, prazo.id_prazo "
    Sql = Sql & " FROM prazo"
    Sql = Sql & " where"
    Sql = Sql & " prazo.id_cliente = '" & txtid_cliente.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_prazo")) <> vbNull Then txtid_prazo.text = Tabela("id_prazo")
    Else
        campo = "id_cliente"
        Scampo = "'" & txtid_cliente.text & "'"

        campo = campo & ", data_Venda"
        Scampo = Scampo & ", '" & Format(Now, "YYYYMMDD") & "'"

        sqlIncluir "Prazo", campo, Scampo, Me, "N"

        Buscar_id
    End If

    Incluir_item_prazo

    Sqlconsulta = "id_venda = '" & txtid_venda.text & "'"

    campo = "prazo = '" & FormatValor(txtApagar.text, 1) & "'"
    campo = campo & ", status = 'P'"
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
    Scampo = Scampo & ", '" & Mid("Venda a Prazo:" & txtCliente.text, 1, 100) & "'"

    campo = campo & ", datacaixa"
    Scampo = Scampo & ", '" & Format(Now, "YYYYMMDD") & "'"

    campo = campo & ", valorcaixaprazo"
    Scampo = Scampo & ", '" & FormatValor(txtApagar.text, 1) & "'"

    campo = campo & ", status"
    Scampo = Scampo & ", 'P'"

    sqlIncluir "caixa", campo, Scampo, Me, "N"

    frmVendas.txtStatus.text = "P"
    frmVendasPagamento.txtConfirmaPag.text = "S"

    MsgBox ("Pagamento efetuado com sucesso.."), vbInformation

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Atualiza()
    On Error GoTo trata_erro
    Dim mCredito As Double
    Dim mDebito As Double
    Dim mPagar As Double
    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT clientes.prazo "
    Sql = Sql & " From"
    Sql = Sql & " clientes"
    Sql = Sql & " where"
    Sql = Sql & " clientes.id_cliente = '" & txtid_cliente.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("prazo")) <> vbNull Then txtPrazo.text = Tabela("prazo")
    End If


    Sql = "SELECT prazo.id_cliente,"
    Sql = Sql & " SUM(prazopagto.ValorPagto) As total"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN prazopagto ON prazo.id_prazo = prazopagto.id_prazo"
    Sql = Sql & " where"
    Sql = Sql & " prazo.id_cliente = '" & txtid_cliente.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("total")) <> vbNull Then mCredito = Tabela("total")
    Else
        mCredito = 0
    End If


    Sql = "SELECT prazo.id_cliente,"
    Sql = Sql & " SUM(prazoitem.quantidade * prazoitem.preco_venda) As total"
    Sql = Sql & " From"
    Sql = Sql & " Prazo"
    Sql = Sql & " LEFT JOIN prazoitem ON prazo.id_prazo = prazoitem.id_prazo"
    Sql = Sql & " where"
    Sql = Sql & " prazo.id_cliente = '" & txtid_cliente.text & "'"

    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("total")) <> vbNull Then mDebito = Tabela("total")
    Else
        mDebito = 0
    End If

    mPagar = mDebito - mCredito

    lbltotalDebito.Caption = Format(mDebito, "###,##0.00")
    lbltotalCredito.Caption = Format(mCredito, "###,##0.00")
    lblPagar.Caption = Format(mPagar, "###,##0.00")

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

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

Private Sub Incluir_item_prazo()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "select saida.*"
    Sql = Sql & " from "
    Sql = Sql & " saida"
    Sql = Sql & " where "
    Sql = Sql & " saida.id_venda = '" & txtid_venda.text & "'"

    ' abre um Recrodset da Tabela Tabela
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.BOF = True And Tabela.EOF = True Then Exit Sub
    While Not Tabela.EOF

        campo = "id_prazo"
        Scampo = "'" & txtid_prazo.text & "'"

        campo = campo & ", id_venda"
        Scampo = Scampo & ", '" & txtid_venda.text & "'"

        campo = campo & ", id_estoque"
        Scampo = Scampo & ", '" & Tabela("id_estoque") & "'"

        campo = campo & ", quantidade"
        Scampo = Scampo & ", '" & FormatValor(Tabela("quantidade"), 1) & "'"

        campo = campo & ", preco_venda"
        Scampo = Scampo & ", '" & FormatValor(Tabela("preco_venda"), 1) & "'"

        campo = campo & ", datacompra"
        Scampo = Scampo & ", '" & Format(Tabela("datasaida"), "YYYYMMDD") & "'"

        sqlIncluir "prazoitem", campo, Scampo, Me, "N"

        Tabela.MoveNext
    Wend

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


