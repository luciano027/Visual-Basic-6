VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendasClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendas Cliente"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Saldo"
      Height          =   855
      Left            =   4440
      TabIndex        =   30
      Top             =   4200
      Width           =   1935
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Debito - Credito"
      Height          =   855
      Left            =   2400
      TabIndex        =   28
      Top             =   4200
      Width           =   1935
      Begin VB.Label lblTotalCredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Limite"
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   2175
      Begin VB.Label lblLimite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtPrazo 
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   9135
      Begin RichTextLib.RichTextBox txtObservacao 
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmVendasClientes.frx":0000
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observação"
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
         TabIndex        =   23
         Top             =   0
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtuf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label1 
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
         TabIndex        =   21
         Top             =   0
         Width           =   9135
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.TextBox txtid_cliente 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_venda 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   320
         Width           =   8535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
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
         Width           =   9135
      End
      Begin VB.Image cmdConsultar 
         Height          =   360
         Left            =   8640
         Picture         =   "frmVendasClientes.frx":0082
         Stretch         =   -1  'True
         Top             =   360
         Width           =   360
      End
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   615
      Left            =   8160
      TabIndex        =   0
      Top             =   4440
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
      Picture         =   "frmVendasClientes.frx":038C
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
      TabIndex        =   1
      Top             =   5145
      Width           =   9480
      _ExtentX        =   16722
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
      Left            =   6960
      TabIndex        =   2
      Top             =   4440
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
      Picture         =   "frmVendasClientes.frx":0496
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "frmVendasClientes.frx":09E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9795
   End
End
Attribute VB_Name = "frmVendasClientes"
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
    Atualiza_cliente
End Sub

Private Sub Form_Load()
    Me.Width = 9570
    Me.Height = 5895
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmVendasClientes = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendasClientes = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdConsultar_Click()
    On Error GoTo trata_erro

    With frmConsultaClientes
        .Show 1
    End With

    If IDCliente <> "" Then
        txtid_cliente.text = IDCliente
        IDCliente = ""
    End If

    Atualiza_cliente

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub Atualiza_cliente()
    On Error GoTo trata_erro

    Dim Tabela As ADODB.Recordset

    Set Tabela = CreateObject("ADODB.Recordset")

    Sql = "SELECT clientes.* FROM Clientes WHERE id_Cliente = '" & txtid_cliente.text & "'"
    If Tabela.State = 1 Then Tabela.Close
    Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
    If Tabela.RecordCount > 0 Then
        If VarType(Tabela("id_Cliente")) <> vbNull Then txtid_cliente.text = Tabela("id_Cliente") Else txtid_cliente.text = ""
        If VarType(Tabela("Cliente")) <> vbNull Then txtCliente.text = Tabela("Cliente") Else txtCliente.text = ""
        If VarType(Tabela("Rua")) <> vbNull Then txtRua.text = Tabela("Rua") Else txtRua.text = ""
        If VarType(Tabela("bairro")) <> vbNull Then txtBairro.text = Tabela("bairro") Else txtBairro.text = ""
        If VarType(Tabela("cidade")) <> vbNull Then txtCidade.text = Tabela("cidade") Else txtCidade.text = ""
        If VarType(Tabela("cep")) <> vbNull Then txtCep.text = Tabela("cep") Else txtCep.text = ""
        If VarType(Tabela("uf")) <> vbNull Then txtuf.text = Tabela("uf") Else txtuf.text = ""
        If VarType(Tabela("numero")) <> vbNull Then txtNumero.text = Tabela("numero") Else txtNumero.text = ""
        If VarType(Tabela("observacao")) <> vbNull Then txtObservacao.text = Tabela("observacao") Else txtObservacao.text = ""
        If VarType(Tabela("prazo")) <> vbNull Then txtPrazo.text = Tabela("prazo")
        If VarType(Tabela("limite")) <> vbNull Then lblLimite.Caption = Format(Tabela("limite"), "###,##0.00") Else lblLimite.Caption = "0,00"
    End If

    Atualiza

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata_erro

    If txtid_venda.text = "" Then

        campo = "dataVenda"
        Scampo = "'" & Format(Now, "YYYYMMDD") & "'"

        If txtid_cliente.text <> "" Then
            campo = campo & ", id_cliente"
            Scampo = Scampo & ", '" & txtid_cliente.text & "'"
        End If

        campo = campo & ", maquina"
        Scampo = Scampo & ", '" & MicroBD & "'"

        campo = campo & ", status"
        Scampo = Scampo & ", 'A'"

        ' Incluir valor na tabela EntradaItens
        sqlIncluir "vendas", campo, Scampo, Me, "N"

        Buscar_id

    Else
        If txtid_cliente.text <> "" Then
            Sqlconsulta = "id_venda = '" & txtid_venda.text & "'"

            campo = "id_cliente = '" & txtid_cliente.text & "'"

            sqlAlterar "vendas", campo, Sqlconsulta, Me, "N"
            sqlAlterar "saida", campo, Sqlconsulta, Me, "N"
        End If
    End If

    If txtid_cliente.text <> "" Then
        ' Consulta os dados da tabela clientes
        Sqlconsulta = "id_cliente = '" & txtid_cliente.text & "'"

        campo = "data_cadastro = '" & Format(Now, "YYYYMMDD") & "'"

        If txtCep.text <> "" Then campo = campo & ", cep = '" & txtCep.text & "'"
        If txtuf.text <> "" Then campo = campo & ", uf = '" & txtuf.text & "'"
        If txtNumero.text <> "" Then campo = campo & ", numero = '" & txtNumero.text & "'"
        If txtBairro.text <> "" Then campo = campo & ", bairro = '" & txtBairro.text & "'"
        If txtCidade.text <> "" Then campo = campo & ", cidade = '" & txtCidade.text & "'"
        If txtRua.text <> "" Then campo = campo & ", Rua = '" & txtRua.text & "'"
        If txtObservacao.text <> "" Then campo = campo & ", observacao = '" & txtObservacao.text & "'"

        ' Aletar dos dados da tabela clientes
        sqlAlterar "clientes", campo, Sqlconsulta, Me, "N"

    End If

    With frmVendas
        .lblCliente.Caption = txtCliente.text
        .lblid_cliente.Caption = txtid_cliente.text
        .lblid_venda.Caption = txtid_venda.text
        .lbldataVenda.Caption = Format(Now, "DD/MM/YYYY")
        .lblSaldo.Caption = Format(lblSaldo.Caption, "###,##0.00")
        .txtPrazo.text = txtPrazo.text
    End With

    Unload Me

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

Private Sub Atualiza()
    On Error GoTo trata_erro
    Dim mCredito As Double
    Dim mDebito As Double
    Dim mPagar As Double
    Dim mSaldo As Double
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
    mSaldo = lblLimite.Caption - mPagar

    '  lbltotalDebito.Caption = Format(mDebito, "###,##0.00")

    lbltotalCredito.Caption = Format(mPagar, "###,##0.00")
    lblSaldo.Caption = Format(mSaldo, "###,##0.00")

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

