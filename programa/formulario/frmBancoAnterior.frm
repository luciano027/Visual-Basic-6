VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBancoAnterior 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banco Anterior"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrazo 
      Caption         =   "Prazo"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton cdmFornecedores 
      Caption         =   "Fornecedores"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   5655
   End
   Begin VB.CommandButton cmdAgenda 
      Caption         =   "Agenda"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton cmdClientes 
      Caption         =   "Clientes"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton cmdEstoque 
      Caption         =   "Estoque"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3750
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21944
            MinWidth        =   21944
         EndProperty
      EndProperty
   End
   Begin Vendas.VistaButton cmdSair 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Picture         =   "frmBancoAnterior.frx":0000
      Pictures        =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   65280
      Enabled         =   -1  'True
      NoBackground    =   0   'False
      BackColor       =   16777215
      PictureOffset   =   0
   End
   Begin VB.Image Image1 
      Height          =   4665
      Left            =   0
      Picture         =   "frmBancoAnterior.frx":010A
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   6075
   End
End
Attribute VB_Name = "frmBancoAnterior"
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
Dim chave As String
Dim Sql As String
Dim mID_prazo As String
Dim mID_cliente As String



Private Sub cdmFornecedores_Click()
    On Error GoTo trata_erro
    Dim Fornecedores1 As ADODB.Recordset
    ' conecta ao banco de dados
    Sql = "SELECT * FROM Fornecedores1 order by cod_for"
    Set Fornecedores1 = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela Fornecedores1
    If Fornecedores1.State = 1 Then Fornecedores1.Close
    Fornecedores1.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not Fornecedores1.EOF

        campo = " data_cadastro"
        Scampo = "'" & Format(Date$, "YYYYMMDD") & "'"

        campo = campo & ", status"
        Scampo = Scampo & ", 'A'"

        If VarType(Fornecedores1("cod_for")) <> vbNull Then
            campo = campo & ", cod_for"
            Scampo = Scampo & ", '" & Fornecedores1("cod_for") & "'"
        End If

        If VarType(Fornecedores1("forneced")) <> vbNull Then
            campo = campo & ", fornecedor"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("forneced"), 1, 50) & "'"
        End If

        If VarType(Fornecedores1("fantasia")) <> vbNull Then
            campo = campo & ", fantasia"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("fantasia"), 1, 50) & "'"
        End If

        If VarType(Fornecedores1("cgccpf")) <> vbNull Then
            campo = campo & ", cnpj"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("cgccpf"), 1, 30) & "'"
        End If

        If VarType(Fornecedores1("cepf")) <> vbNull Then
            campo = campo & ", cep"
            Scampo = Scampo & ", '" & Mid(SoNumero(Fornecedores1("cepf")), 1, 10) & "'"
        End If

        If VarType(Fornecedores1("rgie")) <> vbNull Then
            campo = campo & ", inscricao"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("rgie"), 1, 30) & "'"
        End If

        If VarType(Fornecedores1("enderf")) <> vbNull Then
            campo = campo & ", rua"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("enderf"), 1, 100) & "'"
        End If

        If VarType(Fornecedores1("bairrof")) <> vbNull Then
            campo = campo & ", bairro"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("bairrof"), 1, 70) & "'"
        End If

        If VarType(Fornecedores1("cidadef")) <> vbNull Then
            campo = campo & ", cidade"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("cidadef"), 1, 50) & "'"
        End If

        If VarType(Fornecedores1("uf")) <> vbNull Then
            campo = campo & ", uf"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("uf"), 1, 2) & "'"
        End If


        If VarType(Fornecedores1("fonef")) <> vbNull Then
            campo = campo & ", tel2"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("fonef"), 1, 16) & "'"
        End If

        If VarType(Fornecedores1("faxf")) <> vbNull Then
            campo = campo & ", fax"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("faxf"), 1, 16) & "'"
        End If

        If VarType(Fornecedores1("contato")) <> vbNull Then
            campo = campo & ", contato"
            Scampo = Scampo & ", '" & Mid(Fornecedores1("contato"), 1, 30) & "'"
        End If


        sqlIncluir "Fornecedores", campo, Scampo, Me, "N"

        Fornecedores1.MoveNext
    Wend
    If Fornecedores1.State = 1 Then Fornecedores1.Close
    Set Fornecedores1 = Nothing

    MsgBox ("Arquivos incluido com sucesso.."), vbInformation
    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdAgenda_Click()
    On Error GoTo trata_erro
    Dim Agenda1 As ADODB.Recordset
    ' conecta ao banco de dados
    Sql = "SELECT * FROM Agenda1"
    Set Agenda1 = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela Agenda1
    If Agenda1.State = 1 Then Agenda1.Close
    Agenda1.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not Agenda1.EOF

        campo = " data_cadastro"
        Scampo = "'" & Format(Date$, "YYYYMMDD") & "'"

        If VarType(Agenda1("nome_age")) <> vbNull Then
            campo = campo & ", nome"
            Scampo = Scampo & ", '" & Agenda1("nome_age") & "'"
        End If

        If VarType(Agenda1("obs1")) <> vbNull Then
            campo = campo & ", atividade"
            Scampo = Scampo & ", '" & Mid(Agenda1("obs1"), 1, 40) & "'"
        End If

        If VarType(Agenda1("telefone")) <> vbNull Then
            campo = campo & ", telefone"
            Scampo = Scampo & ", '" & Mid(Agenda1("telefone"), 1, 16) & "'"
        End If

        If VarType(Agenda1("celular")) <> vbNull Then
            campo = campo & ", celular"
            Scampo = Scampo & ", '" & Mid(Agenda1("celular"), 1, 16) & "'"
        End If

        If VarType(Agenda1("fax")) <> vbNull Then
            campo = campo & ", telefone2"
            Scampo = Scampo & ", '" & Mid(Agenda1("fax"), 1, 16) & "'"
        End If

        If VarType(Agenda1("obs1")) <> vbNull Then
            campo = campo & ", obs"
            Scampo = Scampo & ", '" & Agenda1("obs1") & "'"
            If VarType(Agenda1("obs1")) <> vbNull Then Scampo = Scampo & "'  " & Agenda1("obs2") & "'"
        End If

        sqlIncluir "Telefone", campo, Scampo, Me, "N"

        Agenda1.MoveNext
    Wend
    If Agenda1.State = 1 Then Agenda1.Close
    Set Agenda1 = Nothing

    MsgBox ("Arquivos incluido com sucesso.."), vbInformation
    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)


End Sub

Private Sub cmdClientes_Click()
    On Error GoTo trata_erro
    Dim clientes1 As ADODB.Recordset
    ' conecta ao banco de dados
    Sql = "SELECT * FROM clientes order by codigo_cli"
    Set clientes1 = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela clientes1
    If clientes1.State = 1 Then clientes1.Close
    clientes1.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not clientes1.EOF

        campo = " prazo = 'S'"

        sqlIncluir "clientes", campo, Scampo, Me, "N"

        clientes1.MoveNext
    Wend
    If clientes1.State = 1 Then clientes1.Close
    Set clientes1 = Nothing

    MsgBox ("Arquivos incluido com sucesso.."), vbInformation
    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub

Private Sub cmdPrazo_Click()
    On Error GoTo trata_erro
    Dim prazo1 As ADODB.Recordset
    Dim Prazo As ADODB.Recordset
    Dim sql1 As String
    Dim sqlconsulta1 As String

    ' conecta ao banco de dados
    Sql = "SELECT *, clientes.id_cliente, clientes1.data_deb, clientes1.debito, clientes1.credito, "
    Sql = Sql & " estoques.id_estoque "
    Sql = Sql & " FROM prazo1 "
    Sql = Sql & " left join clientes on prazo1.codigo_cli = clientes.codigo_cli"
    Sql = Sql & " left join clientes1 on prazo1.codigo_cli = clientes1.codigo_cli"
    Sql = Sql & " left join estoques on prazo1.codigo_est = estoques.codigo_est"
    Sql = Sql & " where"
    Sql = Sql & " prazo1.codigo_cli is not null"
    Sql = Sql & " order by prazo1.codigo_cli"

    Set prazo1 = CreateObject("ADODB.Recordset")
    Set Prazo = CreateObject("ADODB.Recordset")
    ' abre um Recrodset da Tabela prazo1
    If prazo1.State = 1 Then prazo1.Close
    prazo1.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not prazo1.EOF
        sql1 = "Select * from prazo where id_cliente = '" & prazo1("id_cliente") & "'"
        If Prazo.State = 1 Then Prazo.Close
        Prazo.Open sql1, banco, adOpenKeyset, adLockOptimistic

        If Prazo.RecordCount > 0 Then
            mID_prazo = Prazo("id_prazo")
            mID_cliente = prazo1("id_cliente")

            '----------------------
            campo = "id_prazo"
            Scampo = "'" & mID_prazo & "'"

            campo = campo & ", id_estoque"
            Scampo = Scampo & ", '" & prazo1("id_estoque") & "'"

            campo = campo & ", quantidade"
            Scampo = Scampo & ", '" & FormatValor(prazo1("quantidade"), 1) & "'"

            campo = campo & ", preco_venda"
            Scampo = Scampo & ", '" & FormatValor(prazo1("valor"), 1) & "'"

            campo = campo & ", dataCompra"
            Scampo = Scampo & ", '" & Format(prazo1("Data"), "YYYYMMDD") & "'"

            sqlIncluir "prazoitem", campo, Scampo, Me, "N"

        Else
            If VarType(prazo1("id_cliente")) <> vbNull Then
                campo = "id_cliente"
                Scampo = "'" & prazo1("id_cliente") & "'"

                If VarType(prazo1("data_deb")) <> vbNull Then
                    campo = campo & ", data_venda"
                    Scampo = Scampo & ", '" & Format(prazo1("data_deb"), "YYYYMMDD") & "'"
                End If

                sqlIncluir "Prazo", campo, Scampo, Me, "N"

                Buscar_id

                '------------------------ Incluir Credito
                If VarType(prazo1("credito")) <> vbNull And prazo1("credito") > 0 Then
                    campo = "id_prazo"
                    Scampo = "'" & mID_prazo & "'"

                    campo = campo & ", valorpagto"
                    Scampo = "'" & FormatValor(prazo1("credito"), 1) & "'"

                    campo = campo & ", datapagto"
                    Scampo = Scampo & ", '" & Format(prazo1("data_deb"), "YYYYMMDD") & "'"

                    sqlIncluir "prazopagto", campo, Scampo, Me, "N"

                End If

                '----------------------
                campo = "id_prazo"
                Scampo = "'" & mID_prazo & "'"

                campo = campo & ", id_estoque"
                Scampo = Scampo & ", '" & prazo1("id_estoque") & "'"

                campo = campo & ", quantidade"
                Scampo = Scampo & ", '" & FormatValor(prazo1("quantidade"), 1) & "'"

                campo = campo & ", preco_venda"
                Scampo = Scampo & ", '" & FormatValor(prazo1("valor"), 1) & "'"

                campo = campo & ", dataCompra"
                Scampo = Scampo & ", '" & Format(prazo1("Data"), "YYYYMMDD") & "'"

                sqlIncluir "prazoitem", campo, Scampo, Me, "N"

            End If

        End If


        prazo1.MoveNext
    Wend
    If prazo1.State = 1 Then prazo1.Close
    Set prazo1 = Nothing

    MsgBox ("Arquivos incluido com sucesso.."), vbInformation
    Unload Me

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
        If VarType(Tabela("maxid")) <> vbNull Then mID_prazo = Tabela("maxid") Else mID_prazo = ""
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub



Private Sub Form_Load()

    Set Me.Icon = LoadPicture(ICONBD)
End Sub

Private Sub Form_Activate()
    On Error GoTo trata_erro

    Me.Width = 6030
    Me.Height = 4500
    Centerform Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub cmdSair_Click()
    Unload Me
    Set frmBancoAnterior = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBancoAnterior = Nothing
End Sub

Private Sub cmdEstoque_Click()
    On Error GoTo trata_erro
    Dim mSaldo As Double
    Dim mIdestoque As String
    Dim estoque1 As ADODB.Recordset
    Dim Saldo1 As ADODB.Recordset
    Dim Saldo2 As ADODB.Recordset
    ' conecta ao banco de dados

    Set estoque1 = CreateObject("ADODB.Recordset")
    Set Saldo1 = CreateObject("ADODB.Recordset")
    Set Saldo2 = CreateObject("ADODB.Recordset")


    ' abre um Recrodset da Tabela estoque1
    Sql = "SELECT * FROM estoque1 order by codigo_est"
    If estoque1.State = 1 Then estoque1.Close
    estoque1.Open Sql, banco, adOpenKeyset, adLockOptimistic
    While Not estoque1.EOF
        Sql = "SELECT estoques.id_estoque, estoques.codigo_est"
        Sql = Sql & " From"
        Sql = Sql & " Estoques"
        Sql = Sql & " where "
        Sql = Sql & " estoques.codigo_est = '" & estoque1("codigo_est") & "'"

        If Saldo1.State = 1 Then Saldo1.Close
        Saldo1.Open Sql, banco, adOpenKeyset, adLockOptimistic
        If Saldo1.RecordCount > 0 Then


            mIdestoque = Saldo1("id_estoque")

            Sql = "select estoquesaldo.* from estoquesaldo where id_estoque = '" & Saldo1("id_estoque") & "'"

            If Saldo2.State = 1 Then Saldo2.Close
            Saldo2.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If Saldo2.RecordCount > 0 Then

                mSaldo = estoque1("saldo") + Saldo2("saldo")
                campo = "saldo = '" & FormatValor(mSaldo, 1) & "'"
                Sqlconsulta = "estoquesaldo.id_estoque = '" & Saldo1("id_estoque") & "'"
                sqlAlterar "estoquesaldo", campo, Sqlconsulta, Me, "N"
            Else
                campo = "id_estoque"
                Scampo = "'" & Saldo1("id_estoque") & "'"

                campo = campo & ", saldo"
                Scampo = Scampo & ", '" & FormatValor(estoque1("saldo"), 1) & "'"

                sqlIncluir "estoquesaldo", campo, Scampo, Me, "N"
            End If
        End If
        estoque1.MoveNext
    Wend
    If estoque1.State = 1 Then estoque1.Close
    Set estoque1 = Nothing

    campo = "Prazo = 'S'"

    sqlAlterar "Clientes", campo, "1=1", Me, "N"

    MsgBox ("Arquivos incluido com sucesso.."), vbInformation
    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)

End Sub
