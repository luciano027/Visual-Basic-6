VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLojaCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loja"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   9120
      TabIndex        =   33
      Top             =   2160
      Width           =   3495
      Begin VB.TextBox txtcnpj 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtinscricao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Documento Identificação"
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
         TabIndex        =   37
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   8895
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtuf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   20
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
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
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
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
         TabIndex        =   26
         Top             =   0
         Width           =   8895
      End
      Begin VB.Image cmdConsultaCEP 
         Height          =   240
         Left            =   1320
         Picture         =   "frmLojaCadastro.frx":0000
         ToolTipText     =   "Verificar CEP Digitado"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image cmdConsulCEP 
         Height          =   360
         Left            =   1560
         Picture         =   "frmLojaCadastro.frx":0342
         Stretch         =   -1  'True
         ToolTipText     =   "Consulta CEP "
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   12495
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   960
         Left            =   10920
         TabIndex        =   15
         Top             =   450
         Width           =   1455
         Begin VB.OptionButton optInativo 
            Caption         =   "Inativo"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   650
            Width           =   855
         End
         Begin VB.OptionButton optAtivo 
            Caption         =   "Ativo"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label3 
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
            TabIndex        =   18
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.TextBox txtTelefone2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtTelefone1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   10695
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Top             =   1080
         Width           =   7575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Identificação"
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
         TabIndex        =   14
         Top             =   0
         Width           =   12855
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtChave 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_loja 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtid_login 
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   4080
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4875
      Width           =   12765
      _ExtentX        =   22516
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
      Left            =   11520
      TabIndex        =   39
      Top             =   4080
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
      Picture         =   "frmLojaCadastro.frx":064C
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
      Left            =   10320
      TabIndex        =   40
      Top             =   4080
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
      Picture         =   "frmLojaCadastro.frx":0756
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
      Left            =   9120
      TabIndex        =   41
      Top             =   4080
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
      Picture         =   "frmLojaCadastro.frx":0CA8
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
      Picture         =   "frmLojaCadastro.frx":11FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14955
   End
End
Attribute VB_Name = "frmLojaCadastro"
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
    Me.Width = 12855
    Me.Height = 5730
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
    Set frmLojaCadastro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLojaCadastro = Nothing
    MenuPrincipal.AbilidataMenu
End Sub


Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro

    If txtid_loja.text <> "" Then
        confirma = MsgBox("Confirma Exclusão", vbQuestion + vbYesNo)
        If confirma = vbYes Then

            Sqlconsulta = "id_loja = '" & txtid_loja.text & "'"

            sqlDeletar "Loja", Sqlconsulta, Me, "S"

            Unload Me

        End If
    End If

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
            MsgBox ("Nome da Loja não pode ficar em branco..")
            txtDescricao.SetFocus
            Exit Sub
        End If

        If txtTelefone1.text <> "" Then
            campo = campo & ", Telefone1"
            Scampo = Scampo & ", '" & txtTelefone1.text & "'"
        End If

        If txtTelefone2.text <> "" Then
            campo = campo & ", Telefone2"
            Scampo = Scampo & ", '" & txtTelefone2.text & "'"
        End If

        If txtemail.text <> "" Then
            campo = campo & ", email"
            Scampo = Scampo & ", '" & txtemail.text & "'"
        End If

        If txtInscricao.text <> "" Then
            campo = campo & ", inscricao"
            Scampo = Scampo & ", '" & txtInscricao.text & "'"
        End If

        If txtCnpj.text <> "" Then
            campo = campo & ", cnpj"
            Scampo = Scampo & ", '" & txtCnpj.text & "'"
        End If

        If txtCep.text <> "" Then
            campo = campo & ", cep"
            Scampo = Scampo & ", '" & txtCep.text & "'"
        End If
        If txtuf.text <> "" Then
            campo = campo & ", uf"
            Scampo = Scampo & ", '" & txtuf.text & "'"
        End If

        If txtNumero.text <> "" Then
            campo = campo & ", numero"
            Scampo = Scampo & ", '" & txtNumero.text & "'"
        End If

        If txtBairro.text <> "" Then
            campo = campo & ", bairro"
            Scampo = Scampo & ", '" & txtBairro.text & "'"
        End If
        If txtCidade.text <> "" Then
            campo = campo & ", cidade"
            Scampo = Scampo & ", '" & txtCidade.text & "'"
        End If

        If txtRua.text <> "" Then
            campo = campo & ", rua"
            Scampo = Scampo & ", '" & txtRua.text & "'"
        End If

        If optAtivo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'A'"
        End If

        If optInativo.Value = True Then
            campo = campo & ", status"
            Scampo = Scampo & ", 'I'"
        End If

        ' Incluir valor na tabela Loja
        sqlIncluir "Loja", campo, Scampo, Me, "S"

    End If
    ' rotina de gravacao de alteracao dos dados
    If txtTipo.text = "A" Then

        ' Consulta os dados da tabela Loja
        Sqlconsulta = "id_Loja = '" & txtid_loja.text & "'"

        If txtDescricao.text <> "" Then campo = " descricao = '" & txtDescricao.text & "'" Else txtDescricao.SetFocus
        If txtTelefone1.text <> "" Then campo = campo & ", Telefone1 = '" & txtTelefone1.text & "'"
        If txtTelefone2.text <> "" Then campo = campo & ", Telefone2 = '" & txtTelefone2.text & "'"
        If txtemail.text <> "" Then campo = campo & ", email = '" & txtemail.text & "'"
        If txtCep.text <> "" Then campo = campo & ", cep = '" & txtCep.text & "'"
        If txtuf.text <> "" Then campo = campo & ", uf = '" & txtuf.text & "'"
        If txtNumero.text <> "" Then campo = campo & ", numero = '" & txtNumero.text & "'"
        If txtBairro.text <> "" Then campo = campo & ", bairro = '" & txtBairro.text & "'"
        If txtCidade.text <> "" Then campo = campo & ", cidade = '" & txtCidade.text & "'"
        If txtRua.text <> "" Then campo = campo & ", Rua = '" & txtRua.text & "'"
        If txtCnpj.text <> "" Then campo = campo & ", cnpj = '" & txtCnpj.text & "'"
        If txtInscricao.text <> "" Then campo = campo & ", inscricao = '" & txtInscricao.text & "'"
        If optAtivo.Value = True Then campo = campo & ", status = 'A'" Else campo = campo & ", status = 'I'"

        ' Aletar dos dados da tabela Loja
        sqlAlterar "Loja", campo, Sqlconsulta, Me, "S"

    End If

    cmdGravar.Enabled = False

    Unload Me

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub


Private Sub AutalizaCadastro()
    On Error GoTo trata_erro

    If txtTipo.text = "A" Or txtTipo.text = "E" Then
        If txtChave.text = "0" Then
            Dim loja As ADODB.Recordset
            ' conecta ao banco de dados

            Set loja = CreateObject("ADODB.Recordset")    '''

            ' abre um Recrodset da Tabela Loja
            Sql = " select "
            Sql = Sql & " Loja.*"
            Sql = Sql & " from  "
            Sql = Sql & " Loja "
            Sql = Sql & " where "
            Sql = Sql & " id_Loja = '" & txtid_loja.text & "'"

            If loja.State = 1 Then loja.Close
            loja.Open Sql, banco, adOpenKeyset, adLockOptimistic
            If loja.RecordCount > 0 Then

                If VarType(loja("descricao")) <> vbNull Then txtDescricao.text = loja("descricao") Else txtDescricao.text = ""
                If VarType(loja("Telefone1")) <> vbNull Then txtTelefone1.text = loja("Telefone1") Else txtTelefone1.text = ""
                If VarType(loja("Telefone2")) <> vbNull Then txtTelefone2.text = loja("Telefone2") Else txtTelefone2.text = ""
                If VarType(loja("email")) <> vbNull Then txtemail.text = loja("email") Else txtemail.text = ""
                If VarType(loja("Rua")) <> vbNull Then txtRua.text = loja("Rua") Else txtRua.text = ""
                If VarType(loja("bairro")) <> vbNull Then txtBairro.text = loja("bairro") Else txtBairro.text = ""
                If VarType(loja("cidade")) <> vbNull Then txtCidade.text = loja("cidade") Else txtCidade.text = ""
                If VarType(loja("cep")) <> vbNull Then txtCep.text = loja("cep") Else txtCep.text = ""
                If VarType(loja("uf")) <> vbNull Then txtuf.text = loja("uf") Else txtuf.text = ""
                If VarType(loja("numero")) <> vbNull Then txtNumero.text = loja("numero") Else txtNumero.text = ""
                If VarType(loja("cnpj")) <> vbNull Then txtCnpj.text = loja("cnpj") Else txtCnpj.text = ""
                If VarType(loja("inscricao")) <> vbNull Then txtInscricao.text = loja("inscricao") Else txtInscricao.text = ""
                If VarType(loja("status")) <> vbNull Then
                    If loja("status") = "A" Then optAtivo.Value = True
                    If loja("status") = "I" Then optInativo.Value = True
                End If

            End If
            If loja.State = 1 Then loja.Close
            Set loja = Nothing

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
    If KeyAscii = vbKeyReturn Then txtTelefone1.SetFocus
End Sub

'--- Telefone1
Private Sub txtTelefone1_GotFocus()
    txtTelefone1.BackColor = &H80FFFF
End Sub
Private Sub txtTelefone1_LostFocus()
    txtTelefone1.BackColor = &H80000014
    txtTelefone1.text = SoNumero(txtTelefone1.text)
    txtTelefone1.text = FormataTelefone(txtTelefone1.text)
    If Len(txtTelefone1.text) > 16 Then
        MsgBox "Comprimento do campo e de 16 digitos, voce digitou " & Len(txtTelefone1.text)
        txtTelefone1.SetFocus
    End If
End Sub
Private Sub txtTelefone1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtTelefone2.SetFocus
End Sub

'---Telefone2
Private Sub txtTelefone2_GotFocus()
    txtTelefone2.BackColor = &H80FFFF
End Sub
Private Sub txtTelefone2_LostFocus()
    txtTelefone2.BackColor = &H80000014
    txtTelefone2.text = SoNumero(txtTelefone2.text)
    txtTelefone2.text = FormataTelefone(txtTelefone2.text)
    If Len(txtTelefone2.text) > 16 Then
        MsgBox "Comprimento do campo e de 16 digitos, voce digitou " & Len(txtTelefone2.text)
        txtTelefone2.SetFocus
    End If
End Sub
Private Sub txtTelefone2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtemail.SetFocus
End Sub


'-------- email
Private Sub txtemail_GotFocus()
    txtemail.BackColor = &H80FFFF
End Sub
Private Sub txtemail_LostFocus()
    txtemail.BackColor = &H80000014
    If Len(txtemail.text) > 150 Then
        MsgBox "Comprimento do campo e de 150 digitos, voce digitou " & Len(txtemail.text)
        txtemail.SetFocus
    End If
End Sub
Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCep.SetFocus
End Sub


'-------- cep
Private Sub txtcep_GotFocus()
    txtCep.BackColor = &H80FFFF
    If CEPVer <> "" Then
        txtCep.text = CEPVer
        CEPVer = ""
    End If
End Sub
Private Sub txtcep_LostFocus()
    txtCep.BackColor = &H80000014
    txtCep.text = SoNumero(txtCep.text)
    If Len(txtCep.text) > 10 Then
        MsgBox "Comprimento do campo e de 10 digitos, voce digitou " & Len(txtCep.text)
        txtCep.SetFocus
    End If
End Sub
Private Sub txtcep_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtuf.SetFocus
End Sub
Private Sub cmdConsultaCEP_Click()
    On Error GoTo trata_erro
    Dim CEP As ADODB.Recordset
    Dim itemx As ListItem
    Dim Consult As String
    Dim CEPs As String
    ' conecta ao banco de dados
    Set CEP = CreateObject("ADODB.Recordset")

    If txtCep.text <> "" Then
        CEPs = SoNumero(txtCep.text)
        CEPs = Mid(txtCep.text, 1, 5) & "-" & Mid(txtCep.text, 6, 3)
        Aguarde_Process Me, True
        Sql = "select cep_endereco.id_endereco, cep_endereco.id_bairro, cep_endereco.endereco_completo, cep_endereco.cep, " _
            & " cep_cidade.id_cidade, cep_cidade.id_estado, cep_cidade.cidade, " _
            & " cep_bairro.id_bairro, cep_bairro.id_cidade, cep_bairro.bairro, " _
            & " cep_estados.id_estado , cep_estados.uf " _
            & " From " _
            & " cep_endereco " _
            & " left join cep_cidade on cep_endereco.id_cidade = cep_cidade.id_cidade " _
            & " left join cep_bairro on cep_endereco.id_bairro = cep_bairro.id_bairro " _
            & " left join cep_estados on cep_cidade.id_estado = cep_estados.id_estado " _
            & " where cep = '" & CEPs & "'"

        ' abre um Recrodset da Tabela CEP
        If CEP.State = 1 Then CEP.Close
        CEP.Open Sql, banco, adOpenKeyset, adLockOptimistic
        Aguarde_Process Me, False
        If CEP.RecordCount > 0 Then
            If VarType(CEP("uf")) <> vbNull Then txtuf.text = CEP("uf") Else txtuf.text = ""
            If VarType(CEP("bairro")) <> vbNull Then txtBairro.text = CEP("bairro") Else txtBairro.text = ""
            If VarType(CEP("cidade")) <> vbNull Then txtCidade.text = CEP("cidade") Else txtCidade.text = ""
            If VarType(CEP("endereco_completo")) <> vbNull Then txtRua.text = CEP("endereco_completo") Else txtRua.text = ""
            txtCep.text = CEPs
        Else
            MsgBox ("CEP Não encontrado..."), vbExclamation
        End If
    End If

    If CEP.State = 1 Then CEP.Close
    Set CEP = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub
Private Sub cmdConsulCEP_Click()
    With frmCepConsulta
        .Show 1
    End With
End Sub


'-------- uf
Private Sub txtUF_GotFocus()
    txtuf.BackColor = &H80FFFF
End Sub
Private Sub txtuf_LostFocus()
    txtuf.BackColor = &H80000014
    If Len(txtuf.text) > 2 Then
        MsgBox "Comprimento do campo e de 2 digitos, voce digitou " & Len(txtuf.text)
        txtuf.SetFocus
    End If
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNumero.SetFocus
End Sub

'-------- Numero
Private Sub txtNumero_GotFocus()
    txtNumero.BackColor = &H80FFFF
End Sub
Private Sub txtNumero_LostFocus()
    txtNumero.BackColor = &H80000014
    If Len(txtNumero.text) > 5 Then
        MsgBox "Comprimento do campo e de 5 digitos, voce digitou " & Len(txtNumero.text)
        txtNumero.SetFocus
    End If
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtBairro.SetFocus
End Sub
'-------- bairro
Private Sub txtbairro_GotFocus()
    txtBairro.BackColor = &H80FFFF
End Sub
Private Sub txtbairro_LostFocus()
    txtBairro.BackColor = &H80000014
    If Len(txtBairro.text) > 70 Then
        MsgBox "Comprimento do campo e de 70 digitos, voce digitou " & Len(txtBairro.text)
        txtBairro.SetFocus
    End If
End Sub
Private Sub txtbairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCidade.SetFocus
End Sub
'-------- cidade
Private Sub txtcidade_GotFocus()
    txtCidade.BackColor = &H80FFFF
End Sub
Private Sub txtcidade_LostFocus()
    txtCidade.BackColor = &H80000014
    If Len(txtCidade.text) > 50 Then
        MsgBox "Comprimento do campo e de 50 digitos, voce digitou " & Len(txtCidade.text)
        txtCidade.SetFocus
    End If
End Sub
Private Sub txtcidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtRua.SetFocus
End Sub


'-------- Rua
Private Sub txtRua_GotFocus()
    txtRua.BackColor = &H80FFFF
End Sub
Private Sub txtRua_LostFocus()
    txtRua.BackColor = &H80000014
    If Len(txtRua.text) > 100 Then
        MsgBox "Comprimento do campo e de 100 digitos, voce digitou " & Len(txtRua.text)
        txtRua.SetFocus
    End If
End Sub
Private Sub txtRua_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCnpj.SetFocus
End Sub

'-------- cnpj
Private Sub txtcnpj_GotFocus()
    txtCnpj.BackColor = &H80FFFF
End Sub
Private Sub txtcnpj_LostFocus()
    txtCnpj.BackColor = &H80000014
    If Len(txtCnpj.text) > 30 Then
        MsgBox "Comprimento do campo e de 30 digitos, voce digitou " & Len(txtCnpj.text)
        txtCnpj.SetFocus
    End If
End Sub

Private Sub txtcnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtInscricao.SetFocus
End Sub


'-------- inscricao
Private Sub txtInscricao_GotFocus()
    txtInscricao.BackColor = &H80FFFF
End Sub
Private Sub txtInscricao_LostFocus()
    txtInscricao.BackColor = &H80000014
    If Len(txtInscricao.text) > 20 Then
        MsgBox "Comprimento do campo e de 20 digitos, voce digitou " & Len(txtInscricao.text)
        txtInscricao.SetFocus
    End If
End Sub

Private Sub txtInscricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdGravar.SetFocus
End Sub

