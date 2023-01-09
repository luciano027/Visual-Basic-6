VERSION 5.00
Begin VB.Form frmVendedorCodigo 
   BorderStyle     =   0  'None
   Caption         =   "Vendedor Codigo"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frCodigoVendedor 
      Caption         =   "Frame4"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtid_cliente 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtid_vendedor 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtid_prazo 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTipo 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtACesso 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin Vendas.VistaButton cmdSair 
         Height          =   615
         Left            =   3720
         TabIndex        =   2
         Top             =   480
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
         Picture         =   "frmVendedorCodigo.frx":0000
         Pictures        =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   65280
         Enabled         =   -1  'True
         NoBackground    =   0   'False
         BackColor       =   16777215
         PictureOffset   =   0
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "Senha de Acesso"
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
         TabIndex        =   3
         Top             =   0
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmVendedorCodigo"
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

Private Sub Form_Load()
    Me.Width = 4905
    Me.Height = 1335
    ' Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)

    MenuPrincipal.AbilidataMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendedorCodigo = Nothing
    MenuPrincipal.AbilidataMenu
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub VendasPagamento()
    On Error GoTo trata_erro
    Dim Tabela As ADODB.Recordset
    Dim mAcesso As String
    Dim mIdVendedor As String
    Dim mVendedor As String
    Dim mTipoAcesso As String
    ' conecta ao banco de dados
    Set Tabela = CreateObject("ADODB.Recordset")

    If txtAcesso.text <> "" Then
        Sql = "Select Vendedores.id_vendedor, vendedores.acesso, vendedores.vendedor, vendedores.tipo_acesso "
        Sql = Sql & " from "
        Sql = Sql & " vendedores"
        Sql = Sql & " where vendedores.acesso = '" & txtAcesso.text & "'"

        ' abre um Recrodset da Tabela Tabela
        If Tabela.State = 1 Then Tabela.Close
        Tabela.Open Sql, banco, adOpenKeyset, adLockOptimistic
        Aguarde_Process Me, False
        If Tabela.RecordCount > 0 Then
            If VarType(Tabela("vendedor")) <> vbNull Then mVendedor = Tabela("vendedor")
            If VarType(Tabela("id_vendedor")) <> vbNull Then mIdVendedor = Tabela("id_vendedor")
            If VarType(Tabela("acesso")) <> vbNull Then mAcesso = Tabela("acesso")
            If VarType(Tabela("tipo_acesso")) <> vbNull Then mTipoAcesso = Tabela("tipo_acesso")
            If txtTipo.text = "V" Then
                With frmVendas
                    .txtid_vendedor.text = mIdVendedor
                    .txtVendedor.text = mVendedor
                    .txtAcesso.text = mAcesso
                    .Show 1
                End With
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "L" Then
                If mTipoAcesso = "A" Then
                    With frmClientesCadastro
                        .txtid_cliente.text = txtid_cliente.text
                        .txtTipo.text = "L"
                        .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If
            
             If txtTipo.text = "Acerto" Then
                If mTipoAcesso = "A" Then
                    With frmSaidaConsulta
                         .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If


            If txtTipo.text = "F" Then
                With frmFichaFicha
                    .txtid_vendedor.text = mIdVendedor
                    .txtVendedor.text = mVendedor
                    .txtAcesso.text = mAcesso
                    .txtid_prazo.text = txtid_prazo.text
                    .txtTipo_acesso.text = mTipoAcesso
                    .Show 1
                End With
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "C" Then
                If mTipoAcesso = "P" Then
                    With frmcaixaConsulta
                        .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "R" Then
                If mTipoAcesso = "P" Then
                    With frmCaixaRetirada
                        .txtid_vendedor.text = mIdVendedor
                        .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "I" Then
                If mTipoAcesso = "P" Then
                    With frmCaixaInserir
                        .txtid_vendedor.text = mIdVendedor
                        .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "D" Then
                If mTipoAcesso = "A" Then
                    With frmCaixaDevolucao
                        .txtid_vendedor.text = mIdVendedor
                        .Show 1
                    End With
                Else
                    MsgBox ("Você, não esta autorizado para este acesso, chame o Administrador do sistema"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "A" Then
                If mTipoAcesso = "A" Then
                    confirma = MsgBox("Confirma Vendedor como Administrador do Sistema", vbQuestion + vbYesNo)
                    If confirma = vbYes Then

                        ' ---- retira o accesso para todos os vendedores
                        Sqlconsulta = " Tipo_acesso = 'A'"
                        campo = " Tipo_acesso = ''"
                        sqlAlterar "vendedores", campo, Sqlconsulta, Me, "N"
                        '---- inclui acesso para um vendedor exclusivo
                        Sqlconsulta = "id_vendedor = '" & txtid_vendedor.text & "'"
                        campo = " Tipo_acesso = 'A'"
                        sqlAlterar "vendedores", campo, Sqlconsulta, Me, "S"

                        Unload Me

                    End If
                Else
                    MsgBox ("Você, não esta autorizado para este acesso"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If

            If txtTipo.text = "P" Then
                If mTipoAcesso = "A" Then
                    confirma = MsgBox("Confirma Vendedor como Caixa do Sistema", vbQuestion + vbYesNo)
                    If confirma = vbYes Then

                        ' ---- retira o accesso para todos os vendedores
                        Sqlconsulta = " Tipo_acesso = 'P'"
                        campo = " Tipo_acesso = ''"
                        sqlAlterar "vendedores", campo, Sqlconsulta, Me, "N"
                        '---- inclui acesso para um vendedor exclusivo
                        Sqlconsulta = "id_vendedor = '" & txtid_vendedor.text & "'"
                        campo = " Tipo_acesso = 'P'"
                        sqlAlterar "vendedores", campo, Sqlconsulta, Me, "S"

                        Unload Me

                    End If
                Else
                    MsgBox ("Você, não esta autorizado para este acesso, chame o Administrador do sistema"), vbInformation
                End If
                Unload Me
                Exit Sub
            End If
            Exit Sub
        Else
            MsgBox ("Vendedor não cadastrado..."), vbInformation

        End If
    End If

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub

'-------- Acesso
Private Sub txtAcesso_GotFocus()
    txtAcesso.BackColor = &H80FFFF
End Sub
Private Sub txtAcesso_LostFocus()
    txtAcesso.BackColor = &H80000014
End Sub
Private Sub txtAcesso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then VendasPagamento
End Sub




