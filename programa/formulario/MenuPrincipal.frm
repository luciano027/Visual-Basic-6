VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MenuPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Vendas"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgRelatorio 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar MDIStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   2850
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "ONLINE:"
            TextSave        =   "ONLINE:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MenuPrincipal.frx":0000
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Hoje:"
            TextSave        =   "Hoje:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "24/02/2021"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "20:38"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   935
            MinWidth        =   935
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menu_cadasto 
      Caption         =   "Cadastro"
      Begin VB.Menu produtos_estoque 
         Caption         =   "Produtos no Estoque"
      End
      Begin VB.Menu fornecedor 
         Caption         =   "Fornecedor"
      End
      Begin VB.Menu clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu vendedores 
         Caption         =   "Vendedores"
      End
   End
   Begin VB.Menu menu_vendas 
      Caption         =   "Vendas"
      Begin VB.Menu registrar_vendas 
         Caption         =   "Registrar Vendas"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu menu_movimento 
      Caption         =   "Movimento"
      Begin VB.Menu entrada_estoque 
         Caption         =   "Entrada no Estoque"
      End
      Begin VB.Menu saida_estoque 
         Caption         =   "Saida Acerto no Estoque"
      End
      Begin VB.Menu saida_cliente 
         Caption         =   "Consulta saida Cliente"
      End
      Begin VB.Menu movimento_cliente 
         Caption         =   "Consulta Ficha Cliente"
      End
   End
   Begin VB.Menu menu_financeiro 
      Caption         =   "Financeiro"
      Begin VB.Menu contas_pagar 
         Caption         =   "Contas a Pagar"
      End
      Begin VB.Menu contas_receber 
         Caption         =   "Ficha do Cliente"
         Shortcut        =   {F6}
      End
      Begin VB.Menu separador1 
         Caption         =   "-"
      End
      Begin VB.Menu caixa 
         Caption         =   "Caixa"
         Begin VB.Menu caixa_caixa 
            Caption         =   "Caixa"
            Shortcut        =   {F9}
         End
         Begin VB.Menu caixa_retirada 
            Caption         =   "Retirada do Caixa"
         End
         Begin VB.Menu caixa_inlcuir 
            Caption         =   "Incluir no Caixa"
         End
         Begin VB.Menu caixa_devolucao 
            Caption         =   "Devolução"
         End
      End
   End
   Begin VB.Menu menu_relatorios 
      Caption         =   "Relatorios"
   End
   Begin VB.Menu Menu_utilitarios 
      Caption         =   "Utilitarios"
      Begin VB.Menu consulta_estoque 
         Caption         =   "Consulta Estoque"
         Shortcut        =   {F4}
      End
      Begin VB.Menu agenda_telefone 
         Caption         =   "Agenda Telefone"
         Shortcut        =   {F7}
      End
      Begin VB.Menu consulta_cep 
         Caption         =   "Consulta CEP"
         Shortcut        =   {F8}
      End
      Begin VB.Menu separado 
         Caption         =   "-"
      End
      Begin VB.Menu Backup 
         Caption         =   "Backup"
      End
   End
   Begin VB.Menu Menu_sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Option Explicit

Private Declare Function LoadLibrary Lib "Kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Sub InitCommonControls Lib "ComCtl32.dll" ()
Private hShellLib As Long


Private Sub caixa_caixa_Click()
    With frmVendedorCodigo
        .txtTipo.text = "C"
        .txtAcesso.text = ""
        .Show 1
    End With
End Sub

Private Sub caixa_devolucao_Click()
    With frmVendedorCodigo
        .txtTipo.text = "D"
        .txtAcesso.text = ""
        .Show 1
    End With
End Sub

Private Sub caixa_inlcuir_Click()
    With frmVendedorCodigo
        .txtTipo.text = "I"
        .txtAcesso.text = ""
        .Show 1
    End With
End Sub

Private Sub caixa_retirada_Click()
    With frmVendedorCodigo
        .txtTipo.text = "R"
        .txtAcesso.text = ""
        .Show 1
    End With
End Sub

Private Sub menu_relatorios_Click()
    frmRelatorios.Show 1
End Sub

Private Sub movimento_cliente_Click()
    frmFichaConsultaMovimento.Show 1
End Sub

Private Sub vendedores_Click()
    frmVendedoresConsulta.Show 1
End Sub

Private Sub clientes_Click()
    frmClientesConsulta.Show 1
End Sub

Private Sub contas_pagar_Click()
    frmPagarConsulta.Show 1
End Sub

Private Sub fornecedor_Click()
    frmFornecedorConsulta.Show 1
End Sub

Private Sub produtos_estoque_Click()
    frmEstoqueConsulta.Show 1
End Sub

Private Sub saida_cliente_Click()
    frmClientesSaidaConsulta.Show 1
End Sub

Private Sub saida_estoque_Click()
    With frmVendedorCodigo
        .txtTipo.text = "Acerto"
        .txtAcesso.text = ""
        .Show 1
    End With
End Sub

Private Sub consulta_estoque_Click()
    frmConsultaEstoqueGeral.Show 1
End Sub

Private Sub contas_receber_Click()
    frmFichaConsulta.Show 1
End Sub

Private Sub entrada_estoque_Click()
    frmEntradaConsulta.Show 1
End Sub

Private Sub Menu_sair_Click()
    Unload Me
    Set MenuPrincipal = Nothing
    End
End Sub

Private Sub agenda_telefone_Click()
    frmAgendaTelefone.Show 1
End Sub

Private Sub Backup_Click()
    frmBackup_restaura.Show 1
End Sub

Private Sub consulta_cep_Click()
    frmCepConsulta.Show 1
End Sub

Private Sub registrar_vendas_Click()
    With frmVendedorCodigo
        .txtTipo.text = "V"
        .txtAcesso.text = ""
        .Show 1
    End With

End Sub

Private Sub relatorios_clientes_Click()

End Sub

Private Sub relatorios_estoque_Click()

End Sub

Private Sub relatorios_fornecedor_Click()

End Sub

Private Sub relatorios_vendedores_Click()

End Sub

'****************************************************************************************************************
'****************************************************************************************************************
'*****************************************************************************************************1***********
Private Sub MDIForm_Unload(Cancel As Integer)
    Set MenuPrincipal = Nothing
    End
End Sub

Private Sub Centralize_Back_Form()
    With fMdiBack
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
        .Left = 0
        .Top = 0
        .ZOrder 1
    End With
End Sub

Private Sub MDIForm_Activate()
    Centralize_Back_Form

    ' frmLogin.Show 1

    '  Verificar_menu_login

End Sub

Private Sub MDIForm_Load()
    On Error GoTo trata_erro



    MicroBD = ReadINI("Geral", "Maquina", App.Path & "\vendas.ini")
    TipoBD = ReadINI("Geral", "Tipo", App.Path & "\vendas.ini")    ' se e cliente ou servidor
    ICONBD = ReadINI("Geral", "NomeIco", App.Path & "\vendas.ini")    ' endereco onde se encotra a icone
    Logo = ReadINI("Geral", "Logo", App.Path & "\vendas.ini")    ' endereco onde se encotra a icone
    salvaric = ReadINI("Geral", "salvaric", App.Path & "\vendas.ini")    ' arquivo onde ira salvar as fotos dos Consultas
    ArqEmail = ReadINI("Arquivos", "ArqEmail", App.Path & "\vendas.ini")
    ArqTecnico = ReadINI("Arquivos", "ArqTecnico", App.Path & "\vendas.ini")
    ArqImprime = ReadINI("Arquivos", "ArqImprime", App.Path & "\vendas.ini")
    ArqTemp = ReadINI("Arquivos", "ArqTemp", App.Path & "\vendas.ini")
    FotoICo = ReadINI("Geral", "FotoICO", App.Path & "\vendas.ini")    ' arquivo onde ira salvar as fotos dos Consultas

    strgExtrato = ArqTecnico & "\" & "maq" & MicroBD & ".txt"

    MDIStatus.Panels.Item(4).text = "  F4-Consulta Estoque     | F5-Venda                | F6-Ficha Cliente       | F7-Agenda Telêfone      | F8-Consulta CEP | F9 - Caixa    "
    MDIStatus.Panels.Item(7).text = Format(DtaSistema, "DD/MM/YYYY")

    Set Me.Icon = LoadPicture(ICONBD)

    Conectar

    Exit Sub
trata_erro:
    Exibe_Erros (Err.Description)
End Sub



Public Sub DesabilitaMenu()
'  If mCadastro = 1 Then MenuPrincipal.menu_cadastro.Enabled = False Else MenuPrincipal.menu_cadastro.Enabled = False
'  If mAtestado = 1 Then MenuPrincipal.menu_atestado.Enabled = False Else Me.menu_atestado.Enabled = False
'  MenuPrincipal.menu_faturamento.Enabled = False
'  MenuPrincipal.menu_utilitarios.Enabled = False
End Sub

Public Sub AbilidataMenu()
' If mCadastro = 1 Then MenuPrincipal.menu_cadastro.Enabled = True Else MenuPrincipal.menu_cadastro.Enabled = False
' If mAtestado = 1 Then MenuPrincipal.menu_atestado.Enabled = True Else MenuPrincipal.menu_atestado.Enabled = False
' If mFaturamento = 1 Then MenuPrincipal.menu_faturamento.Enabled = True Else MenuPrincipal.menu_faturamento.Enabled = False
' If mUtilitarios = 1 Then MenuPrincipal.menu_utilitarios.Enabled = True Else MenuPrincipal.menu_utilitarios.Enabled = False
End Sub
