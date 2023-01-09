VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscaSimples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmBuscaSimples.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14888
            MinWidth        =   14888
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Gif89a1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbEntity 
      Height          =   4860
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   5160
      Index           =   1
      Left            =   0
      Picture         =   "frmBuscaSimples.frx":146E2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   5640
   End
End
Attribute VB_Name = "frmBuscaSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim vKey As Integer

Private Sub Form_Load()
    Me.Width = 5415
    Me.Height = 7140
    Centerform Me
    Set Me.Icon = LoadPicture(ICONBD)
End Sub



Private Sub cmbEntity_DblClick()
    cmbEntity_KeyPress 13
End Sub

Private Sub cmbEntity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13     'enter
        If Not cmbEntity.ListIndex = -1 Then
            vKey = cmbEntity.ItemData(cmbEntity.ListIndex)
            Unload Me
        End If

    Case 27     'Esc
        vKey = -1
        Unload Me
    End Select
End Sub

Public Function getKey(ByVal pTblName As String, ByVal pFldName As String, Optional pCriteria As String)
    vKey = -1
    Call fillCombo(cmbEntity, pTblName, pFldName, pCriteria)
    Me.Show vbModal

    getKey = vKey
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmBuscaSimples = Nothing
End Sub
