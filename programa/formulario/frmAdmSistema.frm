VERSION 5.00
Begin VB.Form frmAdmSistema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador do Sistema"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTipo_acesso 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtAcesso 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtVendedor 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtid_vendedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmAdmSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
