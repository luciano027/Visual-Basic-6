VERSION 5.00
Begin VB.Form fMdiBack 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin VB.Image ImgBack 
      Height          =   7020
      Left            =   0
      Picture         =   "fMdiBack.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11940
   End
End
Attribute VB_Name = "fMdiBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this is to avoid or minimize SubResize reentrance flicker
Private aTop As Long
Private aLef As Long
Private aWid As Long
Private aHei As Long
'this is a variable text existence only requested;
'supposed Label is Logo containned:
Private aTopDiff As Long
Private aLeftDiff As Long

'- To avoid click effects
'
Private Sub Form_Click()
    Me.ZOrder 1
End Sub

Private Sub Form_DblClick()
    Me.ZOrder 1
End Sub


Private Sub Form_Load()
' getting initial measures and position
    aTop = Me.Top
    aLef = Me.Left
    aWid = Me.Width
    aHei = Me.Height

    ' getting relative lbModulo positionning

End Sub

Private Sub Arrange_Me()
    If aTop = Me.Top And _
       aLef = Me.Left And _
       aWid = Me.Width And _
       aHei = Me.Height Then
        With ImgBack
            .Top = 0
            .Left = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight
        End With
        DoEvents
        DoEvents
        DoEvents
    End If
End Sub

Private Sub Form_Resize()
    If Not aTop = Me.Top Then
        aTop = Me.Top
        Arrange_Me
    End If
    If Not aLef = Me.Left Then
        aLef = Me.Left
        Arrange_Me
    End If
    If Not aWid = Me.Width Then
        aWid = Me.Width
        Arrange_Me
    End If
    If Not aHei = Me.Height Then
        aHei = Me.Height
        Arrange_Me
    End If
End Sub



