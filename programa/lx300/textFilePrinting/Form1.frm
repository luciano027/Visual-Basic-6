VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Report Printing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Report"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Please add "MicroSoft Scripting Runtime" reference in your project before run this routine

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

Private lhPrinter As Long
Private BtsKiri As String
Private FntSize As String
Private TblFnt As Boolean

Dim Fs, Glfile                 ' File scripting reference
Attribute Fs.VB_VarUserMemId = 1073938432
Attribute Glfile.VB_VarUserMemId = 1073938432
Dim fvlonLineCount As Long
Attribute fvlonLineCount.VB_VarUserMemId = 1073938434
Dim fvbooExitPrintLine As Boolean
Attribute fvbooExitPrintLine.VB_VarUserMemId = 1073938435
Dim fvintPageNo As Integer
Attribute fvintPageNo.VB_VarUserMemId = 1073938436

Sub PageHeader()
' This text file can be printed on Dot Metrix Printers like Epson LX-300 ,etc
' and other compatiable printers for fast printing than graphical report

' chr(14) for Large font
' chr(12) for page eject
' chr(18) for normal font
' chr(15) for condence printing

    PrintLine ""
    PrintLine Chr(14) & "Auther: Fayyaz Butt"
    PrintLine "For Comments Plz email to : fayyaz_a@hotmail.com"
    PrintLine "ANY HEADING GOES HERE..     " & String(20, " ") & Format(Format(Date, "mmmm d, yyyy"), String(20, "@"))
    PrintLine ""
    PrintLine String(75, "=")
    PrintLine " Code    Description "
    PrintLine String(75, "=")

End Sub

Sub PageFooter()
    fvintPageNo = fvintPageNo + 1
    PrintLine ""
    PrintLine String(75, "-")
    PrintLine " " & "Any text you want to print at footer" & String(10, " ") & "Page : " & Format(fvintPageNo, "@@@") & Chr(12)
End Sub

Sub PrintLine(AnyLine As String)
    fvlonLineCount = fvlonLineCount + 1
    Glfile.WriteLine AnyLine
End Sub

Private Sub Command1_Click()

    Dim rsGeneric As ADODB.Recordset          ' any Generic Recordset
    Dim CN As ADODB.Connection                ' set your connection

    Set CN = New ADODB.Connection
    CN.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\myDataBase.MDB"
    CN.Open

    Set rsGeneric = CN.Execute("select * from myTable")

    Dim lvstrCritaria As String
    Dim lvstrPreviousACode As String
    Dim lvintWait As Integer
    Dim lvbytCount As Byte


    fvlonLineCount = 0
    fvbooExitPrintLine = False
    fvintPageNo = 0

    Set Fs = CreateObject("scripting.filesystemobject")
    Set Glfile = Fs.CreateTextFile(App.Path & "\MyTextfile.txt", True)
    '*** Report Header ***
    Call PageHeader
    '*** Detail Section ***
    rsGeneric.MoveFirst
    Do Until rsGeneric.EOF
        ' set your table fields here
        PrintLine rsGeneric!acno & " | " & rsGeneric!head
        rsGeneric.MoveNext
    Loop
    '*** Report Footer ***
    PrintLine String(75, "-")
    PrintLine "*** End of Report ***"
    Call PageFooter
    '*** End Of Report ***

    Glfile.Close
    Screen.MousePointer = 0
    MsgBox "Your report has been printed to File" & Chr(13) & App.Path & "\MyTextfile.TXT", vbInformation, "Report Printed"
    CN.Close
    Set CN = Nothing

End Sub


Private Sub Command2_Click()
'This is the sample how to print in Dos Mode and How to print it Bold!
'http://www.geocities.com/mdgnn/xcontrols.htm

    Open "Lpt1" For Output As #1

    Print #1, Chr(27) & "@"    'Initialize printer
    Print #1, Chr(27) & "A" & Chr(11)
    Print #1, Chr(27) & "E"    'Set Font Bold
    Print #1, "Printer is Bold"
    Print #1, "Printer is Bold"
    Print #1, Chr(27) & "F"    ' Set Font Normal

    Print #1, "Printer is Normal"
    Print #1, "Printer is Normal"

    Close #1
End Sub

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type


Public Property Get BatasKiri() As String
    BatasKiri = BtsKiri
End Property
Public Property Let BatasKiri(ByVal StrBtsNew As String)
    BtsKiri = Space(Val(StrBtsNew))
End Property
Public Property Get FontSize() As String
    FontSize = FntSize
End Property
Public Property Let FontSize(ByVal StrFntNew As String)
    FntSize = Chr$(27) + Chr$(StrFntNew)
End Property
Public Property Get TebalFont() As Boolean
    TebalFont = TblFnt
End Property
Public Property Let TebalFont(ByVal BTblFntNew As Boolean)
    TblFnt = BTblFntNew
End Property

Private Sub Class_Initialize()
    Dim lReturn As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)

    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = "Kopebi"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    If Printer.DriverName <> "SRP250" Then
        BtsKiri = Space(5)
        FntSize = Chr$(27) + Chr$(80)
    End If
    TblFnt = False
End Sub

Private Sub Class_Terminate()
    Dim lReturn As Long
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
End Sub

Public Sub PrintData(ByVal sWrittenData As String)
    Dim lReturn As Long
    Dim lpcWritten As Long

    If Printer.DriverName <> "SRP250" Then
        If TblFnt = False Then
            'Chr$(27) + Chr$(15) = CONDESED
            'Chr$(27) + Chr$(120) + "0" = DRAFT MODE
            sWrittenData = Chr$(27) + Chr$(120) + "0" + Chr$(27) + Chr$(15) + FntSize + BtsKiri & sWrittenData & vbCrLf
        Else
            sWrittenData = Chr$(27) + Chr$(120) + "0" + Chr$(27) + Chr$(15) + Chr$(27) + Chr$(71) + BtsKiri + FntSize & sWrittenData & vbCrLf
        End If
    Else
        sWrittenData = sWrittenData & vbCrLf
    End If
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
End Sub

Public Sub PrintHead(ByVal sWrittenData As String)
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim temp As String
    temp = Chr$(27) + Chr$(120) + "0" + Chr$(27) + Chr$(15) + FntSize & ""
    lReturn = WritePrinter(lhPrinter, ByVal temp, Len(temp), lpcWritten)
    sWrittenData = Chr$(27) + Chr$(120) + "0" + Chr$(27) + Chr$(71) + Chr$(27) + Chr$(14) + FntSize & sWrittenData & vbCrLf
    'sWrittenData = Chr$(27) + Chr$(52) + Chr$(14) + FntSize & sWrittenData & Chr$(27) + Chr$(53) + vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
End Sub

