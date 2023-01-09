VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   3930
   ClientTop       =   4470
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4530
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   795
      Left            =   780
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Customer List"
      Height          =   795
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
Private Sub cmdPrint_Click()
'-----------------------------------------------------------------------------
    
    Dim intLineCtr          As Integer
    Dim intPageCtr          As Integer
    Dim intX                As Integer
    Dim strCustFileName     As String
    Dim strBackSlash        As String
    Dim intCustFileNbr      As Integer
    
    Dim strFirstName        As String
    Dim strLastName         As String
    Dim strAddr             As String
    Dim strCity             As String
    Dim strState            As String
    Dim strZip              As String

    Const intLINE_START_POS As Integer = 6
    Const intLINES_PER_PAGE As Integer = 60
    
    ' Have the user make sure his/her printer is ready ...
    If MsgBox("Make sure your printer is on-line and " _
            & "loaded with paper.", vbOKCancel, "Check Printer") = vbCancel _
    Then
        Exit Sub
    End If
    
    ' Set the printer font to Courier, if available (otherwise, we would be
    ' relying on the default font for the Windows printer, which may or
    ' may not be set to an appropriate font) ...
    For intX = 0 To Printer.FontCount - 1
        If Printer.Fonts(intX) Like "Courier*" Then
            Printer.FontName = Printer.Fonts(intX)
            Exit For
        End If
    Next
    
    Printer.FontSize = 10
    
    ' initialize report variables ...
    intPageCtr = 0
    intLineCtr = 99 ' initialize line counter to an arbitrarily high number
                    ' to force the first page break
                    
    ' prepare file name & number
    strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
    strCustFileName = App.Path & strBackSlash & "customer.txt"
    intCustFileNbr = FreeFile

    ' open the input file
    Open strCustFileName For Input As #intCustFileNbr
    
    ' read and print all the records in the input file
    Do Until EOF(intCustFileNbr)
        ' read a record from the input file and store the fields there into VB variables
        Input #intCustFileNbr, strLastName, strFirstName, strAddr, strCity, strState, strZip
        ' if the number of lines printed so far exceeds the maximum number of lines
        ' allowed on a page, invoke the PrintHeadings subroutine to do a page break
        If intLineCtr > intLINES_PER_PAGE Then
            GoSub PrintHeadings
        End If
        ' print a line of data
        Printer.Print Tab(intLINE_START_POS); _
                      strFirstName & " " & strLastName; _
                      Tab(21 + intLINE_START_POS); _
                      strAddr; _
                      Tab(48 + intLINE_START_POS); _
                      strCity; _
                      Tab(72 + intLINE_START_POS); _
                      strState; _
                      Tab(76 + intLINE_START_POS); _
                      strZip
        ' increment the line count
        intLineCtr = intLineCtr + 1
    Loop

    ' close the input file
    Close #intCustFileNbr

    ' Important! When done, the EndDoc method of the Printer object must be invoked.
    ' The EndDoc method terminates a print operation sent to the Printer object,
    ' releasing the document to the print device or spooler.
    Printer.EndDoc
    
    cmdExit.SetFocus
    
    Exit Sub


' internal subroutine to print report headings
'------------
PrintHeadings:
'------------
    ' If we are about to print any page other than the first, invoke the NewPage
    ' method to perform a page break. The NewPage method advances to the next
    ' printer page and resets the print position to the upper-left corner of the
    ' new page.
    If intPageCtr > 0 Then
        Printer.NewPage
    End If
    ' increment the page counter
    intPageCtr = intPageCtr + 1
    
    ' Print 4 blank lines, which provides a for top margin. These four lines do NOT
    ' count toward the limit of 60 lines.
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    
    ' Print the main headings
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Date: "; _
                  Format$(Date, "mm/dd/yy"); _
                  Tab(intLINE_START_POS + 31); _
                  "THE VBPROGRAMMER.COM"; _
                  Tab(intLINE_START_POS + 73); _
                  "Page:"; _
                  Format$(intPageCtr, "@@@")
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Time: "; _
                  Format$(Time, "hh:nn:ss"); _
                  Tab(intLINE_START_POS + 33); _
                  "CUSTOMER LIST"
    Printer.Print
    ' Print the column headings
    Printer.Print Tab(intLINE_START_POS); _
                  "CUSTOMER NAME"; _
                  Tab(21 + intLINE_START_POS); _
                  "ADDRESS"; _
                  Tab(48 + intLINE_START_POS); _
                  "CITY"; _
                  Tab(72 + intLINE_START_POS); _
                  "ST"; _
                  Tab(76 + intLINE_START_POS); _
                  "ZIP"
    Printer.Print Tab(intLINE_START_POS); _
                  "-------------"; _
                  Tab(21 + intLINE_START_POS); _
                  "-------"; _
                  Tab(48 + intLINE_START_POS); _
                  "----"; _
                  Tab(72 + intLINE_START_POS); _
                  "--"; _
                  Tab(76 + intLINE_START_POS); _
                  "---"
    Printer.Print
    ' reset the line counter to reflect the number of lines that have now
    ' been printed on the new page.
    intLineCtr = 6
    Return

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdExit_Click()
'-----------------------------------------------------------------------------
    End
End Sub

