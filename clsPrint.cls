VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Do not remember where I found this piece of code.  If it is yours, let me
' know.  I will place your name and URL here.  Thank you.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants (Margin values)
' ***************************************************************************
  Private Const MARGIN_LEFT = 1440   ' Equal one inch
  Private Const MARGIN_RIGHT = 1440
  Private Const MARGIN_TOP = 1440
  Private Const MARGIN_BOTTOM = 1440


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' 18-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added second title line if needed

Public Sub PrintText(ByVal strDataToPrint As String, _
                     ByVal strTitleLine1 As String, _
            Optional ByVal strTitleLine2 As String = vbNullString)

    Dim intPosition  As Integer   ' index counter
    Dim intMaxLength As Integer   ' max output record length
    Dim strPrintLine As String    ' output print line
    
    ' Initialize first page
    DoNewPage strTitleLine1, strTitleLine2      ' print page header/footer
    strDataToPrint = strDataToPrint & Chr$(32)  ' add 1 trailing blank
    intMaxLength = 78
    
    ' Build the output data prior to dumping to the printer
    Do
        ' see if there is anything left to print
        If Len(Trim$(strDataToPrint)) = 0 Then
            Exit Do
        End If
        
        'Get next string of data
        strPrintLine = Left$(strDataToPrint, intMaxLength)
                
        ' verify we have something to print
        If Len(Trim$(strPrintLine)) = 0 Then
            Exit Do
        End If
        
        intPosition = InStrRev(strPrintLine, Chr$(32), Len(strPrintLine))
        strPrintLine = Trim$(Left$(strPrintLine, intPosition))
        
        ' Check if word will fit on this line
        If Len(strPrintLine) <= intMaxLength Then
            
            ' Resize input data string
            strDataToPrint = Trim$(Mid$(strDataToPrint, intPosition)) & Chr$(32)
            
            Printer.Print     ' Send carriage-return line-feed to printer
            
            'Check if we need to start a new page
            If Printer.CurrentY > (Printer.ScaleHeight - MARGIN_BOTTOM) Then
                DoNewPage strTitleLine1, strTitleLine2, True     ' print page header/footer
            Else
                Printer.CurrentX = MARGIN_LEFT   ' set position to left margin
            End If
                    
            Printer.Print strPrintLine;          ' Print this line on the page
            
        ElseIf Len(strPrintLine) = 0 Then
            Exit Do
        End If
    Loop
    
    ' Notify printer we are finished
    Printer.EndDoc

End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Sub DoNewPage(ByVal strTitleLine1 As String, _
             Optional ByVal strTitleLine2 As String = vbNullString, _
             Optional ByVal blnEjectPage As Boolean = False)
    
    Dim strBuffer As String
    
    ' Start new page if requested
    If blnEjectPage Then
        Printer.NewPage
    End If
    
    ' Define printer font for header/footer
    Printer.Print
    Printer.Font = "Times New Roman"
    Printer.FontSize = 12
    Printer.FontBold = True
    
    ' Print page header
    Printer.CurrentY = (MARGIN_TOP - Printer.TextHeight(strTitleLine1)) / 2
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTitleLine1)) / 2
    Printer.Print strTitleLine1;
    Printer.Print
    
    Printer.CurrentY = ((MARGIN_TOP - Printer.TextHeight(strTitleLine2)) * 1.5) / 2
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTitleLine2)) / 2
    Printer.Print strTitleLine2;
    
    ' Print page footer
    strBuffer = StrConv(Format$(Now(), "d mmm yyyy  hh:mm:ss"), vbProperCase)
    Printer.CurrentX = MARGIN_LEFT
    Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
    Printer.Print strBuffer;
    
    strBuffer = "Page  " & Printer.Page
    Printer.CurrentX = Printer.ScaleWidth - (Printer.TextWidth(strBuffer) + MARGIN_RIGHT)
    Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
    Printer.Print strBuffer;
    
    ' Reset position to top of page and redefine font for page body print.
    Printer.Print
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentX = MARGIN_LEFT
    Printer.CurrentY = MARGIN_TOP

End Sub

