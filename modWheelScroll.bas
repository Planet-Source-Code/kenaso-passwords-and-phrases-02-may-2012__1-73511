Attribute VB_Name = "modWheelScroll"
' ***************************************************************************
' Routine:    modWheelScroll
'
' Purpose:    Wheel mouse support
'
' Reference:  Joe Fisher  27-Mar-2008
'             MSHFlexGrid Select individual/multiple rows and Export to
'             Excel/Also Search Text in MSHFlexGrid
'             http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=70322&lngWId=1
'
' Warning:    Just a note of caution. This solution makes use of a "hook"
'             into the Windows message stream directed at your program
'             form. If you introduce an error into the WindowProc()
'             function (detailed below) then you will crash the Visual
'             Basic IDE. Please make sure that you save your program
'             before testing and that you try and eliminate any errors in
'             the specified routine. Once up and running this solution is
'             entirely stable.
'
'             To activate the hook into the Windows message stream that
'             detects the mouse wheel "event" you should call the WheelHook()
'             routine from the relevant Form Activate event. You should also
'             remember to call the WheelUnHook() routine from the Deactivate
'             event. This cleans up by deactivating the hook into the relevant
'             message stream but also means that you can apply this technique
'             to multiple forms in the same application.
'
' NOTE:       If you want to step thru the code, do the following:
'
'                 1.  Navigate to Form_Load() in frmMain
'                 2.  Comment out the following line:
'
'                        WheelHook frmMain.hwnd
'
'             This will deactivate the wheel scrolling capability while
'             you are in the VB IDE.  If you do not do this, you will
'             have great difficulty walking thru the code.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Created module
' 13-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug in grid scrolling back to top of grid.  Thanks to
'              GioRock for spotting this.
' 18-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug in grid scrolling back to top of grid if there were
'              fixed rows. Also updated left and right mouse scrolling.
'              See bottom of module.
' ***************************************************************************
Option Explicit
' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const MK_CONTROL As Long = &H8
  Public Const MK_LBUTTON As Long = &H1
  Public Const MK_RBUTTON As Long = &H2
  Public Const MK_MBUTTON As Long = &H10
  Public Const MK_SHIFT   As Long = &H4
  
' ***************************************************************************
' Module constants
' ***************************************************************************
  Private Const PREV_WND_PROC      As String = "PrevWndProc"
  Private Const WM_DESTROY         As Long = &H2
  Private Const WM_MOUSEWHEEL      As Long = &H20A
  Private Const GWL_WNDPROC        As Long = -4
  Private Const CB_GETDROPPEDSTATE As Long = &H157

' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type RECT
      Left   As Long
      Top    As Long
      Right  As Long
      Bottom As Long
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CallWindowProc function passes message information to the specified
  ' window procedure. Use the CallWindowProc function for window subclassing.
  ' Usually, all windows with the same class share one window procedure.  A
  ' subclass is a window or set of windows with the same class whose messages
  ' are intercepted and processed by another window procedure (or Procedures)
  ' before being passed to the window procedure of the class
  Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
          (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long

  ' The SetWindowLong function changes an attribute of the specified window.
  ' The function also sets the 32-bit (long) value at the specified offset
  ' into the extra window memory.
  Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
          (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
  ' Sends the specified message to a window or windows. The SendMessage
  ' function calls the window procedure for the specified window and does
  ' not return until the window procedure has processed the message.
  Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
          (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, _
          lParam As Any) As Long

  ' The GetWindowRect function retrieves the dimensions of the bounding
  ' rectangle of the specified window. The dimensions are given in screen
  ' coordinates that are relative to the upper-left corner of the screen.
  Private Declare Function GetWindowRect Lib "user32" _
          (ByVal hwnd As Long, lpRect As RECT) As Long
                
  ' The GetProp function retrieves a data handle from the property list
  ' of the specified window. The character string identifies the handle
  ' to be retrieved. The string and handle must have been added to the
  ' property list by a previous call to the SetProp function.
  Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" _
          (ByVal hwnd As Long, ByVal lpString As String) As Long

  ' The SetProp function adds a new entry or changes an existing entry
  ' in the property list of the specified window. The function adds a
  ' new entry to the list if the specified character string does not
  ' exist already in the list. The new entry contains the string and the
  ' handle. Otherwise, the function replaces the string's current handle
  ' with the specified handle.
  Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" _
          (ByVal hwnd As Long, ByVal lpString As String, _
          ByVal hData As Long) As Long
  
  ' The RemoveProp function removes an entry from the property list of
  ' the specified window. The specified character string identifies the
  ' entry to be removed.
  Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" _
          (ByVal hwnd As Long, ByVal lpString As String) As Long

  ' The GetParent function retrieves a handle to the specified window's
  ' parent or owner.
  Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

Public Sub WheelHook(ByVal lngHandleID As Long)

    ' Normally called from Form_Load() event
    
    On Error Resume Next
    SetProp lngHandleID, PREV_WND_PROC, SetWindowLong(lngHandleID, GWL_WNDPROC, AddressOf WindowProc)
    On Error GoTo 0
    
End Sub

Public Sub WheelUnHook(ByVal lngHandleID As Long)
    
    ' Normally called from Form_QueryUnload() event
    
    On Error Resume Next
    SetWindowLong lngHandleID, GWL_WNDPROC, GetProp(lngHandleID, PREV_WND_PROC)
    RemoveProp lngHandleID, PREV_WND_PROC
    lngHandleID = 0
    On Error GoTo 0

End Sub

Public Function IsOver(ByVal lngHandleID As Long, _
                       ByVal lngXpos As Long, _
                       ByVal lngYpos As Long) As Boolean
  
    ' Called by MouseWheel() in Form
    '           WindowProc() in this module
    
    Dim typRECT As RECT
    
    GetWindowRect lngHandleID, typRECT
    
    With typRECT
        
        If lngXpos >= .Left And _
           lngXpos <= .Right And _
           lngYpos >= .Top And _
           lngYpos <= .Bottom Then
            
            IsOver = True
        Else
            IsOver = False
        End If
        
    End With
  
End Function

Public Sub FlexGridScroll(ByRef FG As MSFlexGrid, _
                          ByVal lngRotation As Long)
  
    ' This piece code is for the mouse wheel.
    ' For the left and right mouse click scroll
    ' go to bottom of this module.
    
    Dim lngNewValue As Long
    Dim sngStep     As Single
    
    On Error Resume Next
    
    With FG
        sngStep = .Height / .RowHeight(0)
        sngStep = Int(sngStep)
        
        If .Rows < sngStep Then
            Exit Sub
        End If
            
        Do While Not (.RowIsVisible(.TopRow + sngStep))
            sngStep = sngStep - 1
        Loop
        
        If lngRotation > 0 Then
        
            lngNewValue = .TopRow - sngStep
            
            If lngNewValue < 1 Then
                If .FixedRows > 0 Then
                    lngNewValue = .FixedRows
                Else
                    lngNewValue = 0
                End If
                
                .Row = lngNewValue
            End If
        Else
            lngNewValue = .TopRow + sngStep
            
            If lngNewValue > .Rows - 1 Then
                lngNewValue = .Rows - 1
            End If
        End If
        
        .TopRow = lngNewValue
    End With

    On Error GoTo 0

End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function WindowProc(ByVal lngHandleID As Long, _
                            ByVal lngMsg As Long, _
                            ByVal lngWparam As Long, _
                            ByVal lngLparam As Long) As Long

    Dim lngMouseKeys As Long
    Dim lngRotation  As Long
    Dim lngXpos      As Long
    Dim lngYpos      As Long
    Dim frm          As Form

    Select Case lngMsg
    
           Case WM_MOUSEWHEEL
                lngMouseKeys = lngWparam And 65535
                lngRotation = lngWparam / 65536
                lngXpos = lngLparam And 65535
                lngYpos = lngLparam / 65536
        
                Set frm = GetForm(lngHandleID)
                        
                If frm Is Nothing Then
                   
                    ' it's not a form
                    If Not IsOver(lngHandleID, lngXpos, lngYpos) And _
                       IsOver(GetParent(lngHandleID), lngXpos, lngYpos) Then
                       
                        ' it's not over the control and is over the form,
                        ' so fire mousewheel on form (if it's not a dropped down combo)
                        If SendMessage(lngHandleID, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
                            
                            GetForm(GetParent(lngHandleID)).MouseWheel lngMouseKeys, _
                                                                       lngRotation, _
                                                                       lngXpos, lngYpos
                            Exit Function ' Discard scroll message to control
                        
                        End If
                    End If
                
                Else
                    ' it's a form so fire mousewheel
                    If IsOver(frm.hwnd, lngXpos, lngYpos) Then
                        frm.MouseWheel lngRotation, lngXpos, lngYpos
                    End If
                End If
                   
           Case WM_DESTROY
                ' PREV_WND_PROC will be gone after UnSubClass is called!
                If CBool(CallWindowProc(GetProp(lngHandleID, PREV_WND_PROC), _
                                        lngHandleID, lngMsg, lngWparam, lngLparam)) Then
                    
                    WheelUnHook lngHandleID
                End If
                
                Exit Function
    End Select
    
    WindowProc = CallWindowProc(GetProp(lngHandleID, PREV_WND_PROC), _
                                lngHandleID, lngMsg, lngWparam, lngLparam)

End Function
  
Private Function GetForm(ByVal lngHandleID As Long) As Form
    
    For Each GetForm In Forms
    
        If GetForm.hwnd = lngHandleID Then
            Exit Function
        End If
    
    Next GetForm
    
    Set GetForm = Nothing

End Function


' ***************************************************************************
' ****            MSFlexgrid right & left mouse scroll                   ****
' ***************************************************************************

' 18-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
' Insert this code into the MSFlexgrid MouseUp event.  When the user holds
' down the CTRL key and clicks the left or right mouse button, the grid will
' scroll accordingly to the left or right.  Right now the number of columns
' to move is three (3).  This can be changed to what ever value you want.
' You could also move this code to operate under any of the other special
' keys as shown below.
'
' Reference:  VB6 IDE help "Detecting SHIFT, CTRL, and ALT States"

'Private Sub MSFlexgrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    Dim ShiftTest  As Integer
'
'    Static sintPos As Integer
'
'    Const COL_SHIFT As Long = 3  ' Number of columns to jump
'
'    '---------------------------------------
'    ' Hold down the appropriate key
'    ' Shift = 1  Shift key
'    ' Shift = 2  Control key
'    ' Shift = 3  SHIFT and CTRL keys
'    ' Shift = 4  ALT key
'    ' Shift = 5  SHIFT and ALT keys
'    ' Shift = 6  CTRL and ALT keys
'    ' Shift = 7  SHIFT, CTRL, and ALT keys
'    '---------------------------------------
'    ShiftTest = Shift And 7  ' Calculate special key
'
'    Select Case ShiftTest
'           Case 1  ' SHIFT key
'           Case 2  ' CTRL key
'
'                ' This piece of code could easily be moved under
'                ' one of the other special keys in this section.
'                With grdTest
'                    ' Left mouse button
'                    If Button = 1 Then
'                        sintPos = sintPos - COL_SHIFT   ' Calculate jump to new column
'
'                        ' If new position is greater than
'                        ' the left most column by x columns
'                        ' then move left x columns
'                        If sintPos > (.LeftCol + COL_SHIFT) Then
'                            .Col = sintPos   ' Jump x columns to the left
'                        Else
'                            ' See if there are fixed columns
'                            If .FixedCols > 0 Then
'
'                                ' if the new position is less than
'                                ' or equal to the number of fixed
'                                ' columns then move to the first
'                                ' non-fixed column.
'                                If sintPos <= .FixedCols Then
'                                    .Col = .FixedCols  ' first non-fixed column
'                                Else
'                                    .Col = sintPos     ' Move left to new column
'                                End If
'                            Else
'                                ' If no fixed columns and the new
'                                ' position is less than or equal
'                                ' to zero then move tothe first
'                                ' column.
'                                If sintPos <= 0 Then
'                                    .Col = 0        ' Move to first column
'                                Else
'                                    .Col = sintPos  ' Move left to new column
'                                End If
'                            End If
'
'                            sintPos = .Col     ' Save new column position
'                            .LeftCol = .Col    ' This is the new left-most column
'
'                        End If
'                    End If
'
'                    ' Right mouse button
'                    If Button = 2 Then
'                        sintPos = sintPos + COL_SHIFT  ' Calculate jump to new column
'
'                        ' If new position is greater than
'                        ' or equal to the last column then
'                        ' move to last column
'                        If sintPos >= (.Cols - 1) Then
'                            .Col = .Cols - 1   ' Move to last column
'                        Else
'                            .Col = sintPos     ' Move to right to new column
'                        End If
'
'                        sintPos = .Col         ' Save new column position
'                        .LeftCol = .Col        ' This is the new left-most column
'                    End If
'                End With
'
'           Case 3  ' SHIFT and CTRL keys
'           Case 4  ' ALT key
'           Case 5  ' SHIFT and ALT keys
'           Case 6  ' CTRL and ALT keys
'           Case 7  ' SHIFT, CTRL, and ALT keys
'    End Select
'
'End Sub


