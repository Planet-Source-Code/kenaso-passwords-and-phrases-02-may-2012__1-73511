VERSION 5.00
Begin VB.Form frmSpecial 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   6555
   ClipControls    =   0   'False
   Icon            =   "frmSpecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6555
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   225
      TabIndex        =   32
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   225
      TabIndex        =   31
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   " @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   225
      TabIndex        =   30
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   " #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   225
      TabIndex        =   29
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   " $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   1080
      TabIndex        =   28
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   " %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   1080
      TabIndex        =   27
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   1080
      TabIndex        =   26
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   1080
      TabIndex        =   25
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ("
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   8
      Left            =   1935
      TabIndex        =   24
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   9
      Left            =   1935
      TabIndex        =   23
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   10
      Left            =   1935
      TabIndex        =   22
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  _"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   11
      Left            =   1935
      TabIndex        =   21
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   12
      Left            =   2790
      TabIndex        =   20
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   13
      Left            =   2790
      TabIndex        =   19
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ["
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   14
      Left            =   2790
      TabIndex        =   18
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   15
      Left            =   2790
      TabIndex        =   17
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  {"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   16
      Left            =   3645
      TabIndex        =   16
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  }"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   17
      Left            =   3645
      TabIndex        =   15
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  \"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   18
      Left            =   3645
      TabIndex        =   14
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  /"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   19
      Left            =   3645
      TabIndex        =   13
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  |"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   20
      Left            =   4500
      TabIndex        =   12
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ;"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   21
      Left            =   4500
      TabIndex        =   11
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   22
      Left            =   4500
      TabIndex        =   10
      Top             =   1980
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  <"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   23
      Left            =   4500
      TabIndex        =   9
      Top             =   2610
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   24
      Left            =   5355
      TabIndex        =   8
      Top             =   720
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   25
      Left            =   5355
      TabIndex        =   7
      Top             =   1350
      Width           =   750
   End
   Begin VB.CheckBox chkFlags 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 0 - 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   26
      Left            =   5355
      TabIndex        =   6
      Top             =   1980
      Width           =   1020
   End
   Begin VB.CommandButton cmdFlags 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   5535
      TabIndex        =   4
      Top             =   3330
      Width           =   870
   End
   Begin VB.CommandButton cmdFlags 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   4590
      TabIndex        =   3
      Top             =   3330
      Width           =   870
   End
   Begin VB.CommandButton cmdFlags 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   3645
      TabIndex        =   2
      Top             =   3330
      Width           =   870
   End
   Begin VB.CommandButton cmdFlags 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   2700
      TabIndex        =   1
      Top             =   3330
      Width           =   870
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   3510
      Width           =   1290
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   6045
   End
End
Attribute VB_Name = "frmSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmSpecial
'
' NOTE:          If you want to step thru the code, do the follwing:
'
'                     1.  Navigate to Form_Load() in frmMain
'                     2.  Comment out the following line.
'
'                             WheelHook frmMain.hwnd
'
'               This will deactivate the wheel scrolling capability while
'               you are in the VB IDE.  If you do not do this, you will
'               have great difficulty walking thru the code.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected loading of special character array
' ***************************************************************************
Option Explicit

Private Sub cmdFlags_Click(Index As Integer)

    Dim bytIndex As Byte
        
    Select Case Index
            
           Case 0  ' Reset flags
                Erase gastrChars()   ' Always start with an empty array
    
                With frmSpecial
                    ' Loop thru and reset values to unchecked
                    For bytIndex = 0 To 25
                        .chkFlags(bytIndex).Value = vbChecked
                        gastrChars(bytIndex) = Trim$(.chkFlags(bytIndex).Caption)
                    Next bytIndex
                    
                    ' Uncheck the numbers box
                    .chkFlags(26).Value = vbChecked
                                        
                    ' Insert numbers into array
                    For bytIndex = 26 To 35
                        gastrChars(bytIndex) = CStr(bytIndex - 26)
                    Next bytIndex
        
                End With
                
                Exit Sub
    
           Case 1  ' Save checked flags
                Erase gastrChars()   ' Always start with an empty array
    
                With frmSpecial
                    If gblnSpecialWithNbrs Then
                            
                        ' Symbols and numbers
                        ' Loop thru and save values to an array
                        For bytIndex = 0 To 25
                            If CBool(.chkFlags(bytIndex).Value) Then
                                gastrChars(bytIndex) = Trim$(.chkFlags(bytIndex).Caption)
                            Else
                                gastrChars(bytIndex) = vbNullString
                            End If
                        Next bytIndex
                        
                        ' Is numbers box checked?
                        If CBool(.chkFlags(26).Value) Then
                            ' Insert numbers into array
                            For bytIndex = 26 To 35
                                gastrChars(bytIndex) = CStr(bytIndex - 26)
                            Next bytIndex
                        Else
                            ' Remove numbers from array
                            For bytIndex = 26 To 35
                                gastrChars(bytIndex) = vbNullString
                            Next bytIndex
                        End If
                    Else
                        ' Symbols only
                        ' Loop thru and save values to an array
                        For bytIndex = 0 To 25
                            
                            If CBool(.chkFlags(bytIndex).Value) Then
                                gastrChars(bytIndex) = Trim$(.chkFlags(bytIndex).Caption)
                            Else
                                gastrChars(bytIndex) = vbNullString
                            End If
                            
                        Next bytIndex
                        
                        ' Uncheck the numbers box
                        .chkFlags(26).Value = vbUnchecked
                        
                        ' Remove numbers from array
                        For bytIndex = 26 To 35
                            gastrChars(bytIndex) = vbNullString
                        Next bytIndex
                    End If
                End With
    
    End Select
    
    ' Return to main screen
    frmSpecial.Hide
    frmMain.Show

End Sub

Private Sub Form_Load()

    Dim bytIndex As Byte
    Dim strMsg   As String
    Dim objEdit  As cKeyEdit
    
    strMsg = "Remove checkmark from characters " & _
             "that will not be used in a password."
    
    With frmSpecial
        .Hide
        .Caption = "Special characters"
        .lblInfo.Caption = strMsg
        
        ' Loop thru and reset values to unchecked
        For bytIndex = 0 To 25
            .chkFlags(bytIndex).Value = vbChecked
            gastrChars(bytIndex) = Trim$(.chkFlags(bytIndex).Caption)
        Next bytIndex
        
        .chkFlags(26).Value = vbChecked
        
        For bytIndex = 26 To 35
            gastrChars(bytIndex) = CStr(bytIndex - 26)
        Next bytIndex
        
        ' Center caption on form
        Set objEdit = New cKeyEdit
        objEdit.CenterCaption frmSpecial
        Set objEdit = Nothing
        
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With
        
End Sub

Public Sub ShowForm()

    Dim bytIndex As Byte
        
    With frmSpecial
                
        ' Set checkmarks for special characters
        For bytIndex = 0 To 25
            
            ' Loop thru and flag appropriate items
            If .chkFlags(bytIndex).Enabled Then
                
                If Len(gastrChars(bytIndex)) = 0 Then
                    .chkFlags(bytIndex).Value = vbUnchecked
                Else
                    .chkFlags(bytIndex).Value = vbChecked
                End If
            
            End If
        
        Next bytIndex
    
        ' Enable or disable numbers
        If gblnSpecialWithNbrs Then
        
            ' using numbers
            .chkFlags(26).Visible = True
            .chkFlags(26).Value = vbChecked
    
            For bytIndex = 26 To 35
                gastrChars(bytIndex) = CStr(bytIndex - 26)
            Next bytIndex
        
        Else
            ' Not using numbers
            .chkFlags(26).Value = vbUnchecked
            .chkFlags(26).Visible = False
    
            For bytIndex = 26 To 35
                gastrChars(bytIndex) = vbNullString
            Next bytIndex
        
        End If
    End With
    
    frmMain.Hide     ' Hide main form
    frmSpecial.Show  ' Show this form

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' "X" selected in upper right
    ' corner. Return to main screen.
    If UnloadMode = 0 Then
        frmSpecial.Hide  ' Hide this form
        frmMain.Show     ' Show main form
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

