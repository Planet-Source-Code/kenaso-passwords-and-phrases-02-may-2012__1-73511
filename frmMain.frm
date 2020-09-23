VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6165
   ClientLeft      =   1860
   ClientTop       =   2700
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8655
   Begin VB.CheckBox chkUseNumbers 
      Caption         =   "Make one word numeric if creating more than two words"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3285
      TabIndex        =   31
      Top             =   5490
      Width           =   2850
   End
   Begin VB.CheckBox chkSortData 
      Caption         =   "Sort passwords"
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
      Left            =   3300
      TabIndex        =   29
      Top             =   5535
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   225
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtGridEdit 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   990
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "txtGridEdit - Hidden"
      Top             =   5175
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox chkPassword 
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
      Left            =   3285
      TabIndex        =   23
      Top             =   5175
      Width           =   3135
   End
   Begin VB.PictureBox picWorking 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   1943
      ScaleHeight     =   1200
      ScaleWidth      =   4695
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   4755
      Begin VB.Label lblWorking 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Working"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   4410
      End
   End
   Begin VB.Frame fraPasswords 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   8475
      Begin MSFlexGridLib.MSFlexGrid grdPasswords 
         Height          =   2730
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2220
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   4815
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   7740
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   20
         Top             =   300
         Width           =   510
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   225
         Picture         =   "frmMain.frx":0614
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   300
         Width           =   510
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1148
         TabIndex        =   8
         Top             =   915
         Width           =   6315
         Begin VB.ComboBox cboSpecial 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   675
            Width           =   2340
         End
         Begin VB.TextBox txtQuantity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   300
            MaxLength       =   5
            TabIndex        =   0
            Top             =   675
            Width           =   1365
         End
         Begin VB.ComboBox cboNbrCount 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4635
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   675
            Width           =   1560
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Special Considerations"
            Height          =   240
            Left            =   2250
            TabIndex        =   18
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Password count Min-1  Max-5000"
            Height          =   495
            Index           =   1
            Left            =   225
            TabIndex        =   17
            Top             =   225
            Width           =   1530
         End
         Begin VB.Label lblNbrCount 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "  Number of letters per password to convert"
            Height          =   465
            Left            =   4455
            TabIndex        =   11
            Top             =   225
            Width           =   1905
         End
      End
      Begin VB.Label lblPasswordTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Passwords"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   2730
         TabIndex        =   21
         Top             =   105
         Width           =   3150
      End
   End
   Begin VB.Frame fraPhrase 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   90
      TabIndex        =   5
      Top             =   0
      Width           =   8475
      Begin VB.TextBox txtPassphrase 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         HideSelection   =   0   'False
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2205
         Width           =   8205
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   7740
         Picture         =   "frmMain.frx":091E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   300
         Width           =   510
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   225
         Picture         =   "frmMain.frx":0C28
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   300
         Width           =   510
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1148
         TabIndex        =   6
         Top             =   915
         Width           =   6315
         Begin VB.TextBox txtNbrOfWords 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   28
            Text            =   "99"
            Top             =   675
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of words per phrase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2325
            TabIndex        =   12
            Top             =   195
            Width           =   1650
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label lblPassPhrase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   2625
         TabIndex        =   15
         Top             =   105
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdChoice 
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
      Height          =   600
      Index           =   1
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5355
      Width           =   855
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5355
      Width           =   855
   End
   Begin VB.Label lblHidden 
      BackColor       =   &H00FFFFC0&
      Caption         =   "lblHidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblDupes 
      BackStyle       =   0  'Transparent
      Caption         =   "9999 Duplicates removed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      TabIndex        =   27
      Top             =   5880
      Width           =   2670
   End
   Begin VB.Label lblWarningMsg 
      BackStyle       =   0  'Transparent
      Height          =   675
      Left            =   180
      TabIndex        =   26
      Top             =   5160
      Width           =   2445
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   180
      TabIndex        =   24
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptLength 
         Caption         =   "&Min word length"
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "3 Chars"
            Index           =   0
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "4 Chars"
            Index           =   1
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "5 Chars"
            Index           =   2
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "6 Chars"
            Index           =   3
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "7 Chars"
            Index           =   4
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "8 Chars"
            Index           =   5
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "9 Chars"
            Index           =   6
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "10 Chars"
            Index           =   7
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "11 Chars"
            Index           =   8
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "12 Chars"
            Index           =   9
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "13 Chars"
            Index           =   10
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "14 Chars"
            Index           =   11
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "15 Chars"
            Index           =   12
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "16 Chars"
            Index           =   13
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "17 Chars"
            Index           =   14
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "18 Chars"
            Index           =   15
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "19 Chars"
            Index           =   16
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "20 Chars"
            Index           =   17
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "21 Chars"
            Index           =   18
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "22 Chars"
            Index           =   19
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "23 Chars"
            Index           =   20
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "24 Chars"
            Index           =   21
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "25 Chars"
            Index           =   22
         End
      End
      Begin VB.Menu mnuOptTypeCase 
         Caption         =   "&Type of case"
         Begin VB.Menu mnuOptCase 
            Caption         =   "Lowercase"
            Index           =   0
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Uppercase"
            Index           =   1
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Propercase"
            Index           =   2
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Mixed Case"
            Index           =   3
         End
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptSpecial 
         Caption         =   "&Special Characters"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmMain
'
' Description:   This is the main form that has multiple layers.
'                There are four parts:
'                    1.  The password generation.
'                    2.  The passphrase generation.
'                    3.  The passphrase display
'                    4.  The Working message.
'
'                The user can stop this process immediately by pressing the
'                STOP or EXIT button.
'
' IMPORTANT:  See Form_Load() routine.
'             Mouse wheel scroll does not work while in the VB IDE because
'             If you are attempting to debug the code, you could lock up to
'             the point that you would have to reboot.  I added some code
'             at the top of the Form_Load() event to prevent this.  Please
'             look at it.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 03-DEC-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module updated
' 20-DEC-2001  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete variables
' 27-Jun-2006  Kenneth Ives  kenaso@tx.rr.com
'              Replaced database with random access file.
' 05-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added mouse wheel scroll support.
'              Added ability to copy a single password from the grid to the
'              clipboard using Ctrl+C so it can be pasted into another
'              document.
'              Thanks to Herman McCrea for suggesting these additions.
' 18-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote ResizeColumns() routine.
'              Updated mnuFilePrint_Click() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "frmMain"
  Private Const MSG_1       As String = "No passphrase is offensive when protecting your data."
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' Reduce flicker while loading a control
  ' Lock the control to prevent redrawing
  '     Syntax:  LockWindowUpdate ctl_name.hWnd
  ' Unlock the control
  '     Syntax:  LockWindowUpdate 0&
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hwnd As Long) As Long
  
' ***************************************************************************
' Variables
' ***************************************************************************
  Private mblnFirstTime        As Boolean
  Private mblnUseNumbers       As Boolean   ' Used with passphrases
  Private mblnPlusNumbers      As Boolean   ' Used with passwords
  Private mblnSortPasswords    As Boolean
  Private mblnCreatePassword   As Boolean
  Private mblnPlusSpecialChars As Boolean   ' Used with passwords
  Private mstrFilename         As String
  Private mstrTypeCase         As String
  Private mlngDupes            As Long
  Private mlngColumns          As Long
  Private mlngPwdCount         As Long
  Private mlngTypeCase         As Long
  Private mlngWordLength       As Long
  Private mlngNbrOfWords       As Long
  Private mlngCharsToConv      As Long
  Private mintMinLength        As Integer
  Private mintMaxLength        As Integer
  Private mcolWords            As Collection  ' Collection of passwords/phrases
  Private mobjKeyEdit          As cKeyEdit
  

Private Sub cboNbrCount_Click()
    gstrNumberIndex = CStr(cboNbrCount.ListIndex)
End Sub

Private Sub cboSpecial_Click()

    Dim lngIndex  As Long
    Dim lngChoice As Long
          
    gblnSpecialWithNbrs = False       ' Assume FALSE
    mlngCharsToConv = 0               ' Reset number of letters to convert
    lngChoice = cboSpecial.ListIndex
    gstrSpecialIndex = CStr(lngChoice)
    
    ' Determine the number of characters
    ' to convert in a password
    Select Case lngChoice
         
           Case 0  ' Alphabetic Only
                mblnPlusNumbers = False
                mblnPlusSpecialChars = False
                cboNbrCount.Clear
                cboNbrCount.Enabled = False
                lblNbrCount.Enabled = False
                mnuOptSpecial.Enabled = False
              
           Case 1  ' Alphabetic with Numeric mix
                mblnPlusNumbers = True
                mblnPlusSpecialChars = False
                lblNbrCount.Enabled = True
                cboNbrCount.Enabled = True
                cboNbrCount.Clear
                mnuOptSpecial.Enabled = False
                
                ' load the combo box with numbers
                ' always one short so the first
                ' character is alphabetic
                If mlngWordLength > 0 Then
                    For lngIndex = 1 To (mlngWordLength - 1)
                        cboNbrCount.AddItem " " & lngIndex
                    Next lngIndex
                        
                    cboNbrCount.ListIndex = 0
                End If
                
                cboNbrCount_Click
            
            Case 2  ' Alphabetic with Special characters mix
                mblnPlusNumbers = False
                mblnPlusSpecialChars = True
                lblNbrCount.Enabled = True
                cboNbrCount.Enabled = True
                cboNbrCount.Clear
                mnuOptSpecial.Enabled = True
                
                ' load the combo box with numbers
                ' always one short so the first
                ' character is alphabetic
                If mlngWordLength > 0 Then
                    For lngIndex = 1 To (mlngWordLength - 1)
                        cboNbrCount.AddItem " " & lngIndex
                    Next lngIndex
                        
                    cboNbrCount.ListIndex = 0
                End If
                
                For lngIndex = 26 To 35
                    gastrChars(lngIndex) = vbNullString
                Next lngIndex
           
                cboNbrCount_Click
            
           Case 3  ' Alphabetic with Numeric and Special characters mix
                mblnPlusNumbers = True
                mblnPlusSpecialChars = True
                gblnSpecialWithNbrs = True
                lblNbrCount.Enabled = True
                cboNbrCount.Enabled = True
                cboNbrCount.Clear
                mnuOptSpecial.Enabled = True
                
                ' load the combo box with numbers
                ' always one short so the first
                ' character is alphabetic
                If mlngWordLength > 0 Then
                    For lngIndex = 1 To (mlngWordLength - 1)
                        cboNbrCount.AddItem " " & lngIndex
                    Next lngIndex
                        
                    cboNbrCount.ListIndex = 0
                End If
                
                For lngIndex = 26 To 35
                    gastrChars(lngIndex) = CStr(lngIndex - 26)
                Next lngIndex
    
                cboNbrCount_Click
    End Select
  
End Sub

Private Sub chkPassword_Click()

    mblnCreatePassword = CBool(chkPassword.Value)
    
    LoadSpecialCombo
    
    ' Set the password flag
    If mblnCreatePassword Then
        chkPassword.Caption = "Uncheck to create a Passphrase"
        chkSortData.Visible = True
        chkUseNumbers.Visible = False
        mblnSortPasswords = True
        lblDupes.Visible = True
        lblDupes.Caption = vbNullString
        lblWarningMsg.Caption = CStr(mlngWordLength) & " Characters" & vbNewLine & mstrTypeCase
        mnuOptLength.Enabled = True
        mnuOptCharLength_Click mlngWordLength - 3
        UpdateScreen 0
        
        If mblnPlusSpecialChars Then
            mnuOptSpecial.Enabled = True
        Else
            mnuOptSpecial.Enabled = False
        End If
        
    Else
        chkPassword.Caption = "Check to create Passwords"
        chkSortData.Visible = False
        chkUseNumbers.Visible = True
        mblnSortPasswords = False
        lblDupes.Visible = False
        lblWarningMsg.Caption = MSG_1
        mnuOptLength.Enabled = False
        mnuOptSpecial.Enabled = False
        UpdateScreen 1
    End If

End Sub

Private Sub chkSortData_Click()
    
    mblnSortPasswords = CBool(chkSortData.Value)
    
    If mcolWords Is Nothing Then
        Exit Sub
    End If
    
    If mcolWords.Count > 1 Then
        LoadPasswordGrid
    End If
    
End Sub

Private Sub chkUseNumbers_Click()

    mblnUseNumbers = CBool(chkUseNumbers.Value)
    
End Sub

' ***************************************************************************
' Routine:       cmdChoice_Click
'
' Description:   Perform the commands associated with a command button
'
' Parameters:    Index - which button was pressed
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Private Sub cmdChoice_Click(Index As Integer)
  
    Select Case Index
           Case 0:  ' Start and Stop
                If cmdChoice(0).Caption = "&Start" Then
                    
                    ' if START is displayed on the command button change
                    ' it to STOP and set the cancellation flag
                    gblnStopProcessing = False
                    gobjPrng.StopProcessing = gblnStopProcessing
                    DoEvents: DoEvents: DoEvents: DoEvents
                    
                    cmdChoice(0).Caption = "&Stop"
                    cmdChoice(1).Enabled = False
                    chkPassword.Enabled = False
                    chkSortData.Enabled = False
                    
                    txtPassphrase.Text = vbNullString   ' Empty output boxes
                    grdPasswords.Clear        ' Empty grid
                    
                    EmptyCollection mcolWords       ' make sure collection is empty
                    Set mcolWords = New Collection  ' Instantiate new collection
    
                    ' see if we are to create passwords
                    If mblnCreatePassword Then
                        
                        lblDupes.Caption = vbNullString
                        picWorking.Visible = True
                        DoEvents

                        ' Create passwords
                        BeginPasswords
                         
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            gobjPrng.StopProcessing = gblnStopProcessing
                            chkPassword.Enabled = True
                            chkSortData.Enabled = True
                            cmdChoice(0).Caption = "&Start"
                            cmdChoice(1).Enabled = True
                            picWorking.Visible = False    ' Hide working message
                            Exit Sub
                        End If
    
                    Else
                        ' Create a passphrase
                        ' Gather all the information from the screen
                        ' see how many words are to be used in the
                        ' passphrase
                        mlngNbrOfWords = Val(txtNbrOfWords.Text)
                        
                        If mlngNbrOfWords < 1 Or mlngNbrOfWords > MAX_WORDS Then
                            InfoMsg "Number of words used in a phrase is limited to 1-" & CStr(MAX_WORDS)
                            chkPassword.Enabled = True
                            chkSortData.Enabled = True
                            cmdChoice(0).Caption = "&Start"
                            picWorking.Visible = False    ' Hide the working message
                            Exit Sub
                        End If
                        
                        picWorking.Visible = True
                        CreatePassphrase mlngTypeCase, mlngNbrOfWords, mblnUseNumbers, mcolWords
                        
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            gobjPrng.StopProcessing = gblnStopProcessing
                            chkPassword.Enabled = True
                            chkSortData.Enabled = True
                            cmdChoice(0).Caption = "&Start"
                            cmdChoice(1).Enabled = True
                            picWorking.Visible = False    ' Hide the working message
                            Exit Sub
                        End If
    
                        ' Display the passphrase
                        UpdateScreen 2
                    End If
                        
                    chkPassword.Enabled = True
                    chkSortData.Enabled = True
                    cmdChoice(0).Caption = "&Start"
                    cmdChoice(1).Enabled = True
                    picWorking.Visible = False    ' Hide the working message
                    DoEvents
                    
                ElseIf cmdChoice(0).Caption = "&Stop" Then
                    ' if STOP is displayed on the command button change
                    ' it to START and reset the cancellation switch
                    DoEvents
                    gblnStopProcessing = True
                    gobjPrng.StopProcessing = gblnStopProcessing
                    chkPassword.Enabled = True
                    chkSortData.Enabled = True
                    cmdChoice(0).Caption = "&Start"
                    cmdChoice(1).Enabled = True
                    picWorking.Visible = False    ' Hide the working message
                    EmptyCollection mcolWords     ' Empty collection
                    DoEvents
                End If
           
           Case 1:  ' Exit application
                DoEvents
                gblnStopProcessing = True
                gobjPrng.StopProcessing = gblnStopProcessing
                
                DoEvents
                EmptyCollection mcolWords     ' Empty collection
                TerminateProgram
    End Select
    
End Sub

' ***************************************************************************
' Routine:       UpdateScreen
'
' Description:   Prepare a screen for a display update.  this can be either
'                the initial input or output.
'
' Parameters:    Index - which screen to update
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Private Sub UpdateScreen(Index As Integer)
   
    Dim lngIdx As Long
    
    ' Update the display
    Select Case Index
           Case 0: ' Prepare the password frame
                fraPhrase.Visible = False
                fraPasswords.Visible = True
                grdPasswords.Clear
           
           Case 1: ' Prepare the passphrase frame
                fraPasswords.Visible = False
                fraPhrase.Visible = True
                txtPassphrase.Text = vbNullString
           
           Case 2: ' display the passphrase
                txtPassphrase.Text = vbNullString
                If mcolWords.Count > 0 Then
                    For lngIdx = 1 To mcolWords.Count
                        txtPassphrase.Text = txtPassphrase.Text & mcolWords.Item(lngIdx) & " "
                    Next lngIdx
                End If
    End Select
  
End Sub

Private Sub Form_Load()
  
    ' See if we are in the developement environment
'    If gblnIDE_Environment Then
'        ' Deactivate the wheel scrolling capability while in the
'        ' VB IDE.  If you do not allow this to happen, you will
'        ' have great difficulty stepping thru the code.
'        WheelUnHook frmMain.hwnd
'    Else
        WheelHook frmMain.hwnd     ' Activate mouse wheel
'    End If
    
    Set mobjKeyEdit = New cKeyEdit   ' Instantiate class object
    
    ' Initialize combo boxes and variables
    mintMinLength = 0
    mintMaxLength = 0
    mlngColumns = 0
    mlngCharsToConv = 0
    mstrFilename = vbNullString
    mblnFirstTime = True
    mblnUseNumbers = False
    mblnCreatePassword = False
    
    FillComboBoxes   ' load combo boxes
    
    mnuOptCharLength_Click Val(gstrPwdLenIndex)   ' default is eight characters
    mnuOptCase_Click Val(gstrCaseIndex)           ' default is lowercase display
    
    chkPassword.Value = vbUnchecked
    chkPassword_Click    ' show passphrase window
    
    ' set up the form to be displayed
    With frmMain
        .Caption = gstrVersion
        .lblWarningMsg.Caption = MSG_1
        .mnuOptLength.Enabled = False
        .mnuOptSpecial.Enabled = False
        .picWorking.Visible = False
        .txtQuantity.Text = "1"
        .txtNbrOfWords.Text = "1"
        .grdPasswords.Rows = 0
        .grdPasswords.Cols = 0
        .grdPasswords.AllowUserResizing = flexResizeColumns
        .chkPassword.Value = vbUnchecked
        .chkUseNumbers.Value = vbUnchecked
        .chkUseNumbers.Visible = True
        .chkSortData.Value = vbChecked
        .chkSortData.Visible = False
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless
        .Refresh
    End With

    DisableX frmMain
    mobjKeyEdit.CenterCaption frmMain

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    WheelUnHook frmMain.hwnd    ' Deactivate mouse wheel
    Clipboard.Clear             ' Clear clipboard of captured data
    Set mobjKeyEdit = Nothing   ' Free class object from memory
     
End Sub

Private Sub Form_Resize()

    If frmMain.WindowState = vbMinimized Then
        frmMain.Caption = "Passphrase"
    Else
        frmMain.Caption = gstrVersion
        mobjKeyEdit.CenterCaption frmMain
    End If
    
End Sub

Private Sub grdPasswords_EnterCell()
  
    ' When first entering cell, change the background color to yellow
    grdPasswords.CellBackColor = vbYellow
  
End Sub

Private Sub grdPasswords_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim CtrlDown   As Integer
    Dim PressedKey As Integer
    
    ' Initialize  variables
    CtrlDown = (Shift And vbCtrlMask) > 0    ' Define control key
    PressedKey = Asc(UCase$(Chr$(KeyCode)))  ' Convert to uppercase
      
    ' Ctrl + C was pressed
    If CtrlDown And PressedKey = vbKeyC Then
        
        Clipboard.Clear                      ' clear the clipboard
        Clipboard.SetText grdPasswords.Text  ' load clipboard with highlighted text

    End If

End Sub

Private Sub grdPasswords_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    
        Case vbKeyReturn, vbKeyTab
            'move to next cell.
            With grdPasswords
                If .Col + 1 <= .Cols - 1 Then
                    .Col = .Col + 1
                Else
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                        .Col = 0
                    Else
                        .Row = 1
                        .Col = 0
                    End If
                End If
            End With
            
        Case vbKeyBack
            With grdPasswords
                'remove the last character, if any.
                If Len(.Text) Then
                    .Text = Left$(.Text, Len(.Text) - 1)
                End If
            End With
            
        Case Is < 32
        
        Case Else
            With grdPasswords
                .Text = .Text & Chr$(KeyAscii)
            End With
            
    End Select

End Sub

Private Sub grdPasswords_LeaveCell()

    ' When leaving the cell, change the background color to white
    grdPasswords.CellBackColor = vbWhite
  
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub mnuAbout_Click()

    Dim strMsg As String
    
    ' Initialize variables
    strMsg = StrConv(App.EXEName & ".exe", vbProperCase) & vbNewLine
    strMsg = strMsg & "Written by " & App.CompanyName & vbNewLine & vbNewLine
    strMsg = strMsg & "Password file contains several thousand English    " & vbNewLine
    strMsg = strMsg & "words that are 3-10 characters in length."
    
    InfoMsg strMsg
   
End Sub

Private Sub mnuOptCase_Click(Index As Integer)
  
    Dim lngIndex As Long
    
    ' Set the appropriate check mark
    For lngIndex = 0 To 3
        If lngIndex = Index Then
            mnuOptCase(lngIndex).Checked = True
            mlngTypeCase = lngIndex
            mstrTypeCase = mnuOptCase(lngIndex).Caption
        Else
            ' remove all other check marks
            mnuOptCase(lngIndex).Checked = False
        End If
    Next lngIndex
    
    lblWarningMsg.Caption = CStr(mlngWordLength) & " Characters" & vbNewLine & mstrTypeCase
    gstrCaseIndex = CStr(mlngTypeCase)
    
End Sub

Private Sub mnuOptCharLength_Click(Index As Integer)
  
    Dim lngIndex As Long
    
    ' Set the appropriate check mark
    For lngIndex = 0 To 22
        If lngIndex = Index Then
            mnuOptCharLength(lngIndex).Checked = True
            mlngWordLength = lngIndex + 3      ' calc the word length
            
            If mblnPlusNumbers Or _
               mblnPlusSpecialChars Then
                
                cboSpecial_Click   ' refill letters to convert box
            End If
        Else
            ' remove all other check marks
            mnuOptCharLength(lngIndex).Checked = False
        End If
    Next lngIndex
    
    lblWarningMsg.Caption = CStr(mlngWordLength) & " Characters" & vbNewLine & mstrTypeCase
    gstrPwdLenIndex = CStr(mlngWordLength - 3)
    
End Sub

Private Sub mnuFileExit_Click()
    TerminateProgram     ' Shutdown this application
End Sub

' ***************************************************************************
' Routine:       mnuOptSpecial_Click
'
' Description:   display an input box with a list of special characters to
'                be omitted
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Private Sub mnuOptSpecial_Click()
    frmSpecial.ShowForm
End Sub

Private Sub mnuFileOpen_Click()

    On Error GoTo ErrHandler
    cmDialog.CancelError = True     ' Set Cancel to True.
  
    ' Setup and display the "FILE OPEN" dialog box
    With cmDialog
         .Flags = cdlOFNHideReadOnly Or _
                  cdlOFNExplorer Or _
                  cdlOFNLongNames Or _
                  cdlOFNFileMustExist
         .FileName = vbNullString
         
         ' Set filters
         .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         .FilterIndex = 1   ' Specify default filter
         .ShowOpen          ' Display the Open dialog box
    End With
    
    ' Capture name of selected file
    mstrFilename = cmDialog.FileName
    
    If Len(Trim$(mstrFilename)) = 0 Then
        Exit Sub
    End If
    
    ' Display this file using the default text editor
    DisplayFile mstrFilename, frmMain
    Exit Sub
    
ErrHandler:
    ' User pressed the Cancel button
    Exit Sub

End Sub

Private Sub mnuFilePrint_Click()

    Dim lngCol        As Long
    Dim lngMax        As Long
    Dim lngIndex      As Long
    Dim strFmt        As String
    Dim strTmpLine    As String
    Dim strTextData   As String
    Dim strTypeCase   As String
    Dim strTitleLine1 As String
    Dim strTitleLine2 As String
    Dim cPrint        As cPrint
    
    On Error GoTo FilePrint_CleanUp
    
    strTitleLine1 = vbNullString
    strTextData = vbNullString
    cmDialog.CancelError = True   ' Set CancelError to True.
    
    ' Capture and test to see if we have any data
    If mblnCreatePassword Then
    
        If mcolWords Is Nothing Then
            GoTo FilePrint_CleanUp
        
        ElseIf mcolWords.Count < 1 Then
            GoTo FilePrint_CleanUp
        
        Else
            Select Case mlngTypeCase
                   Case 0: strTypeCase = "All lowercase letters"
                   Case 1: strTypeCase = "All uppercase letters"
                   Case 2: strTypeCase = "First character is uppercase"
                   Case 3: strTypeCase = "Mixed case letters"
            End Select
            
            strTitleLine1 = Format$(mlngPwdCount, "#,##0") & " Passwords " & _
                            Format$(mlngWordLength, "#,##0") & " characters long"
            strTitleLine2 = strTypeCase & _
                            IIf(mlngCharsToConv = 0, "", " with " & CStr(mlngCharsToConv) & " characters converted")

            strFmt = String$(Len(mcolWords.Item(1)), "@")
            lngMax = Int(78 / (Len(strFmt) + 2))
            strTmpLine = vbNullString
            lngCol = 0
            
            For lngIndex = 1 To mcolWords.Count
            
                strTmpLine = strTmpLine & Format$(mcolWords.Item(lngIndex), strFmt) & Space$(2)
                lngCol = lngCol + 1
                
                If lngMax = lngCol Or _
                   lngIndex = mcolWords.Count Then
                    
                    strTextData = strTextData & strTmpLine
                    lngCol = 0
                    strTmpLine = vbNullString
                End If
                
            Next lngIndex
            
        End If
    Else
        If Len(Trim$(txtPassphrase.Text)) = 0 Then
            GoTo FilePrint_CleanUp
        Else
            strTitleLine1 = "Passphrase"
            strTitleLine2 = "Word count:  " & Format$(mcolWords.Count, "#,##0")
            strTextData = txtPassphrase.Text
        End If
    End If
      
    ' See if we have data to print
    If Len(Trim$(strTextData)) = 0 Then
        Exit Sub
    End If
    
    ' Display the "Print" dialog box
    With cmDialog
         ' default to print all pages, no saving to a file,
         ' and no selective printing
         .Flags = cdlPDAllPages Or _
                  cdlPDHidePrintToFile Or _
                  cdlPDNoSelection
         .ShowPrinter   ' Display the Print dialog box
    End With
    
    ' change the curser to an hourglass
    Screen.MousePointer = vbHourglass
    
    ' Print the data
    strTextData = Trim$(strTextData)
    
    Set cPrint = New cPrint
    cPrint.PrintText strTextData, strTitleLine1, strTitleLine2
        
FilePrint_CleanUp:
    ' Normal exit. Jump to here if user presses
    ' Cancel button or no data to process.
    Set cPrint = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuFileSaveAs_Click()

    Dim hFile       As Long
    Dim lngCol      As Long
    Dim lngMax      As Long
    Dim lngIndex    As Long
    Dim lngPosition As Long
    Dim strFmt      As String
    Dim strCase     As String
    Dim strName     As String
    Dim strTmpLine  As String
    Dim strTextData As String
    
    On Error GoTo FileSaveAs_CleanUp
    
    strTextData = vbNullString
    strTmpLine = vbNullString
    cmDialog.CancelError = True  ' Set CancelError is True
    
    ' Capture and test to see if we have any data
    If mblnCreatePassword Then
    
        If mcolWords Is Nothing Then
            Exit Sub
        
        ElseIf mcolWords.Count < 1 Then
            Exit Sub
        
        Else
            strFmt = String$(Len(mcolWords.Item(1)), "@")
            lngMax = Int(80 / (Len(strFmt) + 2))
            lngCol = 0
            
            For lngIndex = 1 To mcolWords.Count
            
                strTmpLine = strTmpLine & Format$(mcolWords.Item(lngIndex), strFmt) & Space$(2)
                lngCol = lngCol + 1
                
                If lngMax = lngCol Or _
                   lngIndex = mcolWords.Count Then
                    
                    strTextData = strTextData & Trim$(strTmpLine) & vbNewLine
                    lngCol = 0
                    strTmpLine = vbNullString
                End If
                
            Next lngIndex
        End If
    Else
        If Len(Trim$(txtPassphrase.Text)) = 0 Then
            Exit Sub
        Else
            strTextData = txtPassphrase.Text
        End If
    End If
    
    ' See if we have data to save
    If Len(Trim$(strTextData)) = 0 Then
        Exit Sub
    End If
    
    ' Display the "FILE SAVE AS" dialog box
    With cmDialog
         ' Set flags
         .Flags = cdlOFNExplorer Or _
                  cdlOFNLongNames Or _
                  cdlOFNHideReadOnly Or _
                  cdlOFNOverwritePrompt
         
         .FileName = vbNullString   ' empty filename selection
         
         ' Set filters
         .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         .FilterIndex = 1   ' Specify default filter
         .ShowOpen          ' Display File Open dialog box
         
         mstrFilename = cmDialog.FileName    ' capture filename
    End With
    
    ' see if a file name was entered or selected
    If Len(Trim$(mstrFilename)) = 0 Then
        Exit Sub
    Else
        ' save just the filename
        lngPosition = InStrRev(mstrFilename, "\", Len(mstrFilename))
        strName = Mid$(mstrFilename, lngPosition + 1)
    End If
    
    Select Case mlngTypeCase
           Case 0: strCase = " (Lowercase characters)"
           Case 1: strCase = " (Uppercase characters)"
           Case 2: strCase = " (Propercase characters)"
           Case 3: strCase = " (Mixed case characters)"
    End Select

    ' save the file with a max line length of 80
    hFile = FreeFile
    Open mstrFilename For Output As #hFile
    Print #hFile, " "
    Print #hFile, "Filename:  " & strName
    Print #hFile, "Created:   " & Format$(Now(), "dddd  d mmmm yyyy  h:mm ampm")
    Print #hFile, " "
    
    If mblnCreatePassword Then
        strTmpLine = "Password length:  " & Format$(mlngWordLength, "#,##0") & strCase
    
        If mlngCharsToConv > 0 Then
            Print #hFile, strTmpLine & vbNewLine & Space$(18) & CStr(mlngCharsToConv) & _
                          " letters converted to numbers and/or symbols"
        Else
            Print #hFile, strTmpLine
        End If
    
        Print #hFile, "Password count:   " & Format$(mlngPwdCount - mlngDupes, "#,##0") & " (" & _
                      Format$(mlngDupes, "#,##0") & " duplicates were removed)"
    Else
        strTmpLine = "Passphrase count:  " & Format$(mlngNbrOfWords, "#,##0") & strCase
        Print #hFile, strTmpLine
    End If
    
    Print #hFile, " "
    Print #hFile, String$(80, 45)
    Print #hFile, " "
    Print #hFile, strTextData
    Print #hFile, String$(80, 45)
    Print #hFile, " "
            
FileSaveAs_CleanUp:
    ' User pressed Cancel button
    Close #hFile   ' close output file
    
End Sub

' ***************************************************************************
' Routine:       FillComboBoxes
'
' Description:   Fill all the combo boxes with data
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Private Sub FillComboBoxes()
    
    ' Load all the combo boxes for read only access
    With frmMain
        .cboNbrCount.Clear  ' empty combobox
        LoadSpecialCombo    ' load combobox
    End With
    
End Sub

Private Sub LoadSpecialCombo()

    ' Special character options
    With cboSpecial
        .Clear
        .AddItem " Alphabetic Only"
        .AddItem " Numeric Mix"
        .AddItem " Special Char Mix"
        .AddItem " Numeric and Special"
        .ListIndex = Val(gstrSpecialIndex)
    End With

End Sub

' ***************************************************************************
' Routine:       BeginPasswords
'
' Description:   Verify the input data to start building passwords
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Private Sub BeginPasswords()

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    ' Evaluate the count
    If Len(Trim$(txtQuantity.Text)) = 0 Or _
       Val(Trim$(txtQuantity.Text)) = 0 Then
           txtQuantity.Text = "0"
           Exit Sub
    End If
    
    ' read the combo box
    mlngPwdCount = Val(Trim$(txtQuantity.Text))
    mlngCharsToConv = Val(Trim$(cboNbrCount.Text))
    gstrNumberIndex = CStr(cboNbrCount.ListIndex)
    
    ' Create the passwords
    CreatePasswords mlngPwdCount, _
                    mlngWordLength, _
                    mlngCharsToConv, _
                    mlngTypeCase, _
                    mblnPlusNumbers, _
                    mblnPlusSpecialChars, _
                    mcolWords
                              
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    If mcolWords.Count > 0 Then
        LoadPasswordGrid  ' Format the output display
    End If
    
End Sub

' ***************************************************************************
' Routine:       LoadPasswordGrid
'
' Description:   Formats the output display of the passwords for the screen
'
' Parameters:    lngCount - number of passwords to format
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 30-NOV-2000  Kenneth Ives  kenaso@tx.rr.com
'              Modified the output display
' 03-DEC-2000  Kenneth Ives  kenaso@tx.rr.com
'              Modified the output display and used the MS FlexGrid control
' 07-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated collection tests
' ***************************************************************************
Private Sub LoadPasswordGrid()

    Dim lngIndex       As Long
    Dim lngPwdCount    As Long
    Dim lngRowCount    As Long
    Dim lngWordCounter As Long
    Dim strOutput      As String
    Dim astrPWord()    As String
    
    Const ROUTINE_NAME As String = "LoadPasswordGrid"

    On Error GoTo LoadPasswordGrid_Error

    ' if the stop button was pressed then leave
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    ' See if collection has been instatiated
    If mcolWords Is Nothing Then
        Exit Sub
    End If
    
    ' Does collection have any data
    If mcolWords.Count < 1 Then
        Exit Sub
    End If
    
    lngWordCounter = 0
    lngRowCount = 0
    mlngDupes = 0
    lngPwdCount = mcolWords.Count
    
    Erase astrPWord()
    ReDim astrPWord(lngPwdCount)
             
    For lngIndex = 1 To lngPwdCount
        astrPWord(lngIndex - 1) = mcolWords.Item(lngIndex)
    Next lngIndex
    
    ' if the stop button was pressed then leave
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
        
    If mcolWords.Count > 1 Then
        
        With gobjPrng
            If mblnSortPasswords Then
                If mcolWords.Count = 2 Then
                    ' See if first password is out of sequence
                    If StrComp(astrPWord(0), astrPWord(1), vbTextCompare) = 1 Then
                        .SwapData astrPWord(0), astrPWord(1)
                    End If
                Else
                    ' Sort the data in Ascending order
                    .CombSort astrPWord()                 ' Sort passwords in Ascending order
                    .RemoveDupes astrPWord(), mlngDupes   ' Remove any duplicate values
                End If
            Else
                If mcolWords.Count = 2 Then
                    ' See if first password is in sequence
                    If StrComp(astrPWord(0), astrPWord(1), vbTextCompare) = -1 Then
                        .SwapData astrPWord(0), astrPWord(1)
                    End If
                Else
                    .CombSort astrPWord()                 ' Sort passwords in Ascending order
                    .RemoveDupes astrPWord(), mlngDupes   ' Remove any duplicate values
                    .ReshuffleData astrPWord()            ' Reshuffle passwords
                End If
            End If
        End With
                
    End If
    
    lngPwdCount = UBound(astrPWord)  ' Capture number of valid passwords
        
    If mlngDupes < 1 Then
        mlngDupes = 0
    End If
    
    EmptyCollection mcolWords       ' Empty collection
    Set mcolWords = New Collection  ' Instantiate new collection
    
    ' Load with new password list
    For lngIndex = 0 To lngPwdCount - 1
        mcolWords.Add astrPWord(lngIndex)
    Next lngIndex
        
    ' set up number of grid columns.  This was tedious.  There sometimes can be
    ' a difference in display between IDE and the compiled version.  Be sure to
    ' thoroughly check both.  Been there and it is not fun.  :-)
    Select Case mlngWordLength
           Case 3:        mlngColumns = 16
           Case 4:        mlngColumns = 13
           Case 5:        mlngColumns = 10
           Case 6:        mlngColumns = 9
           Case 7:        mlngColumns = 8
           Case 8:        mlngColumns = 7
           Case 9:        mlngColumns = 6
           Case 10, 11:   mlngColumns = 5
           Case 12 To 14: mlngColumns = 4
           Case 15 To 20: mlngColumns = 3
           Case Else:     mlngColumns = 2
    End Select
    
    ' Temporarily lock the grid control while loading.
    ' This will speed things up and reduce the amount of
    ' flicker
    LockWindowUpdate frmMain.hwnd
    
    ' Prepare the grid
    With grdPasswords
         .Clear               ' remove any previous data
         .Rows = 1            ' number of current rows
         .Cols = mlngColumns  ' number of columns
    End With
         
    strOutput = vbNullString
    
    ' loop thru and build the password display output
    For lngIndex = 0 To lngPwdCount - 1
               
        lngWordCounter = lngWordCounter + 1
        
        If lngWordCounter < mlngColumns Then
            ' append password to output string and then
            ' append a TAB as a delimiter
            strOutput = strOutput & astrPWord(lngIndex) & vbTab
        Else
            strOutput = strOutput & astrPWord(lngIndex)   ' Append one more password
            grdPasswords.AddItem strOutput, lngRowCount   ' Add data to grid
            strOutput = vbNullString                                ' empty output string
            lngWordCounter = 0                            ' reset word counter
            lngRowCount = lngRowCount + 1                 ' update the row position
            grdPasswords.Rows = lngRowCount               ' append an empty row
        End If
        
        ' if the stop button was pressed then leave
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
        
    Next lngIndex
          
    ' if the stop button was pressed then leave
    DoEvents
    If gblnStopProcessing Then
        GoTo LoadPasswordGrid_CleanUp
    End If
        
    If Len(strOutput) > 0 Then
        grdPasswords.AddItem strOutput, lngRowCount
        strOutput = vbNullString
    End If
        
    DoEvents
    lblDupes.Caption = Format$(mlngDupes, "#,##0") & " duplicates removed"
    ResizeColumns
    
LoadPasswordGrid_CleanUp:
    DoEvents
    ' unlock grid control after loading the data
    LockWindowUpdate 0&
    grdPasswords.SetFocus
    Erase astrPWord()
    On Error GoTo 0
    Exit Sub

LoadPasswordGrid_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume LoadPasswordGrid_CleanUp
    
End Sub

' ***************************************************************************
'  Routine:         ResizeColumns
'
'  Description:     Adjust grid column width
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Sub ResizeColumns()

    Dim intCol  As Integer
    Dim strText As String
    Dim objFont As Font
    
    With frmMain.grdPasswords
        .Row = 0   ' Go to top row, first column
        .Col = 0
        Set objFont = .Parent.Font
        Set .Parent.Font = .Font
    
        ' Adjust column width
        For intCol = 0 To .Cols - 1
            strText = .Clip
            .ColWidth(intCol) = .Parent.TextWidth(strText) + 120
        Next intCol
        
        Set .Parent.Font = objFont
        
        .Row = 0   ' Go to top row, first column
        .Col = 0
    End With
    
    Set objFont = Nothing   ' Free object from memory
    
End Sub

Private Sub txtGridEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Process possible key combinations
    mobjKeyEdit.TextBoxKeyDown txtGridEdit, KeyCode, Shift

End Sub

Private Sub txtNbrOfWords_Change()

    ' Prevent user from pasting a non-numeric value
    ' into this textbox
    If Not IsNumeric(txtNbrOfWords.Text) Then
        txtNbrOfWords.Text = vbNullString
    End If

End Sub

Private Sub txtNbrOfWords_GotFocus()

    ' Highlight everything in the text box
    mobjKeyEdit.TextBoxFocus txtNbrOfWords
    
End Sub

Private Sub txtNbrOfWords_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Process possible key combinations
    mobjKeyEdit.TextBoxKeyDown txtNbrOfWords, KeyCode, Shift

End Sub

Private Sub txtNbrOfWords_KeyPress(KeyAscii As Integer)

    ' Allow only numbers and backspace to be entered
    mobjKeyEdit.ProcessNumericOnly KeyAscii
  
End Sub

Private Sub txtNbrOfWords_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySeparator Or _
       KeyCode = vbEnter Or _
       KeyCode = vbKeyTab Then
        
        cmdChoice_Click 0
    End If
    
End Sub

Private Sub txtNbrOfWords_LostFocus()

    ' Test minimum and maximum value of word count
    If Val(Trim$(txtNbrOfWords.Text)) > MAX_WORDS Then
        InfoMsg "Maximum value allowed is " & CStr(MAX_WORDS)
    End If
  
End Sub

Private Sub txtPassphrase_GotFocus()

    ' Highlight everything in the text box
    mobjKeyEdit.TextBoxFocus txtPassphrase
  
End Sub

Private Sub txtPassphrase_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Process possible key combinations
    mobjKeyEdit.TextBoxKeyDown txtPassphrase, KeyCode, Shift

End Sub

Private Sub txtQuantity_Change()

    ' Prevent user from pasting a non-numeric value
    ' into this textbox
    If Not IsNumeric(txtQuantity.Text) Then
        txtQuantity.Text = vbNullString
    End If

End Sub

Private Sub txtQuantity_GotFocus()

    ' Highlight everything in the text box
    mobjKeyEdit.TextBoxFocus txtQuantity
    
End Sub

Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Process possible key combinations
    mobjKeyEdit.TextBoxKeyDown txtQuantity, KeyCode, Shift

End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)

    ' Allow only numbers and backspace to be entered
    mobjKeyEdit.ProcessNumericOnly KeyAscii
  
End Sub

Private Sub txtQuantity_LostFocus()
  
    ' Minimum value of one and maximum value of 5000
    
    If Val(Trim$(txtQuantity.Text)) < 1 Then
        InfoMsg "Minimum value allowed is 1"
        Exit Sub
    End If
  
    If Val(Trim$(txtQuantity.Text)) > 5000 Then
        InfoMsg "Maximum value allowed is 5000"
    End If
  
End Sub

' Here you can add scrolling support to controls that don't normally respond
Public Sub MouseWheel(ByVal lngRotation As Long, _
                      ByVal lngXpos As Long, _
                      ByVal lngYpos As Long)

    Dim ctl As Control

    For Each ctl In Me.Controls

        If TypeOf ctl Is MSFlexGrid Then
            If IsOver(ctl.hwnd, lngXpos, lngYpos) Then
                FlexGridScroll ctl, lngRotation
            End If
        End If

'        If TypeOf ctl Is MSHFlexGrid Then
'            If IsOver(ctl.hWnd, lngXpos, lngYpos) Then
'                HorFlexGridScroll ctl, lngRotation
'            End If
'        End If

'        If TypeOf ctl Is DataGrid Then
'            If IsOver(ctl.hWnd, lngXpos, lngYpos) Then
'                DataGridScroll ctl, lngRotation
'            End If
'        End If

    Next ctl

End Sub




