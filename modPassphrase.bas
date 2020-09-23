Attribute VB_Name = "modPassphrase"
' ***************************************************************************
' Module:        modPassphrase
'
' Description:   This module creates the passwords and passphrases.
'
' ---------------------------------------------------------------------------
' Passwords, for the most part, follow a pattern and are case sensitive.
' See if your personal password is within these guidelines.
'
' This information can be obtained via personal observation, viewing
' data posted on the internet or through casual conversation with the
' user or co-worker.
'
' 1.  The weakest are:
'     - Telephone numbers (home, spouse, office, etc.)
'     - Birthdays (spouse, children, siblings, etc.)
'     - Combination of birth dates (your month, wife's day, child's year)
'     - Street address (home, work, ball park, lover, etc.)
'     - Five or less characters
'
' 2.  Very common (Weak to moderate security)
'     - first name (yours, spouse, children, pets, siblings)
'     - last name (yours, wife's maiden, mother's maiden, etc.)
'     - a series of letters or numbers, usually repeated
'       (i.e. 9876543210 or 555555 or aaaaaaaa)
'     - a row of keyboard characters starting from the left or right side
'       (i.e. asdfghjkl or zxcvbnm or qwertyui)
'     - a valid word from a dictionary
'
' 3.  Very common (Weak to moderate security)
'     Same as number 2 above with 1-4 numbers appended.
'     (i.e.  judy1234, ken2011, etc.)
'
' 4.  Strong security
'     - Excellent security if eight or more characters and all of the numbers
'       do not append the password.
'     - A mixture of letters and/or numbers.  The letters can be a mixture
'       of upper and lower case.
'     - A mixture of any of the printable keyboard characters with an ASCII
'       decimal value of 32-126.
'     - Use phrases sometimes mixed with unrelated data.  Quite often these
'       systems are designed to use 5-50 character input.  Below are a
'       coouple of example phrases:
'
'              "Soylent green is people x-4478"    <- 30 characters
'              "hornbill temp classics"            <- 3 random words
'
' Note:  There are some systems that use the ESCape key and one or more
'        function keys but they are too few and definitely not within the
'        realm of this application.  Usually these are high security
'        computer systems with a very limited personnel access. Tampering
'        with these systems will result in several years within a public
'        funded facility.  A positive note is you will not have to worry
'        about what fashions are in that season.
'
' ---------------------------------------------------------------------------
' Suggestions to slow down a potential security breach:
'
'     - Require a minimum password length (eight characters or more) with
'       at least two non-alphabetic characters not adjacent to each other.
'       Ex: 0-9, !, @, #, $, %, ^, &, *, (, )
'
'     - Allow ten tries and then disable the account after forcing a reboot.
'       User must contact security to reactivate their account.
'
'     - Force a reboot after every three attempts.
'
'     - Force the user to change their password every 30-45 days.
'
'     - Once a password has been changed and accepted, a user must wait 24
'       hours before being allowed to change their password again.
'
'     - Maintain a database of previous passwords for this user and not
'       allow any of their previous 10-25 passwords be used again.  Store
'       passwords in a hashed format.
'
'     - A written reprimand and\or unpaid time off if a user writes down
'       their password and leaves it in the open for others to view or
'       they share their password with others.  Also consider locking their
'       account until they have attended retraining concerning password
'       security and personal counciling by their immediate supervisor.
'
' These are just suggestions in ways to get you started in thinking in terms
' of data security.
'
' ---------------------------------------------------------------------------
' If you are curious about how many possible passwords could be generated,
' use these calculations:
'
'  chars_used = 62 = (10 + 26 + 26) = (0-9, A-Z, a-z)
'
'      (chars_used ^ word_length)                            possible
'        plus prev_possibles                               combinations
'        (62^2) + 62                                              3,906
'        (62^3) + 3,906                                         242,234
'        (62^4) + 242,234                                    15,018,570
'        (62^5) + 15,018,570                                916,132,832
'        (62^6) + 916,132,832                            57,731,386,986
'        (62^7) + 57,731,386,986                      3,579,345,993,194
'        (62^8) + 3,579,345,993,194                 221,919,451,578,090
'        (62^9) + 221,919,451,578,090            13,759,005,997,841,642
'        (62^10) + 13,759,005,997,841,642       853,058,371,866,181,866
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Created module
' 04-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected pointer array sizing and loading in CreatePasswords()
' 17-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated CreatePassphrase() and CreatePasswords() routines
' 14-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected pointer position for character changes if password
'              three characters in length in CreatePassphrase() routine
' ***************************************************************************
  
' ---------------------------------------------------------------------------
' Components:
'     Microsoft Common Dialog Control 6.0 (Sp3)
'     Microsoft FlexGrid Control 6.0 (Sp3)
' ---------------------------------------------------------------------------
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "modPassPhrase"
  Private Const MAXLONG     As Long = &H7FFFFFFF   ' 2147483647
      
' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type PWORD_DATA
      ID   As Long
      Word As String * 10
  End Type
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  
' ***************************************************************************
' Routine:       CreatePassphrase
'
' Description:   Build a passphrase using varying length words.  Never start
'                a passphrase with a number.  Most security systems will
'                only accept an alpha character a the first value.
'
' Parameters:    strWordFile - Name of the word file to be read
'                lngNbrOfWords - Number of words to extract from the table
'                blnUseNumbers - Replace one of the words in the phrase with
'                    a string of numbers varying 3-10 digits long if more
'                    than two words are to be created.
'                colWords - Collection of words that make up a passphrase
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 17-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to insert a string of numbers if creating
'              a phrase with more than two words
' ***************************************************************************
Public Sub CreatePassphrase(ByVal lngTypeCase As Long, _
                            ByVal lngNbrOfWords As Long, _
                            ByVal blnUseNumbers As Boolean, _
                            ByRef colWords As Collection)
  
    Dim hFile        As Long
    Dim lngIndex     As Long
    Dim lngMaxSize   As Long
    Dim lngPosition  As Long
    Dim lngRecordCnt As Long
    Dim alngRecID()  As Long
    Dim strWord      As String
    Dim strNumber    As String
    Dim typPWord     As PWORD_DATA
  
    Const ROUTINE_NAME As String = "CreatePassphrase"

    On Error GoTo CreatePassphrase_Error

    ' Calculate number of records in random access file
    lngRecordCnt = FileLen(gstrPwdFile) \ Len(typPWord)
    
    ' if file is empty, display a message and leave
    If lngRecordCnt < 1 Then
        InfoMsg "Passphrase file is empty.  [CreatePassphrase]" & _
                vbNewLine & gstrPwdFile
        GoTo CreatePassphrase_CleanUp
    End If
           
    ' Reset boolean flag if less than
    ' three words are to be created
    If lngNbrOfWords < 3 Then
        blnUseNumbers = False
    End If
        
    lngMaxSize = (MAX_WORDS * 2)  ' Calc record ID array size
    Erase alngRecID()             ' Always start with empty arrays
    ReDim alngRecID(lngMaxSize)   ' Size record ID array
    
    With gobjPrng
    
        ' Load record ID array with twice as many
        ' words allowed for a phrase in case of
        ' duplicate values generated
        For lngIndex = 0 To (lngMaxSize - 1)
            alngRecID(lngIndex) = .GetRndValue(1, lngRecordCnt)
        Next lngIndex
        
        .CombSort alngRecID()       ' Sort record ID array
        .RemoveDupes alngRecID()    ' Remove any duplicates
        .ReshuffleData alngRecID()  ' Mix record ID array
        
        If lngNbrOfWords > 2 Then
            
            ' User opted to replace one of
            ' the words with a numeric value
            If blnUseNumbers Then
                strNumber = CStr(.GetRndValue(10000, MAXLONG))      ' Create random numeric value
                lngPosition = .GetRndValue(3, 10)                   ' Create length of numeric string
                strNumber = Left$(strNumber, lngPosition)           ' Resize numeric string
                lngPosition = .GetRndValue(1, (lngNbrOfWords - 1))  ' Determine which word to replace
            End If
        End If
    End With
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CreatePassphrase_CleanUp
    End If
      
    ' Open password file and
    ' capture specific words
    hFile = FreeFile
    Open gstrPwdFile For Random As #hFile Len = Len(typPWord)
    
    ' start looping thru the file and
    ' capture number of words requested
    For lngIndex = 0 To lngNbrOfWords - 1
    
        ZeroMemory typPWord, Len(typPWord)          ' clear receiving type structure
        Get #hFile, alngRecID(lngIndex), typPWord   ' capture a specific record
                
        If blnUseNumbers Then
        
            ' See if current word is
            ' the one to be replaced
            If lngIndex = lngPosition Then
                strWord = strNumber      ' Use numeric string as a word
            Else
                strWord = typPWord.Word  ' Use word from pwd.dat file
            End If
        Else
            strWord = typPWord.Word  ' Use word from pwd.dat file
        End If
        
        colWords.Add Trim$(strWord)  ' Insert word into collection
    
    Next lngIndex
    
    Close #hFile   ' close PWD.DAT file
    
    ' Format type case of words collected
    If colWords.Count > 0 Then
        FormatPassphrase lngTypeCase, colWords
    End If
    
CreatePassphrase_CleanUp:
    Erase alngRecID()                   ' Always empty arrays when not needed
    ZeroMemory typPWord, Len(typPWord)  ' clear receiving type structure
    On Error GoTo 0                     ' Nullify this error trap
    Exit Sub

CreatePassphrase_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CreatePassphrase_CleanUp
    
End Sub

' ***************************************************************************
' Routine:       CreatePasswords
'
' Description:   Builds the passwords.  Most security systems will not accept
'                a number or special character as the first value in a password
'
' Parameters:    lngPwdCount    - Number of words to create
'                lngWordLength  - Length of each word
'                lngCharsToConv - number of digits in the word
'                lngTypeCase    - Return format of the word
'                blnNumbersOnly - Insert numbers into password (True or False)
'                blnSpecialChars - Insert special chars into password (True or False)
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-JAN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Updated the way random data is gathered.
' 04-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected pointer array sizing and loading
' 14-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected pointer position for character changes if password
'              three characters in length
' ***************************************************************************
Public Sub CreatePasswords(ByVal lngPwdCount As Long, _
                           ByVal lngWordLength As Long, _
                           ByVal lngCharsToConv As Long, _
                           ByVal lngTypeCase As Long, _
                           ByVal blnNumbersOnly As Boolean, _
                           ByVal blnSpecialChars As Boolean, _
                           ByRef colWords As Collection)
                           
    Dim strOnePWord   As String
    Dim intChar       As Integer
    Dim intLoop       As Integer
    Dim lngPointer    As Long
    Dim lngIndex      As Long
    Dim lngLength     As Long
    Dim lngCount      As Long
    Dim alngPointer() As Long
    Dim abytRnd()     As Byte
    Dim abytTemp()    As Byte
    Dim colSpecial    As Collection
  
    Const ROUTINE_NAME As String = "CreatePasswords"

    On Error GoTo CreatePasswords_Error

    Erase abytRnd()     ' Always start with empty arrays
    Erase abytTemp()
    Erase alngPointer()
    
    ' if stop button pressed then leave
    DoEvents
    If gblnStopProcessing Then
        GoTo CreatePasswords_CleanUp
    End If
    
    ' 04-Feb-2009 Load collection with special characters
    '
    ' See if special character flag has been set to TRUE
    If blnSpecialChars Then

        EmptyCollection colSpecial        ' Empty collection
        Set colSpecial = New Collection   ' Instantiate new collection
        lngCount = 0
        
        ' Loop thru special character array and load
        ' only elements that have something in them
        For intLoop = 0 To UBound(gastrChars) - 1
            
            If Len(Trim$(gastrChars(intLoop))) > 0 Then
                colSpecial.Add gastrChars(intLoop)
            End If
            
        Next intLoop
            
        ' Capture number of items in collection
        lngCount = colSpecial.Count
        
        ' If less than one then display
        ' information message and leave
        If lngCount < 1 Then
            InfoMsg "No special characters have been selected.  [CreatePasswords]"
            GoTo CreatePasswords_CleanUp
        End If
        
    End If
    
    With gobjPrng
        
        lngLength = lngPwdCount * lngWordLength                ' desired number of bytes
        abytRnd() = .BuildWithinRange(lngLength + 1, 97, 122)  ' create lowercase random data
        
        ' 04-Feb-2009 Corrected array sizing
        ReDim alngPointer(lngWordLength - 1)
        
        ' 04-Feb-2009 Corrected array looping
        '             to stay within boounds
        ' Load character position in pointer array
        For lngIndex = 0 To UBound(alngPointer) - 1
            alngPointer(lngIndex) = lngIndex + 2
        Next lngIndex
        
        ' Alphabetic only.  No changes.
        If Not blnNumbersOnly And Not blnSpecialChars Then
        
            For lngIndex = 1 To lngLength Step lngWordLength
                
                Erase abytTemp()
                ReDim abytTemp(lngWordLength - 1)
                CopyMemory abytTemp(0), abytRnd(lngIndex - 1), lngWordLength
                
                strOnePWord = StrConv(abytTemp(), vbUnicode)
                colWords.Add strOnePWord
                
                ' see if the stop button was pressed
                DoEvents
                If gblnStopProcessing Then
                    Exit For
                End If
            
            Next lngIndex
        
        Else
            ' insert either numbers, special characters, or both
            
            .RndSeed   ' Reseed VB random number generator
            
            ' see if the stop button was pressed
            DoEvents
            If gblnStopProcessing Then
                GoTo CreatePasswords_CleanUp
            End If
            
            ' if password is only three characters
            ' long then preset the pointer position
            If lngWordLength = 3 Then
                lngPointer = 2
            End If
            
            ' parse random data string and process one password at a time
            For lngIndex = 1 To lngLength Step lngWordLength
                
                Erase abytTemp()
                ReDim abytTemp(lngWordLength - 1)
                
                ' Copy some random characters to a temp array.
                ' Then convert the array to a string.
                CopyMemory abytTemp(0), abytRnd(lngIndex - 1), lngWordLength
                strOnePWord = StrConv(abytTemp(), vbUnicode)
                
                If UBound(alngPointer) > 2 Then
                    .ReshuffleData alngPointer()
                End If
                
                ' see if the stop button was pressed
                DoEvents
                If gblnStopProcessing Then
                    Exit For
                End If
            
                If blnNumbersOnly And Not blnSpecialChars Then
                    ' "Numeric Mix"
                    ' make sure a special character is not the first character
                    For intLoop = 0 To (lngCharsToConv - 1)
                        
                        If lngWordLength = 3 Then
                            lngPointer = IIf(lngPointer = 2, 3, 2)   ' toggle between positions in password
                        Else
                            lngPointer = alngPointer(intLoop)        ' random position in password
                        End If
                        
                        intChar = Int(Rnd() * 10)   ' random value of 0-9
                        Mid$(strOnePWord, lngPointer, 1) = CStr(intChar)
                    
                    Next intLoop
                    
                ElseIf Not blnNumbersOnly And blnSpecialChars Then
                   ' Special characters mix
                    For intLoop = 0 To (lngCharsToConv - 1)
                        
                        If lngWordLength = 3 Then
                            lngPointer = IIf(lngPointer = 2, 3, 2)   ' toggle between positions in password
                        Else
                            lngPointer = alngPointer(intLoop)        ' random position in password
                        End If
                        
                        intChar = Int(Rnd() * lngCount) + 1   ' Random value of 1-26 (if nothing ommitted)
                        Mid$(strOnePWord, lngPointer, 1) = colSpecial.Item(intChar)
                    Next intLoop
                    
                ElseIf blnNumbersOnly And blnSpecialChars Then
                   ' Numbers and special characters mix
                    For intLoop = 0 To (lngCharsToConv - 1)
                        
                        If lngWordLength = 3 Then
                            lngPointer = IIf(lngPointer = 2, 3, 2)   ' toggle between positions in password
                        Else
                            lngPointer = alngPointer(intLoop)        ' random position in password
                        End If
                        
                        intChar = Int(Rnd() * lngCount) + 1   ' Random value of 1-36 (if nothing ommitted)
                        Mid$(strOnePWord, lngPointer, 1) = colSpecial.Item(intChar)
                    
                    Next intLoop
                End If
                          
                ' see if the stop button was pressed
                DoEvents
                If gblnStopProcessing Then
                    Exit For
                End If
            
                ' append one password at a time to the output string
                ' separated by a blank space
                colWords.Add strOnePWord
            
            Next lngIndex
        
        End If
    
    End With
    
    ' see if the stop button was pressed
    DoEvents
    If gblnStopProcessing Then
        GoTo CreatePasswords_CleanUp
    End If
    
    If colWords.Count > 0 Then
        FormatPasswords lngTypeCase, lngWordLength, colWords
    End If
    
CreatePasswords_CleanUp:
    Erase abytRnd()     ' Always empty arrays when not needed
    Erase abytTemp()
    Erase alngPointer()
    EmptyCollection colSpecial

    On Error GoTo 0
    Exit Sub

CreatePasswords_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CreatePasswords_CleanUp
    
End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatPassphrase
'
' Description:   Build the passphrase and then format the final output
'                display
'
' Parameters:    lngTypeCase - Convert the phrase into a particular case
'                colWords - Collection of passwords
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' 17-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated this routine to process in a more logical way.
' ***************************************************************************
Private Sub FormatPassphrase(ByVal lngTypeCase As Long, _
                             ByRef colWords As Collection)

    ' Called by CreatePassphrase()
    
    Dim strChar        As String
    Dim strData        As String
    Dim astrChar()     As String
    Dim lngIndex       As Long
    Dim lngLength      As Long
    Dim lngPointer     As Long
    Dim lngCharsToConv As Long
    Dim alngPointer()  As Long
    
    Erase alngPointer()   ' Start with empty arrays
    Erase astrChar()
    strData = vbNullString
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    If colWords.Count < 1 Then
        Exit Sub
    End If
   
    ' Transfer words to a string delimtied by a space
    For lngIndex = 1 To colWords.Count
        strData = strData & colWords.Item(lngIndex) & " "
    Next lngIndex
    
    strData = Trim$(strData)   ' Remove leading and trailing blanks
    
    Select Case lngTypeCase
           
           Case 0  ' All lowercase letters
                strData = LCase$(strData)
    
           Case 1  ' All Uppercase letters
                strData = UCase$(strData)
    
           Case 2  ' All Propercase letters (first character uppercase)
                strData = StrConv(strData, vbProperCase)
    
           Case 3  ' Mixed case letters
                strData = LCase$(strData)          ' convert to lowercase
                lngLength = Len(strData)           ' Capture length of passphrase
                ReDim alngPointer(lngLength - 1)   ' Resize pointer array
                
                ' load pointer array
                For lngIndex = 1 To lngLength
                    alngPointer(lngIndex - 1) = lngIndex
                Next lngIndex
        
                gobjPrng.ReshuffleData alngPointer()   ' Mix pointers

                ' Calculate how many characters need
                ' to be converted to upper case
                lngCharsToConv = Int(lngLength / 2)
                lngCharsToConv = Int(Rnd() * lngCharsToConv) + Int(lngCharsToConv / 2)
                    
                For lngIndex = 1 To lngCharsToConv
                    
                    lngPointer = alngPointer(lngIndex - 1)   ' Position in password
                    strChar = Mid$(strData, lngPointer, 1)   ' Capture one character
                    
                    ' Skip over if this position is a blank space
                    If Mid$(strData, lngPointer, 1) <> Chr$(32) Then
                        ' Convert character to uppercase
                        Mid$(strData, lngPointer, 1) = UCase$(strChar)
                    End If
                    
                Next lngIndex
                
    End Select
    
    EmptyCollection colWords        ' Empty collection
    Set colWords = New Collection   ' Create new collection
    
    strData = strData & Chr$(32)           ' Add a trailing blank space
    astrChar() = Split(strData, Chr$(32))  ' convert string to array
    
    ' Reload collection
    For lngIndex = 0 To UBound(astrChar) - 1
        colWords.Add astrChar(lngIndex)
    Next lngIndex
    
FormatPassphrase_CleanUp:
    Erase alngPointer()   ' Always empty arrays when not needed
    Erase astrChar()

End Sub

' ***************************************************************************
' Routine:       FormatPasswords
'
' Description:   Build the passphrase and then format the final output
'                display
'
' Parameters:    lngTypeCase - Convert passwords into a particular case
'                lngWordLength - Length of each password
'                colWords - Collection of passwords
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' 17-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated this routine to process in a more logical way.
' ***************************************************************************
Private Sub FormatPasswords(ByVal lngTypeCase As Long, _
                            ByVal lngWordLength As Long, _
                            ByRef colWords As Collection)

    ' Called by CreatePasswords()
    
    Dim strChar        As String
    Dim strData        As String
    Dim lngIdx         As Long
    Dim lngIndex       As Long
    Dim bytPointer     As Byte
    Dim bytCharsToConv As Byte
    Dim abytPointer()  As Byte
    Dim colTemp        As Collection
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    If colWords.Count < 1 Then
        Exit Sub
    End If
   
    Erase abytPointer()            ' Start with an empty array
    EmptyCollection colTemp        ' Start with an empty collection
    Set colTemp = New Collection   ' Create new collection
    bytPointer = 2
    
    For lngIndex = 1 To colWords.Count
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        strData = colWords.Item(lngIndex)   ' Get one password
    
        Select Case lngTypeCase
               
               Case 0  ' All lowercase letters
                    strData = LCase$(strData)
        
               Case 1  ' All Uppercase letters
                    strData = UCase$(strData)
        
               Case 2  ' All Propercase letters (first character uppercase)
                    strData = StrConv(strData, vbProperCase)
    
               Case 3  ' Mixed case letters
                    If lngWordLength = 3 Then
                                                 
                        bytPointer = IIf(bytPointer = 2, 3, 2)           ' toggle between positions in password
                        strChar = Mid$(strData, bytPointer, 1)           ' Capture one character
                        Mid$(strData, bytPointer, 1) = UCase$(strChar)   ' Convert character to uppercase
                        
                    Else
                        ReDim abytPointer(lngWordLength - 1)
                        
                        ' Load position pointer array starting with
                        ' first character position in password.
                        For lngIdx = 1 To lngWordLength
                            abytPointer(lngIdx - 1) = lngIdx
                        Next lngIdx
                        
                        gobjPrng.ReshuffleData abytPointer()   ' Mix pointers
        
                        ' Calculate how many characters need
                        ' to be converted to upper case
                        bytCharsToConv = Int(lngWordLength / 2)
                        bytCharsToConv = Int(Rnd() * bytCharsToConv) + Int(bytCharsToConv / 2)
                        
                        For lngIdx = 1 To bytCharsToConv
                            
                            bytPointer = abytPointer(lngIdx - 1)             ' Position in password
                            strChar = Mid$(strData, bytPointer, 1)           ' Capture one character
                            Mid$(strData, bytPointer, 1) = UCase$(strChar)   ' Convert character to uppercase
                            
                        Next lngIdx
                    
                    End If
        End Select
            
        colTemp.Add strData   ' Update collection
        
    Next lngIndex
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo FormatPasswords_CleanUp
    End If
    
    EmptyCollection colWords       ' Empty collection
    Set colWords = New Collection  ' Create new collection
    
    ' Reload collection
    For lngIndex = 1 To colTemp.Count
        colWords.Add colTemp.Item(lngIndex)
    Next lngIndex
    
FormatPasswords_CleanUp:
    Erase abytPointer()      ' Always empty arrays when not needed
    EmptyCollection colTemp  ' Empty collection
    
End Sub

