Attribute VB_Name = "barcode"

'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************

'START OF DECLARACTIONS
Private I As Integer
Private F As Integer
Private DataToPrint As String
Private DataToEncode As String
Private OnlyCorrectData As String
Private PrintableString As String
Private Encoding As String
Private WeightedTotal As Long
Private WeightValue As Integer
Private CurrentValue As Long
Private CheckDigitValue As Integer
Private Factor As Integer
Private CheckDigit As Integer
Private CurrentEncoding As String
Private NewLine As String
Private msg As String
Private CurrentChar As String
Private CurrentCharNum As Integer
Private C128_StartA As String
Private C128_StartB As String
Private C128_StartC As String
Private C128_Stop As String
Private C128Start As String
Private C128CheckDigit As String
Private StartCode As String
Private StopCode As String
Private Fnc1 As String
Private LeadingDigit As Integer
Private EAN2AddOn As String
Private EAN5AddOn As String
Private EANAddOnToPrint As String
Private HumanReadableText As String
Private StringLength As Integer
Private CorrectFNC As Integer
'END OF DECLARACTIONS


Public Function Code128(DataToFormat As String, Optional ReturnType As Integer, Optional ApplyTilde As Boolean) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
    CorrectFNC = 0
    PrintableString = ""
    
    'Additional logic needed in case ReturnType is not entered
    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
    
    'Additions for ApplyTilde 2-11-2005
    'in case ApplyTilde is null, set it to false
    If ApplyTilde <> True Then ApplyTilde = False
    
    If ApplyTilde Then
        DataToEncode = DataToFormat
        DataToFormat = ""
        OnlyCorrectData = ""
        StringLength = Len(DataToEncode)
        For I = 1 To StringLength
            If (I < StringLength - 2) And Mid(DataToEncode, I, 2) = "~m" And IsNumeric(Mid(DataToEncode, I + 2, 2)) Then
                WeightValue = Val(Mid(DataToEncode, I + 2, 2))
                If (I - WeightValue) < 1 Then WeightValue = I - 1
                CheckDigitValue = MOD10(Mid(DataToEncode, I - WeightValue, WeightValue))
                OnlyCorrectData = OnlyCorrectData & ChrW(CheckDigitValue + 48)
                I = I + 3
            ElseIf (I < StringLength - 2) And Mid(DataToEncode, I, 1) = "~" And IsNumeric(Mid(DataToEncode, I + 1, 3)) Then
                CurrentCharNum = Val(Mid(DataToEncode, I + 1, 3))
                OnlyCorrectData = OnlyCorrectData & ChrW(CurrentCharNum)
                I = I + 3
            Else
               OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
            End If
        Next I
        DataToFormat = OnlyCorrectData
        DataToEncode = ""
    End If
    
    'Here we select character set A, B or C for the START character
    StringLength = Len(DataToFormat)
    CurrentCharNum = AscW(Mid(DataToFormat, 1, 1))
    If CurrentCharNum < 32 Then C128Start = ChrW(203)
    If CurrentCharNum > 31 And CurrentCharNum < 127 Then C128Start = ChrW(204)
    If CurrentCharNum = 197 Then C128Start = ChrW(204) 'Added 2-18-05 for FNC2
    If ((StringLength > 4) And IsNumeric(Mid(DataToFormat, 1, 4))) Then C128Start = ChrW(205)
    '202 & 212-215 is for the FNC1, with this Start C is mandatory
    If CurrentCharNum = 202 Then C128Start = ChrW(205)
    If CurrentCharNum = 212 Then C128Start = ChrW(205)
    If CurrentCharNum = 213 Then C128Start = ChrW(205)
    If CurrentCharNum = 214 Then C128Start = ChrW(205)
    If CurrentCharNum = 215 Then C128Start = ChrW(205)
    If C128Start = ChrW(203) Then CurrentEncoding = "A"
    If C128Start = ChrW(204) Then CurrentEncoding = "B"
    If C128Start = ChrW(205) Then CurrentEncoding = "C"
    For I = 1 To StringLength
    
        'Added 2-18-05 for FNC2 / check for FNC2 which is ASCII 197 in any set other than C
        If (CurrentCharNum = 197) Then
            If CurrentEncoding = "C" Then 'switch to B
                DataToEncode = DataToEncode & ChrW(200)
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(197)
            I = I + 1
        End If
    
        'check for FNC1 in any set which is ASCII 202 and ASCII 212-215
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
        If ((CurrentCharNum = 202) Or (CurrentCharNum = 212) Or (CurrentCharNum = 213) Or (CurrentCharNum = 214) Or (CurrentCharNum = 215)) Then
            DataToEncode = DataToEncode & ChrW(202)
        'check for switching to character set C
        ElseIf ((I < StringLength - 2) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (IsNumeric(Mid(DataToFormat, I, 4)))) Or ((I < StringLength) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (CurrentEncoding = "C")) Then
        'switch to set C if not already in it
            If CurrentEncoding <> "C" Then DataToEncode = DataToEncode & ChrW(199)
            CurrentEncoding = "C"
            CurrentChar = (Mid(DataToFormat, I, 2))
            CurrentValue = CInt(CurrentChar)
        'set the CurrentValue to the number of String CurrentChar
            If (CurrentValue < 95 And CurrentValue > 0) Then DataToEncode = DataToEncode & ChrW(CurrentValue + 32)
            If CurrentValue > 94 Then DataToEncode = DataToEncode & ChrW(CurrentValue + 100)
            If CurrentValue = 0 Then DataToEncode = DataToEncode & ChrW(194)
            I = I + 1
        'check for switching to character set A
        ElseIf (I <= StringLength) And ((AscW(Mid(DataToFormat, I, 1)) < 31) Or ((CurrentEncoding = "A") And (AscW(Mid(DataToFormat, I, 1)) > 32 And (AscW(Mid(DataToFormat, I, 1))) < 96))) Then
        'switch to set A if not already in it
            If CurrentEncoding <> "A" Then DataToEncode = DataToEncode & ChrW(201)
            CurrentEncoding = "A"
        'Get the ASCII value of the next character
            CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & ChrW(194)
            ElseIf CurrentCharNum < 32 Then
                DataToEncode = DataToEncode & ChrW(CurrentCharNum + 96)
            ElseIf CurrentCharNum > 32 Then
                DataToEncode = DataToEncode & ChrW(CurrentCharNum)
            End If
        'check for switching to character set B
        ElseIf (I <= StringLength) And ((AscW(Mid(DataToFormat, I, 1))) > 31 And (AscW(Mid(DataToFormat, I, 1)))) < 127 Then
        'switch to set B if not already in it
            If CurrentEncoding <> "B" Then DataToEncode = DataToEncode & ChrW(200)
            CurrentEncoding = "B"
        'Get the ASCII value of the next character
            CurrentCharNum = (AscW(Mid(DataToFormat, I, 1)))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & ChrW(194)
            Else
                DataToEncode = DataToEncode & ChrW(CurrentCharNum)
            End If
        End If
    Next I
    
    HumanReadableText = ""
'FORMAT TEXT FOR AIs
    StringLength = Len(DataToFormat)
    For I = 1 To StringLength
    CorrectFNC = 0
    'Get ASCII value of each character
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
    'Check for FNC1
        If ((I < StringLength - 2) And ((CurrentCharNum = 202) Or ((CurrentCharNum > 211) And (CurrentCharNum < 219)))) Then
        'It appears that there is an AI
        'Get the value of each number pair (ex: 5 and 6 = 5*10+6 =56)
            CurrentChar = (Mid(DataToFormat, I + 1, 2))
            CurrentCharNum = CInt(CurrentChar)
        'Is 2 digit AI by entering ASCII 212?
            If ((CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 212)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 2)) & ") "
                I = I + 2
                CorrectFNC = 1
        'Is 3 digit AI by entering ASCII 213?
            ElseIf ((I < StringLength - 3) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 213)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 3)) & ") "
                I = I + 3
                CorrectFNC = 1
        'Is 4 digit AI by entering ASCII 214?
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 214)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
        'Is 5 digit AI by entering ASCII 215?
            ElseIf ((I < StringLength - 5) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 215)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 5)) & ") "
                I = I + 5
                CorrectFNC = 1
        'Is 6 digit AI by entering ASCII 216?
            ElseIf ((I < StringLength - 6) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 216)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 6)) & ") "
                I = I + 6
                CorrectFNC = 1
        'Is 7 digit AI by entering ASCII 217?
            ElseIf ((I < StringLength - 7) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 217)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 7)) & ") "
                I = I + 7
                CorrectFNC = 1
        'Is 8 digit AI by entering ASCII 218?
            ElseIf ((I < StringLength - 8) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 218)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 8)) & ") "
                I = I + 8
                CorrectFNC = 1
        'Is 4 digit AI by detection?
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0) And ((CurrentCharNum <= 81 And CurrentCharNum >= 80) Or (CurrentCharNum <= 34 And CurrentCharNum >= 31))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
        'Is 3 digit AI by detection?
            ElseIf ((I < StringLength - 3) And (CorrectFNC = 0) And ((CurrentCharNum <= 49 And CurrentCharNum >= 40) Or (CurrentCharNum <= 25 And CurrentCharNum >= 23))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 3)) & ") "
                I = I + 3
                CorrectFNC = 1
        'Is 2 digit AI by detection?
            ElseIf ((CurrentCharNum <= 30 And (CorrectFNC = 0) And CurrentCharNum >= 0) Or (CurrentCharNum <= 99 And CurrentCharNum >= 90)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 2)) & ") "
                I = I + 2
                CorrectFNC = 1
        'If no AI was detected, set default to 4 digit AI:
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
            End If
        ElseIf (AscW(Mid(DataToFormat, I, 1)) < 32) Then
            HumanReadableText = HumanReadableText & " "
        ElseIf ((AscW(Mid(DataToFormat, I, 1)) > 31) And (AscW(Mid(DataToFormat, I, 1)) < 128)) Then
            HumanReadableText = HumanReadableText & Mid(DataToFormat, I, 1)
        End If
    Next I
    DataToFormat = ""
    '<<<< Calculate Modulo 103 Check Digit >>>>
    WeightedTotal = AscW(C128Start) - 100
    StringLength = Len(DataToEncode)
    For I = 1 To StringLength
        CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
        If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
        If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
        If CurrentCharNum = 194 Then CurrentValue = 0
        CurrentValue = CurrentValue * I
        WeightedTotal = WeightedTotal + CurrentValue
        If CurrentCharNum = 32 Then CurrentCharNum = 194
        PrintableString = PrintableString & ChrW(CurrentCharNum)
    Next I
    CheckDigitValue = (WeightedTotal Mod 103)
    If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = ChrW(CheckDigitValue + 32)
    If CheckDigitValue > 94 Then C128CheckDigit = ChrW(CheckDigitValue + 100)
    If CheckDigitValue = 0 Then C128CheckDigit = ChrW(194)
    DataToEncode = ""
    'ReturnType 0 returns data formatted to the barcode font
    If ReturnType = 0 Then Code128 = C128Start & PrintableString & C128CheckDigit & ChrW(206) & " "
    'ReturnType 1 returns data formatted for human readable text
    If ReturnType = 1 Then Code128 = HumanReadableText
    'ReturnType 2 returns the check digit for the data supplied
    If ReturnType = 2 Then Code128 = C128CheckDigit
End Function



Public Function Code128a(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     PrintableString = ""
     WeightedTotal = 103
     PrintableString = ChrW(203)
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
          CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
          If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
          If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
          CurrentValue = CurrentValue * I
          WeightedTotal = WeightedTotal + CurrentValue
          If CurrentCharNum = 32 Then CurrentCharNum = 194
          PrintableString = PrintableString & ChrW(CurrentCharNum)
     Next I
     CheckDigitValue = (WeightedTotal Mod 103)
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = ChrW(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128CheckDigit = ChrW(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128CheckDigit = ChrW(194)
     PrintableString = PrintableString & C128CheckDigit & ChrW(206) & " "
     Code128a = PrintableString
End Function



Public Function Code128b(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     PrintableString = ""
     WeightedTotal = 104
     PrintableString = ChrW(204)
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
          CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
          If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
          If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
          CurrentValue = CurrentValue * I
          WeightedTotal = WeightedTotal + CurrentValue
          If CurrentCharNum = 32 Then CurrentCharNum = 194
          PrintableString = PrintableString & ChrW(CurrentCharNum)
     Next I
     CheckDigitValue = (WeightedTotal Mod 103)
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = ChrW(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128CheckDigit = ChrW(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128CheckDigit = ChrW(194)
     PrintableString = PrintableString & C128CheckDigit & ChrW(206) & " "
     Code128b = PrintableString
End Function


Public Function Code128c(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
    'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     PrintableString = ""
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
     PrintableString = ChrW(205)
     WeightedTotal = 105
     WeightValue = 1
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength Step 2
          CurrentValue = CInt(Mid(DataToEncode, I, 2))
          If CurrentValue < 95 And CurrentValue > 0 Then PrintableString = PrintableString & ChrW(CurrentValue + 32)
          If CurrentValue > 94 Then PrintableString = PrintableString & ChrW(CurrentValue + 100)
          If CurrentValue = 0 Then PrintableString = PrintableString & ChrW(194)
          CurrentValue = CurrentValue * WeightValue
          WeightedTotal = WeightedTotal + CurrentValue
          WeightValue = WeightValue + 1
     Next I
     CheckDigitValue = (WeightedTotal Mod 103)
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = ChrW(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128CheckDigit = ChrW(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128CheckDigit = ChrW(194)
     If ReturnType = 0 Then Code128c = PrintableString & C128CheckDigit & ChrW(206) & " "
     If ReturnType = 1 Then Code128c = DataToEncode & CheckDigitValue
     If ReturnType = 2 Then Code128c = Str(CheckDigitValue)
End Function


Public Function I2of5(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************

     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
'Check for an even number of digits, add 0 if not even
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
'Assign start and stop codes
     StartCode = ChrW(203)
     StopCode = ChrW(204)
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentCharNum = Val((Mid(DataToEncode, I, 2)))
    'Get the ASCII value of CurrentChar according to chart by to the value
          If CurrentCharNum < 94 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 33)
          If CurrentCharNum > 93 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 103)
     Next I
'Get Printable String
     PrintableString = StartCode + DataToPrint + StopCode & " "
'Return PrintableString
     I2of5 = PrintableString
End Function



Public Function USPS_EAN128(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
'
' Used for 22 digit USPS special services labels such as delivery confirmation in
' EAN128 with Code 128 fonts. This new EAN128 format is mandatory as of
' January 10, 2004 according to the USPS Delivery Confirmation Service
' defined in the September 2002 version of Publication 91. Enter a 19 or
' 20 digit number; only the first 19 are used. This number is made up of
' the following:  2 digit service code + 9 digit customer ID + 8 digit
' sequential package ID + MOD 10 check digit that can be calculated by
' this function if excluded. In this function, the application identifier
' of 91 is automatically added for you.
'
' Other USPS EAN128 barcode types must be created by calling Code128() with the appropriate
' ASCII 0202 and AIs included as documented at:
' http://www.idautomation.com/code128faq.html#EAN128andUCC128
'
' Check to make sure data is numeric and remove dashes, etc.
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     OnlyCorrectData = ""
     Dim DataForCheck As String
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) > "19" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 19))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "19" Then OnlyCorrectData = "0000000000000000000"
'Add in the AI of 91
     DataToEncode = "91" & OnlyCorrectData
'Get the MOD 10 Check Digit
     CheckDigit = MOD10(DataToEncode)
'Now that we have calculated the MOD 10 for the data, send the string
'to the Code128() funtion. This function will:
' - Add in the start and stop codes
' - Add in the AI and START C
' - Calculate the MOD 103 required when using Code 128
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then USPS_EAN128 = Code128(ChrW(202) & DataToEncode & CheckDigit, 0)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then USPS_EAN128 = Mid(DataToEncode, 1, 4) & " " & Mid(DataToEncode, 5, 4) & " " & Mid(DataToEncode, 9, 4) & " " & Mid(DataToEncode, 13, 4) & " " & Mid(DataToEncode, 17, 4) & " " & Mid(DataToEncode, 21, 1) & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then USPS_EAN128 = Str(CheckDigit)
End Function


Public Function Code39Mod43(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     'DataToEncode = RTrim(DataToEncode)
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     DataToEncode = UCase(DataToEncode)
     DataToPrint = ""
     OnlyCorrectData = ""
'only pass correct data
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (AscW(Mid(DataToEncode, I, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    'Space
          If CurrentCharNum = 32 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '-
          If CurrentCharNum = 45 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '.
          If CurrentCharNum = 46 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '$
          If CurrentCharNum = 36 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '/
          If CurrentCharNum = 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '+
          If CurrentCharNum = 43 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    '%
          If CurrentCharNum = 37 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
     WeightedTotal = 0
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (AscW(Mid(DataToEncode, I, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then CurrentValue = CurrentCharNum - 48
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then CurrentValue = CurrentCharNum - 55
    'Space
          If CurrentCharNum = 32 Then CurrentValue = 38
    '-
          If CurrentCharNum = 45 Then CurrentValue = 36
    '.
          If CurrentCharNum = 46 Then CurrentValue = 37
    '$
          If CurrentCharNum = 36 Then CurrentValue = 39
    '/
          If CurrentCharNum = 47 Then CurrentValue = 40
    '+
          If CurrentCharNum = 43 Then CurrentValue = 41
    '%
          If CurrentCharNum = 37 Then CurrentValue = 42
    'To print the barcode symbol representing a space you will
    'to type or print "=" (the equal character) instead of a space character.
          If CurrentCharNum = 32 Then CurrentCharNum = 61
    'gather data to print
          DataToPrint = DataToPrint & ChrW(CurrentCharNum)
    'add the values together
          WeightedTotal = WeightedTotal + CurrentValue
     Next I
'divide the WeightedTotal by 43 and get the remainder, this is the CheckDigit
     CheckDigitValue = (WeightedTotal Mod 43)
    'Assign values to characters
    '0-9
     If CheckDigitValue < 10 Then CheckDigit = CheckDigitValue + 48
    'A-Z
     If CheckDigitValue < 36 And CheckDigitValue > 9 Then CheckDigit = CheckDigitValue + 55
    'Space
     If CheckDigitValue = 38 Then CheckDigit = 61
    '-
     If CheckDigitValue = 36 Then CheckDigit = 45
    '.
     If CheckDigitValue = 37 Then CheckDigit = 46
    '$
     If CheckDigitValue = 39 Then CheckDigit = 36
    '/
     If CheckDigitValue = 40 Then CheckDigit = 47
    '+
     If CheckDigitValue = 41 Then CheckDigit = 43
    '%
     If CheckDigitValue = 42 Then CheckDigit = 37
     
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then Code39Mod43 = "!" & DataToPrint & ChrW(CheckDigit) & "!" & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then Code39Mod43 = DataToPrint & ChrW(CheckDigit)
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then Code39Mod43 = ChrW(CheckDigit)
End Function


Public Function Code39(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************

     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
'Check for spaces in code
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get each character one at a time
          CurrentChar = (Mid(DataToEncode, I, 1))
    'To print the barcode symbol representing a space you will
    'to type or print "=" (the equal character) instead of a space character.
          If CurrentChar = " " Then CurrentChar = "="
          DataToPrint = DataToPrint & CurrentChar
     Next I
'Get Printable String
     PrintableString = "!" & DataToPrint & "!" & " "
'Return PrintableString
     Code39 = PrintableString
End Function





Public Function I2of5Mod10(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToPrint = ""
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     Factor = 3
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
'Add check digit to number to DataToEncode
     DataToEncode = DataToEncode & CheckDigit
'Check for an even number of digits, add 0 if not even
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentCharNum = (Mid(DataToEncode, I, 2))
    'Get the ASCII value of CurrentChar according to chart by to the value
          If CurrentCharNum < 94 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 33)
          If CurrentCharNum > 93 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 103)
     Next I
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then I2of5Mod10 = ChrW(203) & DataToPrint & ChrW(204) & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then I2of5Mod10 = DataToEncode
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then I2of5Mod10 = Str$(CheckDigit)
End Function



Public Function MSI(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
    'Additional logic needed in case ReturnType is not entered
    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
' The MSI encoding function will only accept digits.  Any non-numeric characters
' will be discarded
    Dim DataToPrint As String       'output for function
    Dim OnlyCorrectData As String   'Only numeric characters pulled from DataToEncode
    Dim StringLength As Long        'Length of string
    Dim Idx As Integer              'for loop counter
    Dim OddNumbers As String        'String of odd position numbers used to create check digit
    Dim EvenNumberSum As Long       'all of the even position numbers added up
    Dim OddNumberProduct As Long    'Product of OddNumbers variable
    Dim sOddNumberProduct As String 'String version of OddNumberProduct variable
    Dim OddNumberSum As Long        'Sum of individual digits in sOddNumberProduct
    Dim OddDigit As Boolean         'Used to determine even/odd position digits.
    Dim CheckDigit As String        'This is the CheckDigit
    DataToPrint = ""
    OnlyCorrectData = ""
    'Take off any extra spaces
    DataToEncode = Trim(DataToEncode)
    
    'Check to make sure data is numeric and remove dashes, etc.
     StringLength = Len(DataToEncode)
     For Idx = 1 To StringLength
        'Add all numbers to OnlyCorrectData string
        If IsNumeric(Mid(DataToEncode, Idx, 1)) = True Then
            OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, Idx, 1)
        End If
     Next Idx
     
     DataToEncode = OnlyCorrectData
     
     '<<<< Calculate Check Digit >>>>
     'To create the check digit follow these steps
     '1)Starting from the units position, create a new number with all of the odd
     '  position digits in their original sequence.
     '2)Multiply this new number by 2.
     '3)Add all of the digits of the product from step two.
     '4)Add all of the digits not used in step one to the result in step three.
     '5)Determine the smallest number which when added to the result in step four
     '  will result in a multiple of 10. This is the check character.

    'Step 1 -- Create a new number of the odd position digits starting from the right and going left, but store the
    'digits from left to right.
    'We will create the odd position number & prepare for Step 4 by getting the sum of all even position charactesr
    StringLength = Len(DataToEncode)
    OddNumbers = ""
    OddDigit = True
    EvenNumberSum = 0
    For Idx = StringLength To 1 Step -1
        If OddDigit = True Then
            OddNumbers = Mid(DataToEncode, Idx, 1) & OddNumbers
            OddDigit = False
        Else
            EvenNumberSum = EvenNumberSum + Val(Mid(DataToEncode, Idx, 1))
            OddDigit = True
        End If
    Next Idx
    
    'Step 2 -- Multiply this new number by 2.
    OddNumberProduct = Val(OddNumbers) * 2

    'Step 3 -- Add all of the digits of the product from step two.
    sOddNumberProduct = Format(OddNumberProduct)
    StringLength = Len(sOddNumberProduct)
    OddNumberSum = 0

    For Idx = 1 To StringLength
        OddNumberSum = OddNumberSum + Val(Mid(sOddNumberProduct, Idx, 1))
    Next Idx
    
    'Step 4 -- Add all of the digits not used in step one to the result in step three.
    'We will store the result in OddNumberSum just so we don't have to create another variable
    OddNumberSum = OddNumberSum + EvenNumberSum
    
    'Step 5 -- Determine the smallest number which when added to the result in step four
    '  will result in a multiple of 10. This is the check character.
    OddNumberSum = OddNumberSum Mod 10
    If OddNumberSum <> 0 Then
        CheckDigit = Format(10 - OddNumberSum)
    Else
        CheckDigit = "0"
    End If
    
    Select Case ReturnType
        Case 0  'Returns formatted data for barcode
            DataToPrint = "(" & DataToEncode & CheckDigit & ")" & " "
        Case 1  'Returns data formatted for human readable text.  Which means all of the invalid characters where
                'stripped out.
            DataToPrint = DataToEncode
        Case 2  'Returns just the check digit
            DataToPrint = CheckDigit
    End Select
    
    MSI = DataToPrint
    
End Function


Public Function UPCa(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
'Remove check digits if they added one
     If Len(OnlyCorrectData) < "11" Then OnlyCorrectData = "00000000000"
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = "00000000000"
     If Len(OnlyCorrectData) > "18" Then OnlyCorrectData = "00000000000"
     If Len(OnlyCorrectData) = "12" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 11)
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 2))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 5))
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 16 Then EAN5AddOn = Mid(OnlyCorrectData, 12, 5)
     If Len(OnlyCorrectData) = 13 Then EAN2AddOn = Mid(OnlyCorrectData, 12, 2)
'split 12 digit number from add-on

     DataToEncode = Mid(OnlyCorrectData, 1, 11)
'<<<< Calculate Check Digit >>>>
     Factor = 3
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the number + WeightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
     DataToEncode = DataToEncode & CheckDigit
'Now that have the total number including the check digit, determine character to print
'for proper barcoding
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get the ASCII value of each number
          CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
    'Print different barcodes according to the location of the CurrentChar
          Select Case I
          Case 1
        'For the first character print the human readable character, the normal
        'guard pattern and then the barcode without the human readable character
               If ChrW(CurrentCharNum) > 4 Then DataToPrint = ChrW(CurrentCharNum + 64) & "(" & ChrW(CurrentCharNum + 49)
               If ChrW(CurrentCharNum) < 5 Then DataToPrint = ChrW(CurrentCharNum + 37) & "(" & ChrW(CurrentCharNum + 49)
          Case 2
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 3
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 4
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 5
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 6
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & ChrW(CurrentCharNum) & "*"
          Case 7
        'Add 27 to the ASII value of characters 6-12 to print from character set+ C
        'this is required when printing to the right of the center guard pattern
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 8
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 9
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 10
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 11
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 12
        'For the last character print the barcode without the human readable character,
        'the normal guard pattern and then the human readable character.
               If ChrW(CurrentCharNum) > 4 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 59) & "(" & ChrW(CurrentCharNum + 64)
               If ChrW(CurrentCharNum) < 5 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 59) & "(" & ChrW(CurrentCharNum + 37)
          End Select
     Next I
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          WeightedTotal = 0
          For I = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, I, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then WeightedTotal = WeightedTotal + CurrentCharNum * 3
               If Factor = 1 Then WeightedTotal = WeightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next I
    'Find the CheckDigit by extracting the right-most number from WeightedTotal
          CheckDigit = Val(Right$(WeightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For I = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
            'EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
                    EANAddOnToPrint = ChrW(43) & EANAddOnToPrint & ChrW(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For I = 0 To 99 Step 4
               If Val(EAN2AddOn) = I Then Encoding = "AA"
               If Val(EAN2AddOn) = I + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = I + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = I + 3 Then Encoding = "BB"
          Next I
    'Now that we have the total number including the encoding
    'determine what to print
          For I = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
            'EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
                    EANAddOnToPrint = ChrW(43) & EANAddOnToPrint & ChrW(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
'Get Printable String
     PrintableString = DataToPrint & EANAddOnToPrint & " "
'Return PrintableString
     UPCa = PrintableString
End Function

Public Function UPCe(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToPrint = ""
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
'Remove check digits if they added one
     If Len(OnlyCorrectData) < "11" Then OnlyCorrectData = "00005000000"
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = "00005000000"
     If Len(OnlyCorrectData) > "18" Then OnlyCorrectData = "00005000000"
     If Len(OnlyCorrectData) = "12" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 11)
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 2))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 5))
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 16 Then EAN5AddOn = Mid(OnlyCorrectData, 12, 5)
     If Len(OnlyCorrectData) = 13 Then EAN2AddOn = Mid(OnlyCorrectData, 12, 2)
'split 12 digit number from add-on

     DataToEncode = Mid(OnlyCorrectData, 1, 11)
     
'<<<< Calculate Check Digit >>>>
     Factor = 3
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the number + WeightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
     
     DataToEncode = DataToEncode & CheckDigit
'Compress UPC-A to UPC-E if possible
     Dim D1 As String
     Dim D2 As String
     Dim D3 As String
     Dim D4 As String
     Dim D5 As String
     Dim D6 As String
     Dim D7 As String
     Dim D8 As String
     Dim D9 As String
     Dim D10 As String
     Dim D11 As String
     Dim D12 As String
     D1 = Mid(DataToEncode, 1, 1)
     D2 = Mid(DataToEncode, 2, 1)
     D3 = Mid(DataToEncode, 3, 1)
     D4 = Mid(DataToEncode, 4, 1)
     D5 = Mid(DataToEncode, 5, 1)
     D6 = Mid(DataToEncode, 6, 1)
     D7 = Mid(DataToEncode, 7, 1)
     D8 = Mid(DataToEncode, 8, 1)
     D9 = Mid(DataToEncode, 9, 1)
     D10 = Mid(DataToEncode, 10, 1)
     D11 = Mid(DataToEncode, 11, 1)
     D12 = Mid(DataToEncode, 12, 1)
'Condition A
     If (D11 = "5" Or D11 = "6" Or D11 = "7" Or D11 = "8" Or D11 = "9") And D6 <> "0" And (D7 = "0" And D8 = "0" And D9 = "0" And D10 = "0") Then
          DataToEncode = D2 & D3 & D4 & D5 & D6 & D11
     End If
'Condition B
     If (D6 = "0" And D7 = "0" And D8 = "0" And D9 = "0" And D10 = "0") And D5 <> "0" Then
          DataToEncode = D2 & D3 & D4 & D5 & D11 & "4"
     End If
'Condition C
     If (D5 = "0" And D6 = "0" And D7 = "0" And D8 = "0") And (D4 = "1" Or D4 = "2" Or D4 = "0") Then
          DataToEncode = D2 & D3 & D9 & D10 & D11 & D4
     End If
'Condition D
     If (D5 = "0" And D6 = "0" And D7 = "0" And D8 = "0" And D9 = "0") And (D4 = "3" Or D4 = "4" Or D4 = "5" Or D4 = "6" Or D4 = "7" Or D4 = "8" Or D4 = "9") Then
          DataToEncode = D2 & D3 & D4 & D10 & D11 & "3"
     End If
'
'Run UPC-E compression only if DataToEncode = 6
     If Len(DataToEncode) = 6 Then
    'Now we must encode the check character into the symbol
    'by using variable parity between character sets A and B
          Select Case D12
          Case "0"
               Encoding = "BBBAAA"
          Case "1"
               Encoding = "BBABAA"
          Case "2"
               Encoding = "BBAABA"
          Case "3"
               Encoding = "BBAAAB"
          Case "4"
               Encoding = "BABBAA"
          Case "5"
               Encoding = "BAABBA"
          Case "6"
               Encoding = "BAAABB"
          Case "7"
               Encoding = "BABABA"
          Case "8"
               Encoding = "BABAAB"
          Case "9"
               Encoding = "BAABAB"
          End Select
          StringLength = Len(DataToEncode)
          For I = 1 To StringLength
        'Get the ASCII value of each number
               CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum)
               Case "B"
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 17)
               End Select
        'add in the 1st character along with guard patterns
               Select Case I
               Case 1
            'For the LeadingDigit print the human readable character,
            'the normal guard pattern and then the rest of the barcode
                    DataToPrint = ChrW(85) & "(" & DataToPrint
               Case 6
            'Print the SPECIAL guard pattern and check character
                    If CInt(D12) > 4 Then DataToPrint = DataToPrint & ")" & ChrW(AscW(D12) + 64)
                    If CInt(D12) < 5 Then DataToPrint = DataToPrint & ")" & ChrW(AscW(D12) + 37)
                    
               End Select
          Next I
     End If
     
'determine character to print
'for proper upc-a barcoding
     If Len(DataToEncode) <> 6 Then
          StringLength = Len(DataToEncode)
          For I = 1 To StringLength
        'Get the ASCII value of each number
               CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
        'Print different barcodes according to the location of the CurrentChar
               Select Case I
               Case 1
            'For the first character print the human readable character, the normal
            'guard pattern and then the barcode without the human readable character
                    If ChrW(CurrentCharNum) > 4 Then DataToPrint = ChrW(CurrentCharNum + 64) & "(" & ChrW(CurrentCharNum + 49)
                    If ChrW(CurrentCharNum) < 5 Then DataToPrint = ChrW(CurrentCharNum + 37) & "(" & ChrW(CurrentCharNum + 49)
               Case 2
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum)
               Case 3
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum)
               Case 4
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum)
               Case 5
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum)
               Case 6
            'Print the center guard pattern after the 6th character
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum) & "*"
               Case 7
            'Add 27 to the ASII value of characters 6-12 to print from character set+ C
            'this is required when printing to the right of the center guard pattern
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
               Case 8
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
               Case 9
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
               Case 10
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
               Case 11
                    DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
               Case 12
            'For the last character print the barcode without the human readable character,
            'the normal guard pattern and then the human readable character.
                    If ChrW(CurrentCharNum) > 4 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 59) & "(" & ChrW(CurrentCharNum + 64)
                    If ChrW(CurrentCharNum) < 5 Then DataToPrint = DataToPrint & ChrW(CurrentCharNum + 59) & "(" & ChrW(CurrentCharNum + 37)
               End Select
          Next I
     End If
     
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          WeightedTotal = 0
          For I = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, I, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then WeightedTotal = WeightedTotal + CurrentCharNum * 3
               If Factor = 1 Then WeightedTotal = WeightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next I
    'Find the CheckDigit by extracting the right-most number from WeightedTotal
          CheckDigit = Val(Right$(WeightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
          
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For I = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
            'EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
                    EANAddOnToPrint = ChrW(43) & EANAddOnToPrint & ChrW(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
     
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For I = 0 To 99 Step 4
               If Val(EAN2AddOn) = I Then Encoding = "AA"
               If Val(EAN2AddOn) = I + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = I + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = I + 3 Then Encoding = "BB"
          Next I
    'Now that we have the total number including the encoding
    'determine what to print
          For I = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
            'EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
                    EANAddOnToPrint = ChrW(43) & EANAddOnToPrint & ChrW(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
     
'Get Printable String
     PrintableString = DataToPrint & EANAddOnToPrint & " "
     
'Return PrintableString
     UPCe = PrintableString
     
End Function

Public Function EAN13(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************

     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
     'Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
DataToEncode = OnlyCorrectData
''
'Remove check digits if they added one
     If Len(OnlyCorrectData) < "12" Then OnlyCorrectData = "0000000000000"
     If Len(OnlyCorrectData) = "16" Then OnlyCorrectData = "0000000000000"
     If Len(OnlyCorrectData) = "13" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 12)
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 12) & Mid(OnlyCorrectData, 14, 2))
     If Len(OnlyCorrectData) > "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 12) & Mid(OnlyCorrectData, 14, 5))

'End sub if incorrect number
     Dim EAN2AddOn As String
     Dim EAN5AddOn As String
     Dim EANAddOnToPrint As String
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 17 Then EAN5AddOn = Mid(OnlyCorrectData, 13, 5)
     If Len(OnlyCorrectData) = 14 Then EAN2AddOn = Mid(OnlyCorrectData, 13, 2)
'split 12 digit number from add-on
     DataToEncode = Mid(OnlyCorrectData, 1, 12)
'<<<< Calculate Check Digit >>>>
     Factor = 3
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the number + WeightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
'Now we must encode the leading digit into the left half of the EAN-13 symbol
'by using variable parity between character sets A and B
     LeadingDigit = Mid(DataToEncode, 1, 1)
     Select Case LeadingDigit
     Case 0
          Encoding = "AAAAAACCCCCC"
     Case 1
          Encoding = "AABABBCCCCCC"
     Case 2
          Encoding = "AABBABCCCCCC"
     Case 3
          Encoding = "AABBBACCCCCC"
     Case 4
          Encoding = "ABAABBCCCCCC"
     Case 5
          Encoding = "ABBAABCCCCCC"
     Case 6
          Encoding = "ABBBAACCCCCC"
     Case 7
          Encoding = "ABABABCCCCCC"
     Case 8
          Encoding = "ABABBACCCCCC"
     Case 9
          Encoding = "ABBABACCCCCC"
     End Select
'add the check digit to the end of the barcode & remove the leading digit
     DataToEncode = Mid(DataToEncode, 2, 11) & CheckDigit
'Now that we have the total number including the check digit, determine character to print
'for proper barcoding:
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get the ASCII value of each number excluding the first number because
    'it is encoded with variable parity
          CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
          CurrentEncoding = Mid(Encoding, I, 1)
    'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
          Select Case CurrentEncoding
          Case "A"
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case "B"
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 17)
          Case "C"
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          End Select
    'add in the 1st character along with guard patterns
          Select Case I
          Case 1
        'For the LeadingDigit print the human readable character,
        'the normal guard pattern and then the rest of the barcode
               If LeadingDigit > 4 Then DataToPrint = ChrW((LeadingDigit + 48) + 64) & "(" & DataToPrint
               If LeadingDigit < 5 Then DataToPrint = ChrW((LeadingDigit + 48) + 37) & "(" & DataToPrint
          Case 6
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & "*"
          Case 12
        'For the last character (12) print the the normal guard pattern
        'after the barcode
               DataToPrint = DataToPrint & "("
          End Select
     Next I
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          WeightedTotal = 0
          For I = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, I, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then WeightedTotal = WeightedTotal + CurrentCharNum * 3
               If Factor = 1 Then WeightedTotal = WeightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next I
    'Find the CheckDigit by extracting the right-most number from WeightedTotal
          CheckDigit = Val(Right$(WeightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For I = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
                    EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
          'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & ChrW(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For I = 0 To 99 Step 4
               If Val(EAN2AddOn) = I Then Encoding = "AA"
               If Val(EAN2AddOn) = I + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = I + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = I + 3 Then Encoding = "BB"
          Next I
    'Now that we have the total number including the encoding
    'determine what to print
          For I = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, I, 1)
               CurrentEncoding = Mid(Encoding, I, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & ChrW(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case I
               Case 1
                    EANAddOnToPrint = ChrW(32) & ChrW(43) & EANAddOnToPrint & ChrW(33)
          'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next I
     End If
'Get Printable String
     PrintableString = DataToPrint & EANAddOnToPrint & " "
'Return PrintableString
     EAN13 = PrintableString
End Function


Public Function EAN8(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
' Enter all the numbers without dashes
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
     
     If Len(OnlyCorrectData) > "7" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 7)
     If Len(OnlyCorrectData) < "7" Then OnlyCorrectData = "0000000"
     DataToEncode = OnlyCorrectData
     'If Len(DataToEncode) <> "7" Then
     '     MsgBox "Cannot process; you MUST enter a 7 digit NUMBER for this type of barcode. Do not use any spaces or dashes."
     '     Exit Function
     'End If
'<<<< Calculate Check Digit >>>>
     Factor = 3
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the number + WeightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
     DataToEncode = DataToEncode & CheckDigit
'Now that have the total number including the check digit, determine character to print
'for proper barcoding
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get the ASCII value of each number
          CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
          CurrentEncoding = Mid(Encoding, I, 1)
    'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
    'Print different barcodes according to the location of the CurrentChar
          Select Case I
          Case 1
        'For the first character print the normal guard pattern
        'and then the barcode without the human readable character
               DataToPrint = "(" & ChrW(CurrentCharNum)
          Case 2
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 3
               DataToPrint = DataToPrint & ChrW(CurrentCharNum)
          Case 4
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & ChrW(CurrentCharNum) & "*"
          Case 5
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 6
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 7
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27)
          Case 8
        'Print the check digit as 8th character + normal guard pattern
               DataToPrint = DataToPrint & ChrW(CurrentCharNum + 27) & "("
          End Select
     Next I
'Get Printable String
     PrintableString = DataToPrint & " "
'Display PrintableString in textbox
     EAN8 = PrintableString
End Function

Public Function SSCC18(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
'
' To create more complex UCC/EAN128 barcodes, call Code128() with the appropriate
' ASCII 0202 and AIs included as documented at:
' http://www.idautomation.com/code128faq.html#EAN128andUCC128
'
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) = "18" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 17))
     If Len(OnlyCorrectData) = "19" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
     If Len(OnlyCorrectData) = "20" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
     If Len(OnlyCorrectData) = "21" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "17" Then OnlyCorrectData = "0000000000000"
     DataToEncode = OnlyCorrectData
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     WeightedTotal = 0
     StringLength = Len(DataToEncode)
     For I = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
'Add check digit and Application Identifier (AI) to DataToEncode
'AI = 00 for SSCC18
'DataToEncode = "00" & DataToEncode & CheckDigit
'Now that we have calculated the MOD 10 for the data, send the string
'to the UCC128() funtion. This function will:
' - Add in the Start C and FNC1 required by UCC/EAN
' - Calculate the MOD 103 required by UCC/EAN
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then SSCC18 = UCC128("00" & DataToEncode & CheckDigit)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then SSCC18 = "(00) " & Mid(DataToEncode, 1, 1) & " " & Mid(DataToEncode, 2, 7) & " " & Mid(DataToEncode, 9, 9) & " " & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then SSCC18 = Str(CheckDigit)
End Function


Public Function SCC14(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
'
' To create more complex UCC/EAN128 barcodes, call Code128() with the appropriate
' ASCII 0202 and AIs included as documented at:
' http://www.idautomation.com/code128faq.html#EAN128andUCC128
'
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 13))
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
     If Len(OnlyCorrectData) = "16" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "13" Then OnlyCorrectData = "0000000000000"
     DataToEncode = OnlyCorrectData
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     WeightedTotal = 0
     StringLength = Len(DataToEncode)
     For I = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          WeightedTotal = WeightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next I
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
'Add check digit and Application Identifier (AI) to DataToEncode
'AI = 00 for SSCC18
'DataToEncode = "00" & DataToEncode & CheckDigit
'Now that we have calculated the MOD 10 for the data, send the string
'to the UCC128() funtion. This function will:
' - Add in the Start C and FNC1 required by UCC/EAN
' - Calculate the MOD 103 required by UCC/EAN
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then SCC14 = UCC128("01" & DataToEncode & CheckDigit)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then SCC14 = "(01) " & Mid(DataToEncode, 1, 1) & " " & Mid(DataToEncode, 2, 7) & " " & Mid(DataToEncode, 9, 5) & " " & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then SCC14 = Str(CheckDigit)
End Function


Public Function UCC128(UCCToEncode As String) As String
     UCCToEncode = ChrW(202) & UCCToEncode
     UCCToEncode = Code128(UCCToEncode, 0, True)
     UCC128 = UCCToEncode
End Function


Public Function Code11(DataToEncode As String) As String
' Copyright © 2000-2003 IDautomation.com, Inc.
' For more info visit http://www.IDAutomation.com
'
' You may use our source code in your applications only if you are using barcode fonts
' created by IDautomation.com, Inc. and you do not remove the copyright notices in the source code.
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric or a dash and remove all others.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "-" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     Factor = 1
     WeightedTotal = 0
     For I = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentChar = Mid(DataToEncode, I, 1)
    'Set the "-" character to the value of 10
          If CurrentChar = "-" Then CurrentChar = "10"
    'multiply by the weighting character and add together
          WeightedTotal = WeightedTotal + (Val(CurrentChar) * Factor)
    'change factor for next calculation
          Factor = Factor + 1
     Next I
'Find the Modulo 11 check digit
     CheckDigit = (WeightedTotal Mod 11)
'Get Printable String
     PrintableString = "(" & DataToEncode & CheckDigit & ")" & " "
'Return the PrintableString
     Code11 = PrintableString
End Function


Public Function RM4SCC(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToEncode = UCase(DataToEncode)
'only pass correct data
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (AscW(Mid(DataToEncode, I, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
     DataToPrint = DataToEncode
     
     Dim r As Integer
     Dim c As Integer
     Dim Rtotal As Long
     Dim Ctotal As Long
     Rtotal = 0
     Ctotal = 0
     WeightedTotal = 0
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get each character one at a time
          CurrentChar = Mid(DataToEncode, I, 1)
    'Get the values of CurrentChar
          Select Case CurrentChar
          Case "0"
               r = 1
               c = 1
          Case "1"
               r = 1
               c = 2
          Case "2"
               r = 1
               c = 3
          Case "3"
               r = 1
               c = 4
          Case "4"
               r = 1
               c = 5
          Case "5"
               r = 1
               c = 0
          Case "6"
               r = 2
               c = 1
          Case "7"
               r = 2
               c = 2
          Case "8"
               r = 2
               c = 3
          Case "9"
               r = 2
               c = 4
          Case "A"
               r = 2
               c = 5
          Case "B"
               r = 2
               c = 0
          Case "C"
               r = 3
               c = 1
          Case "D"
               r = 3
               c = 2
          Case "E"
               r = 3
               c = 3
          Case "F"
               r = 3
               c = 4
          Case "G"
               r = 3
               c = 5
          Case "H"
               r = 3
               c = 0
          Case "I"
               r = 4
               c = 1
          Case "J"
               r = 4
               c = 2
          Case "K"
               r = 4
               c = 3
          Case "L"
               r = 4
               c = 4
          Case "M"
               r = 4
               c = 5
          Case "N"
               r = 4
               c = 0
          Case "O"
               r = 5
               c = 1
          Case "P"
               r = 5
               c = 2
          Case "Q"
               r = 5
               c = 3
          Case "R"
               r = 5
               c = 4
          Case "S"
               r = 5
               c = 5
          Case "T"
               r = 5
               c = 0
          Case "U"
               r = 0
               c = 1
          Case "V"
               r = 0
               c = 2
          Case "W"
               r = 0
               c = 3
          Case "X"
               r = 0
               c = 4
          Case "Y"
               r = 0
               c = 5
          Case "Z"
               r = 0
               c = 0
               
          End Select
    'add the values together
          Rtotal = Rtotal + r
          Ctotal = Ctotal + c
     Next I
     
'divide the Totals by 6 and get the remainder, this is a reference
'to the Check Digit.
'set check digit to CurrentChar (a string)
     Rtotal = (Rtotal Mod 6)
     Ctotal = (Ctotal Mod 6)
     Select Case Rtotal
     Case 1
          Select Case Ctotal
          Case 1
               CurrentChar = "0"
          Case 2
               CurrentChar = "1"
          Case 3
               CurrentChar = "2"
          Case 4
               CurrentChar = "3"
          Case 5
               CurrentChar = "4"
          Case 0
               CurrentChar = "5"
          End Select
     Case 2
          Select Case Ctotal
          Case 1
               CurrentChar = "6"
          Case 2
               CurrentChar = "7"
          Case 3
               CurrentChar = "8"
          Case 4
               CurrentChar = "9"
          Case 5
               CurrentChar = "A"
          Case 0
               CurrentChar = "B"
          End Select
     Case 3
          Select Case Ctotal
          Case 1
               CurrentChar = "C"
          Case 2
               CurrentChar = "D"
          Case 3
               CurrentChar = "E"
          Case 4
               CurrentChar = "F"
          Case 5
               CurrentChar = "G"
          Case 0
               CurrentChar = "H"
          End Select
     Case 4
          Select Case Ctotal
          Case 1
               CurrentChar = "I"
          Case 2
               CurrentChar = "J"
          Case 3
               CurrentChar = "K"
          Case 4
               CurrentChar = "L"
          Case 5
               CurrentChar = "M"
          Case 0
               CurrentChar = "N"
          End Select
     Case 5
          Select Case Ctotal
          Case 1
               CurrentChar = "O"
          Case 2
               CurrentChar = "P"
          Case 3
               CurrentChar = "Q"
          Case 4
               CurrentChar = "R"
          Case 5
               CurrentChar = "S"
          Case 0
               CurrentChar = "T"
          End Select
     Case 0
          Select Case Ctotal
          Case 1
               CurrentChar = "U"
          Case 2
               CurrentChar = "V"
          Case 3
               CurrentChar = "W"
          Case 4
               CurrentChar = "X"
          Case 5
               CurrentChar = "Y"
          Case 0
               CurrentChar = "Z"
          End Select
     End Select
'Get Printable String
     PrintableString = "(" & DataToPrint & CurrentChar & ")" & " "
'Return PrintableString
     RM4SCC = PrintableString
End Function


Public Function Codabar(DataToEncode As String) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
     
' Check to make sure data is numeric, $, +, -, /, or :, and remove all others.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "$" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "+" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "-" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "/" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = "." Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
          If Mid(DataToEncode, I, 1) = ":" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToPrint = OnlyCorrectData
'Get Printable String
     PrintableString = "A" & DataToPrint & "B" & " "
'Return PrintableString
     Codabar = PrintableString
End Function

Public Function Postnet(DataToEncode As String, Optional ReturnType As Integer) As String
'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************
     'Additional logic needed in case ReturnType is not entered
     If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 Then ReturnType = 0
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, I, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, I, 1)
     Next I
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     WeightedTotal = 0
     StringLength = Len(DataToEncode)
     For I = 1 To StringLength
    'Get the value of each number
          CurrentCharNum = Mid(DataToEncode, I, 1)
    'add the values together
          WeightedTotal = WeightedTotal + CurrentCharNum
     Next I
'Find the CheckDigit by finding the number + WeightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     I = (WeightedTotal Mod 10)
     If I <> 0 Then
          CheckDigit = (10 - I)
     Else
          CheckDigit = 0
     End If
'Get Printable String
     DataToPrint = DataToEncode
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then Postnet = "(" & DataToPrint & CheckDigit & ")" & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then Postnet = DataToPrint & CheckDigit
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then Postnet = Str$(CheckDigit)
End Function

Public Function MOD10(M10NumberData As String) As Integer
' This is a general MOD10 function like the one required for EAN and UPC
'*********************************************************************
     Dim M10StringLength As Integer
     Dim M10OnlyCorrectData As String
     Dim M10Factor As Integer
     Dim M10WeightedTotal As Integer
     Dim M10CheckDigit As Integer
     Dim M10I As Integer
     
     M10OnlyCorrectData = ""
     M10StringLength = Len(M10NumberData)
     For M10I = 1 To M10StringLength
        'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(M10NumberData, M10I, 1)) Then M10OnlyCorrectData = M10OnlyCorrectData & Mid(M10NumberData, M10I, 1)
     Next M10I
    '<<<< Generate MOD 10 check digit >>>>
     M10Factor = 3
     M10WeightedTotal = 0
     M10StringLength = Len(M10NumberData)
     For M10I = M10StringLength To 1 Step -1
    'Get the value of each number starting at the end
    'CurrentCharNum = Mid(M10NumberData, I, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          M10WeightedTotal = M10WeightedTotal + (Val(Mid(M10NumberData, M10I, 1)) * M10Factor)
    'change factor for next calculation
          M10Factor = 4 - M10Factor
     Next M10I
    'Find the CheckDigit by finding the smallest number that = a multiple of 10
     M10I = (M10WeightedTotal Mod 10)
     If M10I <> 0 Then
          M10CheckDigit = (10 - M10I)
     Else
          M10CheckDigit = 0
     End If
     MOD10 = Str(M10CheckDigit)
End Function


'*********************************************************************
'*  Visual Basic / VBA Functions for Bar Code Fonts 5.01
'*  Copyright, IDAutomation.com, Inc. 2000-2005. All rights reserved.
'*
'*  Visit http://www.idautomation.com/fonts/tools/vba/ for more
'*  information about the functions in this file.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and this text and the copyright notices
'*  are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************





