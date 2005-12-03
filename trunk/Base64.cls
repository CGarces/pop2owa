VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
'Base 64 Encoding class
'
'
'@author Wil Johnson (Wil.Johnson@att.net)
'@version 1.1
'@date 21/03/2000
'@remarks   This code is for example purposes only, and is provided as-is.  While it has
'   worked well under limited testing, the current error handling is minimal and
'   should be expanded upon before release into a production environment.  Please
'   report all bugs found to the author for correction, even if you have already
'   corrected them yourself.
'
'   Again, this code is a rough draft.  Feel free to use it, but do so at your own
'   risk.  These release notes must also remain intact.


Option Explicit

Private m_bytIndex(0 To 63) As Byte
Private m_bytReverseIndex(0 To 255) As Byte

Private Const k_bytEqualSign As Byte = 61

Private Const k_bytMask1 As Byte = 3      '00000011
Private Const k_bytMask2 As Byte = 15     '00001111
Private Const k_bytMask3 As Byte = 63     '00111111

Private Const k_bytMask4 As Byte = 192    '11000000
Private Const k_bytMask5 As Byte = 240    '11110000
Private Const k_bytMask6 As Byte = 252    '11111100



Private Const k_bytShift2 As Byte = 4
Private Const k_bytShift4 As Byte = 16
Private Const k_bytShift6 As Byte = 64


Private Const k_lMaxBytesPerLine As Long = 152

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    


Public Function Encode(ByRef sInput As String) As String
    If sInput = "" Then Exit Function
    Dim bytTemp() As Byte
    bytTemp = StrConv(sInput, vbFromUnicode)
    Encode = EncodeArr(bytTemp)
End Function

Public Function EncodeFromFile(sFileName As String) As String
    On Error GoTo ErrorHandler:
    Dim bytFile() As Byte
    Dim iFile As Integer
    
    'get new file handle
    iFile = FreeFile

    Open sFileName For Input As #iFile
    'size the array to the size of the file
    ReDim bytFile(0 To VBA.LOF(iFile) - 1) As Byte
    'get everything in the file
    Input #iFile, bytFile
    Close #iFile
    
    'encode it
    EncodeFromFile = EncodeArr(bytFile)
    
    GoTo Done:
    
ErrorHandler:
    EncodeFromFile = ""
    Resume Done:

Done:
    On Error Resume Next
    Close #iFile
    
End Function

Public Function EncodeArr(ByRef bytInput() As Byte) As String
    On Error GoTo ErrorHandler:
    Dim bytWorkspace() As Byte      'array for the "rough draft" of the encoded data
    Dim bytResult() As Byte         'array for the "final draft"
    Dim bytCrLf(0 To 3) As Byte     'array that will contain vbCrLf, for CopyMemory purposes
    
    Dim lCounter As Long            'counter used to iterate through input bytes
    Dim lWorkspaceCounter As Long   'counter used to iterate through workspace bytes
    Dim lLineCounter As Long        'counter used when inserting CrLfs
    Dim lCompleteLines As Long      'used for calculations when inserting CrLfs
    Dim lBytesRemaining As Long     'used to determine how much work is left after coming out of a loop

    'pointers
    Dim lpWorkSpace As Long         'pointer to bytWorkspace.  it's offset will change as bytes are copied out of the array
    Dim lpResult As Long            'pointer to bytResult.  it's offset will also change
    Dim lpCrLf As Long              'pointer to bytCrLf.  it is not offset and will not change
    

    'create a workspace larger than we need
    'this is to prevent VB from having to allocate memory constantly
    If UBound(bytInput) < 1024 Then
        ReDim bytWorkspace(LBound(bytInput) To (LBound(bytInput) + 4096)) As Byte
    Else
        ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 4)) As Byte
    End If

    
    lWorkspaceCounter = LBound(bytWorkspace)

    'step through in 3 byte increments
    For lCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 3) + 3)) Step 3
        'result set byte 1 = 6 most significant bits of first byte of input set
        'bits are right shifted by 2
        bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
        
        'result set byte 2 = 2 least significant bits of first byte and 4 most significant bits of second byte of input set
        'bits from first byte are left shifted by 4
        'bits from second byte are right shifted by 4
        bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
        
        'result set byte 3 = 4 least significant bits of second byte and 2 most significant bits of third byte of input set
        'bits from second byte are left shifted by 2
        'bits from third byte are right shifted by 6
        bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + (bytInput(lCounter + 2) \ k_bytShift6))
        
        'result set byte 4 = 6 least significant bits of third byte of input set
        'bits from third byte are not shifted at all
        bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
        lWorkspaceCounter = lWorkspaceCounter + 8
    Next lCounter

    
    Select Case (UBound(bytInput) Mod 3):
        'for information on how bits are masked and shifted, see above
        Case 0:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex((bytInput(lCounter) And k_bytMask1) * k_bytShift4)
            bytWorkspace(lWorkspaceCounter + 4) = k_bytEqualSign
            bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign
           
        Case 1:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
            bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2)
            bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign

        Case 2:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
            bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + ((bytInput(lCounter + 2)) \ k_bytShift6))
            bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
        
    End Select

    lWorkspaceCounter = lWorkspaceCounter + 8

    'base64 encoding allows no more than 76 characters per line,
    'which translates to 152 bytes since the string is unicode
    If lWorkspaceCounter <= k_lMaxBytesPerLine Then
        'no need to line wrap.
        EncodeArr = Left$(bytWorkspace, InStr(1, bytWorkspace, Chr$(0)) - 1)
        'EncodeArr = bytWorkspace
    Else
        'must wrap lines
        'first, populate the CrLf byte array
        bytCrLf(0) = 13
        bytCrLf(1) = 0
        bytCrLf(2) = 10
        bytCrLf(3) = 0
                
        'size the end result array
        ReDim bytResult(LBound(bytWorkspace) To UBound(bytWorkspace))
        
        'get pointers to the various arrays
        lpWorkSpace = VarPtr(bytWorkspace(LBound(bytWorkspace)))
        lpResult = VarPtr(bytResult(LBound(bytResult)))
        lpCrLf = VarPtr(bytCrLf(LBound(bytCrLf)))
        
        'get count of complete lines
        lCompleteLines = Fix(lWorkspaceCounter / k_lMaxBytesPerLine)
        
        For lLineCounter = 0 To lCompleteLines
            'copy first line
            CopyMemory lpResult, lpWorkSpace, k_lMaxBytesPerLine
            
            'offset the workspace and result pointers by k_lMaxBytesPerLine
            lpWorkSpace = lpWorkSpace + k_lMaxBytesPerLine
            lpResult = lpResult + k_lMaxBytesPerLine
            
            'copy CrLf to result
            CopyMemory lpResult, lpCrLf, 4&
            
            'offset result pointer by another 4 bytes to account for the CrLf
            lpResult = lpResult + 4&
        Next lLineCounter
        
        'check if there are any remaining bytes in an incomplete line to be copied
        lBytesRemaining = lWorkspaceCounter - (lCompleteLines * k_lMaxBytesPerLine)
        If lBytesRemaining > 0 Then
            'copy remaining bytes to result
            CopyMemory lpResult, lpWorkSpace, lBytesRemaining
        End If
        
        'no need to resize the result before passing it back to a string,
        'since the empty space is made up of null chars that will terminate the
        'string automatically.
        'CopyMemory StrPtr(EncodeArr), VarPtr(bytResult(LBound(bytResult))), lpResult + lBytesRemaining
        EncodeArr = Left$(bytResult, InStr(1, bytResult, Chr$(0)) - 1)
    End If
    
    Exit Function

ErrorHandler:
    'on error just return an empty array
    Erase bytResult
    EncodeArr = bytResult
End Function

Public Function Decode(sInput As String) As String
    If sInput = "" Then Exit Function
    Decode = StrConv(DecodeArr(sInput), vbUnicode)
End Function

Public Sub DecodeToFile(sInput As String, sFileName As String)
    On Error GoTo ErrorHandler:
    Dim iFile As Integer
    
    'do not overwrite existing files
    If Dir(sFileName) <> "" Then
        Err.Raise vbObjectError + 1000, "Base64.DecodeToFile", "File already exists."
        GoTo Done:
    End If
    
    iFile = FreeFile
    Open sFileName For Binary As #iFile
    Put #iFile, , DecodeArr(sInput)
    Close #iFile
    
    GoTo Done
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Done:
    
Done:
    On Error Resume Next
    Close #iFile
End Sub

Public Function DecodeArr(sInput As String) As Byte()
    'returns a SBCS byte array
    Dim bytInput() As Byte          'base64 encoded string to work with
    Dim bytWorkspace() As Byte      'byte array to use as workspace
    Dim bytResult() As Byte         'array that result will be copied to
    Dim lInputCounter As Long       'iteration counter for input array
    Dim lWorkspaceCounter As Long   'iteration counter for workspace array
    
    
    'get rid of CrLfs, and "="s since they're not required for decoding,
    'and place the input in the byte array
    bytInput = Replace(Replace(sInput, vbCrLf, ""), "=", "")
    
    'size the workspace
    ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 2)) As Byte
    lWorkspaceCounter = LBound(bytWorkspace)
    
    'pass bytes back through index to get original values
    For lInputCounter = LBound(bytInput) To UBound(bytInput)
        bytInput(lInputCounter) = m_bytReverseIndex(bytInput(lInputCounter))
    Next lInputCounter
    
    For lInputCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 8) + 8)) Step 8
        'left shift first input byte by 2 and right shift second input byte by 4
        bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
        
        'mask bits 5-8 of second byte, left shift it by 4
        'right shift third byte by 2, add it to result of second byte
        bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + _
                                              (bytInput(lInputCounter + 4) \ k_bytShift2)
        
        'mask bits 3-8 of third byte, left shift it by 6, add it to fourth byte
        bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + _
                                              bytInput(lInputCounter + 6)
                                              
        lWorkspaceCounter = lWorkspaceCounter + 3
    Next lInputCounter
    
    
    'decode any remaining bytes that are not part of a full 4 byte block
    Select Case (UBound(bytInput) Mod 8):
        Case 3:
            'left shift first input byte by 2 and right shift second input byte by 4
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            
        Case 5:
            'left shift first input byte by 2 and right shift second input byte by 4
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            
            'mask bits 5-8 of second byte, left shift it by 4
            'right shift third byte by 2, add it to result of second byte
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + _
                                                  (bytInput(lInputCounter + 4) \ k_bytShift2)
            lWorkspaceCounter = lWorkspaceCounter + 1
            
        Case 7:
            'left shift first input byte by 2 and right shift second input byte by 4
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            
            'mask bits 5-8 of second byte, left shift it by 4
            'right shift third byte by 2, add it to result of second byte
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + _
                                                  (bytInput(lInputCounter + 4) \ k_bytShift2)
            
            'mask bits 3-8 of third byte, left shift it by 6, add it to fourth byte
            bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + _
                                                  bytInput(lInputCounter + 6)
            lWorkspaceCounter = lWorkspaceCounter + 2
    
        
    End Select
    
    'size the result array
    ReDim bytResult(LBound(bytWorkspace) To lWorkspaceCounter) As Byte
    
    'if option base is set to 1 then don't increment this value
    If LBound(bytWorkspace) = 0 Then
        lWorkspaceCounter = lWorkspaceCounter + 1
    End If
    
    'move decoded data to a properly sized array
    CopyMemory VarPtr(bytResult(LBound(bytResult))), VarPtr(bytWorkspace(LBound(bytWorkspace))), lWorkspaceCounter
    'return
    DecodeArr = bytResult
End Function


Private Sub Class_Initialize()
    m_bytIndex(0) = 65 'Asc("A")
    m_bytIndex(1) = 66 'Asc("B")
    m_bytIndex(2) = 67 'Asc("C")
    m_bytIndex(3) = 68 'Asc("D")
    m_bytIndex(4) = 69 'Asc("E")
    m_bytIndex(5) = 70 'Asc("F")
    m_bytIndex(6) = 71 'Asc("G")
    m_bytIndex(7) = 72 'Asc("H")
    m_bytIndex(8) = 73 'Asc("I")
    m_bytIndex(9) = 74 'Asc("J")
    m_bytIndex(10) = 75 'Asc("K")
    m_bytIndex(11) = 76 'Asc("L")
    m_bytIndex(12) = 77 'Asc("M")
    m_bytIndex(13) = 78 'Asc("N")
    m_bytIndex(14) = 79 'Asc("O")
    m_bytIndex(15) = 80 'Asc("P")
    m_bytIndex(16) = 81 'Asc("Q")
    m_bytIndex(17) = 82 'Asc("R")
    m_bytIndex(18) = 83 'Asc("S")
    m_bytIndex(19) = 84 'Asc("T")
    m_bytIndex(20) = 85 'Asc("U")
    m_bytIndex(21) = 86 'Asc("V")
    m_bytIndex(22) = 87 'Asc("W")
    m_bytIndex(23) = 88 'Asc("X")
    m_bytIndex(24) = 89 'Asc("Y")
    m_bytIndex(25) = 90 'Asc("Z")
    m_bytIndex(26) = 97 'Asc("a")
    m_bytIndex(27) = 98 'Asc("b")
    m_bytIndex(28) = 99 'Asc("c")
    m_bytIndex(29) = 100 'Asc("d")
    m_bytIndex(30) = 101 'Asc("e")
    m_bytIndex(31) = 102 'Asc("f")
    m_bytIndex(32) = 103 'Asc("g")
    m_bytIndex(33) = 104 'Asc("h")
    m_bytIndex(34) = 105 'Asc("i")
    m_bytIndex(35) = 106 'Asc("j")
    m_bytIndex(36) = 107 'Asc("k")
    m_bytIndex(37) = 108 'Asc("l")
    m_bytIndex(38) = 109 'Asc("m")
    m_bytIndex(39) = 110 'Asc("n")
    m_bytIndex(40) = 111 'Asc("o")
    m_bytIndex(41) = 112 'Asc("p")
    m_bytIndex(42) = 113 'Asc("q")
    m_bytIndex(43) = 114 'Asc("r")
    m_bytIndex(44) = 115 'Asc("s")
    m_bytIndex(45) = 116 'Asc("t")
    m_bytIndex(46) = 117 'Asc("u")
    m_bytIndex(47) = 118 'Asc("v")
    m_bytIndex(48) = 119 'Asc("w")
    m_bytIndex(49) = 120 'Asc("x")
    m_bytIndex(50) = 121 'Asc("y")
    m_bytIndex(51) = 122 'Asc("z")
    m_bytIndex(52) = 48 'Asc("0")
    m_bytIndex(53) = 49 'Asc("1")
    m_bytIndex(54) = 50 'Asc("2")
    m_bytIndex(55) = 51 'Asc("3")
    m_bytIndex(56) = 52 'Asc("4")
    m_bytIndex(57) = 53 'Asc("5")
    m_bytIndex(58) = 54 'Asc("6")
    m_bytIndex(59) = 55 'Asc("7")
    m_bytIndex(60) = 56 'Asc("8")
    m_bytIndex(61) = 57 'Asc("9")
    m_bytIndex(62) = 43 'Asc("+")
    m_bytIndex(63) = 47 'Asc("/")
    
    m_bytReverseIndex(65) = 0 'Asc("A")
    m_bytReverseIndex(66) = 1 'Asc("B")
    m_bytReverseIndex(67) = 2 'Asc("C")
    m_bytReverseIndex(68) = 3 'Asc("D")
    m_bytReverseIndex(69) = 4 'Asc("E")
    m_bytReverseIndex(70) = 5 'Asc("F")
    m_bytReverseIndex(71) = 6 'Asc("G")
    m_bytReverseIndex(72) = 7 'Asc("H")
    m_bytReverseIndex(73) = 8 'Asc("I")
    m_bytReverseIndex(74) = 9 'Asc("J")
    m_bytReverseIndex(75) = 10 'Asc("K")
    m_bytReverseIndex(76) = 11 'Asc("L")
    m_bytReverseIndex(77) = 12 'Asc("M")
    m_bytReverseIndex(78) = 13 'Asc("N")
    m_bytReverseIndex(79) = 14 'Asc("O")
    m_bytReverseIndex(80) = 15 'Asc("P")
    m_bytReverseIndex(81) = 16 'Asc("Q")
    m_bytReverseIndex(82) = 17 'Asc("R")
    m_bytReverseIndex(83) = 18 'Asc("S")
    m_bytReverseIndex(84) = 19 'Asc("T")
    m_bytReverseIndex(85) = 20 'Asc("U")
    m_bytReverseIndex(86) = 21 'Asc("V")
    m_bytReverseIndex(87) = 22 'Asc("W")
    m_bytReverseIndex(88) = 23 'Asc("X")
    m_bytReverseIndex(89) = 24 'Asc("Y")
    m_bytReverseIndex(90) = 25 'Asc("Z")
    m_bytReverseIndex(97) = 26 'Asc("a")
    m_bytReverseIndex(98) = 27 'Asc("b")
    m_bytReverseIndex(99) = 28 'Asc("c")
    m_bytReverseIndex(100) = 29 'Asc("d")
    m_bytReverseIndex(101) = 30 'Asc("e")
    m_bytReverseIndex(102) = 31 'Asc("f")
    m_bytReverseIndex(103) = 32 'Asc("g")
    m_bytReverseIndex(104) = 33 'Asc("h")
    m_bytReverseIndex(105) = 34 'Asc("i")
    m_bytReverseIndex(106) = 35 'Asc("j")
    m_bytReverseIndex(107) = 36 'Asc("k")
    m_bytReverseIndex(108) = 37 'Asc("l")
    m_bytReverseIndex(109) = 38 'Asc("m")
    m_bytReverseIndex(110) = 39 'Asc("n")
    m_bytReverseIndex(111) = 40 'Asc("o")
    m_bytReverseIndex(112) = 41 'Asc("p")
    m_bytReverseIndex(113) = 42 'Asc("q")
    m_bytReverseIndex(114) = 43 'Asc("r")
    m_bytReverseIndex(115) = 44 'Asc("s")
    m_bytReverseIndex(116) = 45 'Asc("t")
    m_bytReverseIndex(117) = 46 'Asc("u")
    m_bytReverseIndex(118) = 47 'Asc("v")
    m_bytReverseIndex(119) = 48 'Asc("w")
    m_bytReverseIndex(120) = 49 'Asc("x")
    m_bytReverseIndex(121) = 50 'Asc("y")
    m_bytReverseIndex(122) = 51 'Asc("z")
    m_bytReverseIndex(48) = 52 'Asc("0")
    m_bytReverseIndex(49) = 53 'Asc("1")
    m_bytReverseIndex(50) = 54 'Asc("2")
    m_bytReverseIndex(51) = 55 'Asc("3")
    m_bytReverseIndex(52) = 56 'Asc("4")
    m_bytReverseIndex(53) = 57 'Asc("5")
    m_bytReverseIndex(54) = 58 'Asc("6")
    m_bytReverseIndex(55) = 59 'Asc("7")
    m_bytReverseIndex(56) = 60 'Asc("8")
    m_bytReverseIndex(57) = 61 'Asc("9")
    m_bytReverseIndex(43) = 62 'Asc("+")
    m_bytReverseIndex(47) = 63 'Asc("/")

End Sub


Public Function QP(ByRef sInput As String) As String
If sInput = "" Then Exit Function
Dim lngChar As Long
Dim intAscii As Integer
Dim lngTotal As Long
Dim strReturn As String
Dim strChar As String
lngTotal = Len(sInput)
For lngChar = 1 To lngTotal
    strChar = Mid(sInput, lngChar, 1)
    intAscii = Asc(strChar)
    If (intAscii >= 33 And intAscii <= 60) Or (intAscii >= 62 And intAscii <= 126) Or intAscii = 9 Or intAscii = 10 Or intAscii = 32 Then
        strReturn = strReturn & Mid(sInput, lngChar, 1)
    Else
        strReturn = strReturn & "=" & Right("0" & Hex$(intAscii), 2)
    End If
Next
QP = strReturn
End Function
