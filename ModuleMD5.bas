Attribute VB_Name = "ModuleMD5"
Option Explicit

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Public Function Md5_String_Calc(SourceString As String) As String
    MD5Init
    MD5Update LenB(StrConv(SourceString, vbFromUnicode)), StringToArray(SourceString)
    MD5Final
    Md5_String_Calc = GetValues
End Function

Public Function XMD5(sStr As String) As String      'Simple XMD5 calculate. Ver 0.0.1 (C) 2011 MaxXSoft.
    Dim tmp As String
    Dim intc As Integer
    tmp = sStr
    For intc = 0 To 20
        tmp = Md5_String_Calc(tmp)
        tmp = StrReverse(tmp)
        tmp = tmp & tmp
    Next intc
    tmp = StrReverse(tmp)
    tmp = Md5_String_Calc(tmp)
    For intc = 1 To Abs(Len(tmp) - Len(sStr)) + 1
        tmp = Md5_String_Calc(Md5_String_Calc(tmp) & tmp & Md5_String_Calc(sStr) & sStr)
    Next intc
    XMD5 = tmp
End Function

Private Function StringToArray(InString As String) As Byte()
    Dim i As Integer, bytBuffer() As Byte
    ReDim bytBuffer(LenB(StrConv(InString, vbFromUnicode)))
    bytBuffer = StrConv(InString, vbFromUnicode)
    StringToArray = bytBuffer
End Function

Public Function GetValues() As String
    GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function

Private Function LongToString(Num As Long) As String
        Dim a As Byte, B As Byte, C As Byte, d As Byte
        a = Num And &HFF&
        If a < 16 Then LongToString = "0" & Hex(a) Else LongToString = Hex(a)
        B = (Num And &HFF00&) \ 256
        If B < 16 Then LongToString = LongToString & "0" & Hex(B) Else LongToString = LongToString & Hex(B)
        C = (Num And &HFF0000) \ 65536
        If C < 16 Then LongToString = LongToString & "0" & Hex(C) Else LongToString = LongToString & Hex(C)
        If Num < 0 Then d = ((Num And &H7F000000) \ 16777216) Or &H80& Else d = (Num And &HFF000000) \ 16777216
        If d < 16 Then LongToString = LongToString & "0" & Hex(d) Else LongToString = LongToString & Hex(d)
End Function

Public Sub MD5Init()
    ByteCounter = 0
    State(1) = UnsignedToLong(1732584193#)
    State(2) = UnsignedToLong(4023233417#)
    State(3) = UnsignedToLong(2562383102#)
    State(4) = UnsignedToLong(271733878#)
End Sub

Public Sub MD5Final()
    Dim dblBits As Double, padding(72) As Byte, lngBytesBuffered As Long
    padding(0) = &H80
    dblBits = ByteCounter * 8
    lngBytesBuffered = ByteCounter Mod 64
    If lngBytesBuffered <= 56 Then MD5Update 56 - lngBytesBuffered, padding Else MD5Update 120 - ByteCounter, padding
    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0
    MD5Update 8, padding
End Sub

Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
    Dim II As Integer, i As Integer, j As Integer, K As Integer, lngBufferedBytes As Long, lngBufferRemaining As Long, lngRem As Long

    lngBufferedBytes = ByteCounter Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    ByteCounter = ByteCounter + InputLen

    If InputLen >= lngBufferRemaining Then
        For II = 0 To lngBufferRemaining - 1
            ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
        Next II
        MD5Transform ByteBuffer
        lngRem = (InputLen) Mod 64
        For i = lngBufferRemaining To InputLen - II - lngRem Step 64
            For j = 0 To 63
                ByteBuffer(j) = InputBuffer(i + j)
            Next j
            MD5Transform ByteBuffer
        Next i
        lngBufferedBytes = 0
    Else
      i = 0
    End If
    For K = 0 To InputLen - i - 1
        ByteBuffer(lngBufferedBytes + K) = InputBuffer(i + K)
    Next K
End Sub

Private Sub MD5Transform(Buffer() As Byte)
    Dim x(16) As Long, a As Long, B As Long, C As Long, d As Long
    
    a = State(1)
    B = State(2)
    C = State(3)
    d = State(4)
    Decode 64, x, Buffer
    FF a, B, C, d, x(0), S11, -680876936
    FF d, a, B, C, x(1), S12, -389564586
    FF C, d, a, B, x(2), S13, 606105819
    FF B, C, d, a, x(3), S14, -1044525330
    FF a, B, C, d, x(4), S11, -176418897
    FF d, a, B, C, x(5), S12, 1200080426
    FF C, d, a, B, x(6), S13, -1473231341
    FF B, C, d, a, x(7), S14, -45705983
    FF a, B, C, d, x(8), S11, 1770035416
    FF d, a, B, C, x(9), S12, -1958414417
    FF C, d, a, B, x(10), S13, -42063
    FF B, C, d, a, x(11), S14, -1990404162
    FF a, B, C, d, x(12), S11, 1804603682
    FF d, a, B, C, x(13), S12, -40341101
    FF C, d, a, B, x(14), S13, -1502002290
    FF B, C, d, a, x(15), S14, 1236535329

    GG a, B, C, d, x(1), S21, -165796510
    GG d, a, B, C, x(6), S22, -1069501632
    GG C, d, a, B, x(11), S23, 643717713
    GG B, C, d, a, x(0), S24, -373897302
    GG a, B, C, d, x(5), S21, -701558691
    GG d, a, B, C, x(10), S22, 38016083
    GG C, d, a, B, x(15), S23, -660478335
    GG B, C, d, a, x(4), S24, -405537848
    GG a, B, C, d, x(9), S21, 568446438
    GG d, a, B, C, x(14), S22, -1019803690
    GG C, d, a, B, x(3), S23, -187363961
    GG B, C, d, a, x(8), S24, 1163531501
    GG a, B, C, d, x(13), S21, -1444681467
    GG d, a, B, C, x(2), S22, -51403784
    GG C, d, a, B, x(7), S23, 1735328473
    GG B, C, d, a, x(12), S24, -1926607734

    HH a, B, C, d, x(5), S31, -378558
    HH d, a, B, C, x(8), S32, -2022574463
    HH C, d, a, B, x(11), S33, 1839030562
    HH B, C, d, a, x(14), S34, -35309556
    HH a, B, C, d, x(1), S31, -1530992060
    HH d, a, B, C, x(4), S32, 1272893353
    HH C, d, a, B, x(7), S33, -155497632
    HH B, C, d, a, x(10), S34, -1094730640
    HH a, B, C, d, x(13), S31, 681279174
    HH d, a, B, C, x(0), S32, -358537222
    HH C, d, a, B, x(3), S33, -722521979
    HH B, C, d, a, x(6), S34, 76029189
    HH a, B, C, d, x(9), S31, -640364487
    HH d, a, B, C, x(12), S32, -421815835
    HH C, d, a, B, x(15), S33, 530742520
    HH B, C, d, a, x(2), S34, -995338651

    II a, B, C, d, x(0), S41, -198630844
    II d, a, B, C, x(7), S42, 1126891415
    II C, d, a, B, x(14), S43, -1416354905
    II B, C, d, a, x(5), S44, -57434055
    II a, B, C, d, x(12), S41, 1700485571
    II d, a, B, C, x(3), S42, -1894986606
    II C, d, a, B, x(10), S43, -1051523
    II B, C, d, a, x(1), S44, -2054922799
    II a, B, C, d, x(8), S41, 1873313359
    II d, a, B, C, x(15), S42, -30611744
    II C, d, a, B, x(6), S43, -1560198380
    II B, C, d, a, x(13), S44, 1309151649
    II a, B, C, d, x(4), S41, -145523070
    II d, a, B, C, x(11), S42, -1120210379
    II C, d, a, B, x(2), S43, 718787259
    II B, C, d, a, x(9), S44, -343485551

    State(1) = LongOverflowAdd(State(1), a)
    State(2) = LongOverflowAdd(State(2), B)
    State(3) = LongOverflowAdd(State(3), C)
    State(4) = LongOverflowAdd(State(4), d)
End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
    Dim intDblIndex As Integer, intByteIndex As Integer, dblSum As Double
    For intByteIndex = 0 To Length - 1 Step 4
        dblSum = InputBuffer(intByteIndex) + InputBuffer(intByteIndex + 1) * 256# + InputBuffer(intByteIndex + 2) * 65536# + InputBuffer(intByteIndex + 3) * 16777216#
        OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
        intDblIndex = intDblIndex + 1
    Next intByteIndex
End Sub

Private Function FF(a As Long, B As Long, C As Long, d As Long, x As Long, S As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, (B And C) Or (Not (B) And d), x, ac)
    a = LongLeftRotate(a, S)
    a = LongOverflowAdd(a, B)
End Function

Private Function GG(a As Long, B As Long, C As Long, d As Long, x As Long, S As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, (B And d) Or (C And Not (d)), x, ac)
    a = LongLeftRotate(a, S)
    a = LongOverflowAdd(a, B)
End Function

Private Function HH(a As Long, B As Long, C As Long, d As Long, x As Long, S As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, B Xor C Xor d, x, ac)
    a = LongLeftRotate(a, S)
    a = LongOverflowAdd(a, B)
End Function

Private Function II(a As Long, B As Long, C As Long, d As Long, x As Long, S As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, C Xor (B Or Not (d)), x, ac)
    a = LongLeftRotate(a, S)
    a = LongOverflowAdd(a, B)
End Function

Function LongLeftRotate(value As Long, Bits As Long) As Long
    Dim lngSign As Long, lngI As Long
    Bits = Bits Mod 32
    If Bits = 0 Then LongLeftRotate = value: Exit Function
    For lngI = 1 To Bits
        lngSign = value And &HC0000000
        value = (value And &H3FFFFFFF) * 2
        value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    LongLeftRotate = value
End Function

Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
    Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
    Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function UnsignedToLong(value As Double) As Long
    If value < 0 Or value >= OFFSET_4 Then Error 6
    If value <= MAXINT_4 Then UnsignedToLong = value Else UnsignedToLong = value - OFFSET_4
End Function

Private Function LongToUnsigned(value As Long) As Double
    If value < 0 Then LongToUnsigned = value + OFFSET_4 Else LongToUnsigned = value
End Function
