VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBMPparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' reusable UDT for 1 & 2 dim arrays
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long         ' +4 from .biSize
    biHeight As Long        ' +8
    biPlanes As Integer     ' +12
    biBitCount As Integer   ' +14
    biCompression As Long   ' +16
    biSizeImage As Long     ' +20
    biXPelsPerMeter As Long ' +24
    biYPelsPerMeter As Long ' +28
    biClrUsed As Long       ' +32
    biClrImportant As Long  ' 40th byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette As Long
End Type

' used to transfer a stdPicture bmp,jpg,wmf to a DIB
Private Declare Function GetGDIObject Lib "GDI32.DLL" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "USER32.DLL" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "GDI32.DLL" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Function LoadStream(inStream() As Byte, cHost As c32bppDIB, _
                            Optional ByVal streamOffset As Long = 0, _
                            Optional ByVal streamLength As Long = 0) As Boolean

    ' PURPOSE: Test passed stream for being a 32bpp bitmap.
    ' If not 32bpp, then the stream is converted to a stdPicture and the contents
    ' of that object are drawn to the 32bpp.
    
    ' With the exception of wmf, emf & 32bpp. This class does not handle transparency.
    ' Therefore, the stream should have been passed to the png, gif & icon parsers
    ' first.
    
    ' Parameters.
    ' inStream() :: the byte array containing the image
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    '   - If zero, then size is UBound(inStream)-streamOffset+1
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim lValue As Long, iValue As Integer
    Dim X As Long, Y As Long, lScanWidth As Long
    Dim Offset As Long, iBitCount As Integer
    Dim aDIB() As Byte, tSA As SafeArray
    Dim bAlpha As Boolean
    
    ' manually parse the bitmap header.
    ' Why? because VB's LoadPicture will convert the image into a screen
    ' compatible bitmap; where if screen resolution was less than true color,
    ' a 32bpp image would end up being 24bp or less vs 32bpp
    CopyMemory iValue, inStream(streamOffset), 2&   ' get 1st 2 bytes of the stream
    If iValue = &H4D42 Then                         ' is it a bmp magic number
        CopyMemory iBitCount, inStream(streamOffset + 28), 2& ' bit count
        CopyMemory X, inStream(streamOffset + 18), 4& ' width
        CopyMemory Y, inStream(streamOffset + 22), 4& ' height
        
        ' validate size
        ' width must be at least 1 pixel & height must be a least 1 pixel
        If X < 1 Or Y = 0& Then Exit Function ' -Y indicates top down DIB
        
        On Error Resume Next
        CopyMemory Offset, inStream(streamOffset + 10), 4& ' start of image
        ' validate enough bytes exist for the image
        lValue = (streamOffset + streamLength) - (iparseByteAlignOnWord(iBitCount, X) * Abs(Y) + Offset)
        If Err Then     ' should some overflow occur
            Err.Clear
            lValue = -1&
        End If
        If lValue >= 0& Then              ' is array big enough?
            If iBitCount = 32 Then       ' else we will allow VB to convert it for us
                                         ' because it doesn't contain transparency anyway
                CopyMemory lValue, inStream(streamOffset + 30&), 4& ' compression
                If lValue = 0& Then         ' manually handle no-compression bitmaps
                                            ' else allow VB to convert the bitmap
                    cHost.InitializeDIB X, Abs(Y)
                    With tSA
                        .cbElements = 1
                        .cDims = 2
                        .pvData = cHost.BitsPointer
                        .rgSABound(0).cElements = cHost.Height
                        .rgSABound(1).cElements = cHost.scanWidth
                    End With
                    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4&
                    
                    lScanWidth = cHost.scanWidth
                    If Y < 0& Then ' the dib is top down vs bottom up
                        ' flip the DIB
                        Y = -Y
                        For lValue = 0& To Y - 1&
                            ' start of scan line in source image
                            X = lScanWidth * (Y - lValue - 1&) + Offset
                            ' copy to upside down scan line on our DIB
                            CopyMemory aDIB(0&, lValue), inStream(X), lScanWidth
                        Next
                    Else    ' bottom up dib; simply copy bits
                        CopyMemory ByVal cHost.BitsPointer, inStream(streamOffset + Offset), cHost.Height * lScanWidth
                    End If
                    
                    ' see if 32bpp is premulitplied or not
                    iparseValidateAlphaChannel aDIB(), True, bAlpha, lValue
                    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4& ' remove overlay
                    ' set other properties
                    cHost.Alpha = bAlpha
                    cHost.ImageType = lValue
                    LoadStream = True
                End If
            End If
            
        End If
    End If
    On Error GoTo 0
    
    If cHost.Handle = 0& Then ' we didn't process the image above, try VB's LoadPicture
    
        On Error Resume Next
        Dim tPic As StdPicture
        Set tPic = iparseArrayToPicture(inStream(), streamOffset, streamLength)
        If Err Then
            Err.Clear
        Else
            LoadStream = ConvertstdPicTo32bpp(tPic, 0&, cHost, iBitCount)
        End If
    
    End If

End Function

Public Function ConvertstdPicTo32bpp(stdPic As StdPicture, ByVal Handle As Long, cHost As c32bppDIB, ByVal bitCount As Integer) As Boolean

    ' stdPic is passed from cHost.LoadPicture_StdPicture and Handle=0&
    '   in this case stdPic can be a bmp, jpg, gif, wmf, emf
    ' Handle is passed from chost.LoadPicture_ByHandle and stdPic is Nothing
    '   in this case, Handle only references a bitmap

    Dim tSA As SafeArray, tOBJ As BITMAP, tBMPI As BITMAPINFO
    Dim cx As Long, cy As Long
    Dim tDC As Long, bAlpha As Boolean, iType As Long
    Dim aDIB() As Byte
    Dim bmpDC As Long, bmpOld As Long

    If stdPic Is Nothing Then
        If Handle = 0& Then Exit Function ' couldn't convert image
        If GetGDIObject(Handle, Len(tOBJ), tOBJ) = 0& Then Exit Function
        cx = tOBJ.bmWidth
        cy = Abs(tOBJ.bmHeight)
    ElseIf stdPic.Type = vbPicTypeNone Then
        Exit Function
    Else
        ' get the picture's width & height & initialize DIB
        cx = ConvertHimetrix2Pixels(stdPic.Width, True)
        cy = ConvertHimetrix2Pixels(stdPic.Height, False)
        Handle = 0&
    End If
    
    cHost.InitializeDIB cx, cy
    tDC = cHost.LoadDIBinDC(True)
    
    ' WMF/EMFs are kinda weird, but here is a neat trick to determine if it
    ' has transparency. Fill the entire image with white, then when it is
    ' rendered, any "transparent" areas were not drawn over, left the
    ' alpha byte as 255. Those areas that are drawn over are changed to zero.
    If Handle = 0& Then
        If stdPic.Type = vbPicTypeEMetafile Or stdPic.Type = vbPicTypeMetafile Then
            FillMemory ByVal cHost.BitsPointer, cy * cHost.scanWidth, 255
        End If
    
        ' render the stdPic to the host's dc
        stdPic.Render tDC + 0&, 0&, 0&, cx + 0&, cy + 0&, _
            0, stdPic.Height, stdPic.Width, -stdPic.Height, ByVal 0&
    Else
        ' bitmap checks. Here we can process a 32bpp bitmap loaded into a stdPicture
        ' or loaded from LoadImage API with full confidence if carrying over alpha
        If tOBJ.bmBitsPixel = 32 Then       ' if image is 32bpp then
            If tOBJ.bmBits = 0& Then        ' allow GetDIBits to transfer for us
                With tBMPI.bmiHeader
                    .biBitCount = 32
                    .biHeight = cy
                    .biWidth = cx
                    .biPlanes = 1
                    .biSize = 40
                End With
                If GetDIBits(tDC, Handle, 0, tBMPI.bmiHeader.biHeight, ByVal cHost.BitsPointer, tBMPI, 0) = 0 Then
                    tOBJ.bmBitsPixel = 0&   ' flag to allow BitBlt when this failed
                End If
            Else                ' if we have a bits pointer, simply copy data
                CopyMemory ByVal cHost.BitsPointer, ByVal tOBJ.bmBits, tOBJ.bmWidthBytes * tOBJ.bmHeight
            End If
        End If
        If Not tOBJ.bmBitsPixel = 32& Then          ' use BitBlt
            bmpDC = CreateCompatibleDC(tDC)         ' create a dc & blt, image is not 32bpp
            bmpOld = SelectObject(bmpDC, Handle)
            BitBlt tDC, 0, 0, cx, cy, bmpDC, 0, 0, vbSrcCopy
            SelectObject bmpDC, bmpOld
            DeleteDC bmpDC
        End If
    End If
    ' unmanage the DC if needed
    cHost.LoadDIBinDC False
    
    ' map our array to the host's DIB
    With tSA
        .cbElements = 1 ' as byte array
        .cDims = 2      ' as 1 dimensional
        .pvData = cHost.BitsPointer
        .rgSABound(0).cElements = cy
        .rgSABound(1).cElements = cHost.scanWidth
    End With
    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4& ' apply overlay
    
    If Handle = 0& Then  ' processing wmf, emf
        If stdPic.Type = vbPicTypeEMetafile Or stdPic.Type = vbPicTypeMetafile Then
            ' as mentioned above, any transparent pixels will have alpha value = 255
            For cy = 0& To cHost.Height - 1&
                For cx = 3& To cHost.scanWidth - 1& Step 4&
                    If aDIB(cx, cy) = 255 Then    ' 100% transparent
                        CopyMemory aDIB(cx - 3&, cy), 0&, 4& ' make bits transparent
                        bAlpha = True
                    Else                        ' 100% opaque
                        aDIB(cx, cy) = 255 ' was 255, now 0, change back to 255
                    End If
                Next
            Next
        Else    ' jpg or non-alpha bitmap; should no longer get here.
                ' Most recent update should pass non-wmf/emf by Handle vs stdPicture object
            ' validate first that it has no alpha bytes
            If bitCount = 0 Then
                ' when called from cHost.LoadPicture_Resource then no BitCount is known
                Dim tBMP As BITMAP
                GetGDIObject stdPic.Handle, Len(tBMP), tBMP
                bitCount = tBMP.bmBitsPixel
            End If
            ' Note: I have experienced 32bpp & 24bpp stdPicture.Rendering onto a 32bpp DIB
            ' and writing in the alpha channel. These stdPictures did not use an alpha
            ' channel, and this is a bug of some sort with stdPicture or VB not fully
            ' supporting 32bpp DIBs.  The -1& below forces next routine to fill the
            ' alpha channel with 255 cHost.ImageType = vbPicTypeBitmap
            iparseValidateAlphaChannel aDIB(), True, False, -1&
        End If
        iType = stdPic.Type
    ElseIf tOBJ.bmBitsPixel = 32& Then
        iparseValidateAlphaChannel aDIB(), True, bAlpha, iType
    Else
        iparseValidateAlphaChannel aDIB(), True, False, -1&
        iType = imgBitmap
    End If
    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4&  ' remove overlay
    cHost.Alpha = bAlpha
    cHost.ImageType = iType
    ConvertstdPicTo32bpp = True

End Function

Private Function ConvertHimetrix2Pixels(vHiMetrix As Long, Horizontally As Boolean) As Long
    ' conversion from Himetrics to Pixels when ScaleX/Y is not available
    If Horizontally Then
        ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
    Else
        ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
    End If
End Function

