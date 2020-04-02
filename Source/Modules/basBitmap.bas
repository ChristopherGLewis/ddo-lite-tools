Attribute VB_Name = "basBitmap"
Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO1
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type

Private Type BITMAPINFO8
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Declare Function CreateDIBSection1 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO1, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection8 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO8, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Sub GrayScale(pic As PictureBox)
    ' // Convert pic to GrayScale //
    Dim DeskWnd As Long, DeskDC As Long
    Dim MyDC As Long
    Dim MyDIB As Long, OldDIB As Long
    Dim DIBInf As BITMAPINFO8
    Dim MakePal As Long
    
    pic.AutoRedraw = True
    
    ' Create DC based on desktop DC
    DeskWnd = GetDesktopWindow()
    DeskDC = GetDC(DeskWnd)
    MyDC = CreateCompatibleDC(DeskDC)
    ReleaseDC DeskWnd, DeskDC
    ' Validate DC
    If (MyDC = 0) Then Exit Sub
    ' Set DIB information
    With DIBInf
        With .bmiHeader ' Same size as picture
            .biWidth = pic.ScaleX(pic.ScaleWidth, pic.ScaleMode, vbPixels)
            .biHeight = pic.ScaleY(pic.ScaleHeight, pic.ScaleMode, vbPixels)
            .biBitCount = 8
            .biPlanes = 1
            .biClrUsed = 256
            .biClrImportant = 256
            .biSize = Len(DIBInf.bmiHeader)
        End With
        ' Palette is Greyscale
        For MakePal = 0 To 255
            With .bmiColors(MakePal)
                .rgbRed = MakePal
                .rgbGreen = MakePal
                .rgbBlue = MakePal
            End With
        Next MakePal
    End With
    ' Create the DIBSection
    MyDIB = CreateDIBSection8(MyDC, DIBInf, 0, ByVal 0&, 0, 0)
    If (MyDIB) Then ' Validate and select DIB
        OldDIB = SelectObject(MyDC, MyDIB)
        ' Draw original picture to the greyscale DIB
        BitBlt MyDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, pic.hdc, 0, 0, vbSrcCopy
        ' Draw the greyscale image back to picture box 1
        BitBlt pic.hdc, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, MyDC, 0, 0, vbSrcCopy
        ' Clean up DIB
        SelectObject MyDC, OldDIB
        DeleteObject MyDIB
    End If
    ' Clean up DC
    DeleteDC MyDC
    ' Redraw
    pic.Refresh
End Sub

Public Sub MonoChrome(pic As PictureBox)
    ' // Convert pic to B&W //
    Dim DeskWnd As Long, DeskDC As Long
    Dim MyDC As Long
    Dim MyDIB As Long, OldDIB As Long
    Dim DIBInf As BITMAPINFO1
    
    pic.AutoRedraw = True
    'Create DC based on desktop DC
    DeskWnd = GetDesktopWindow()
    DeskDC = GetDC(DeskWnd)
    MyDC = CreateCompatibleDC(DeskDC)
    ReleaseDC DeskWnd, DeskDC
    'Validate DC
    If (MyDC = 0) Then Exit Sub
    'Set DIB information
    With DIBInf
        With .bmiHeader 'Same size as picture
            .biWidth = pic.ScaleX(pic.ScaleWidth, pic.ScaleMode, vbPixels)
            .biHeight = pic.ScaleY(pic.ScaleHeight, pic.ScaleMode, vbPixels)
            .biBitCount = 1
            .biPlanes = 1
            .biClrUsed = 2
            .biClrImportant = 2
            .biSize = Len(DIBInf.bmiHeader)
        End With
        ' Palette is Black ...
        With .bmiColors(0)
            .rgbRed = &H0
            .rgbGreen = &H0
            .rgbBlue = &H0
        End With
        ' ... and white
        With .bmiColors(1)
            .rgbRed = &HFF
            .rgbGreen = &HFF
            .rgbBlue = &HFF
        End With
    End With
    ' Create the DIBSection
    MyDIB = CreateDIBSection1(MyDC, DIBInf, 0, ByVal 0&, 0, 0)
    If (MyDIB) Then ' Validate and select DIB
        OldDIB = SelectObject(MyDC, MyDIB)
           BitBlt MyDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, pic.hdc, 0, 0, vbSrcCopy
        ' Draw the monochome image back to the picture box
        BitBlt pic.hdc, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, MyDC, 0, 0, vbSrcCopy
        ' Clean up DIB
        SelectObject MyDC, OldDIB
        DeleteObject MyDIB
    End If
    ' Clean up DC
    DeleteDC MyDC
    ' Redraw
    pic.Refresh
End Sub

