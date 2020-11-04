Attribute VB_Name = "mdLoadPng"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal sFilename As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function LoadPng(sFilename As String, lWidth As Long, lHeight As Long, lChannels As Long, baData() As Byte) As Boolean
    Const ImageLockModeRead As Long = 1
    Const PixelFormat32bppPARGB As Long = &HE200B
    Dim aInput(0 To 3)  As Long
    Dim hBitmap         As Long
    Dim uData           As BitmapData
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    If GdipLoadImageFromFile(StrPtr(sFilename), hBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipBitmapLockBits(hBitmap, ByVal 0, ImageLockModeRead, PixelFormat32bppPARGB, uData) <> 0 Then
        GoTo QH
    End If
    lWidth = uData.Width
    lHeight = uData.Height
    lChannels = 4
    ReDim baData(0 To uData.Stride * uData.Height - 1) As Byte
    Call CopyMemory(baData(0), ByVal uData.Scan0, UBound(baData) + 1)
    '--- success
    LoadPng = True
QH:
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hBitmap, uData)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
End Function
