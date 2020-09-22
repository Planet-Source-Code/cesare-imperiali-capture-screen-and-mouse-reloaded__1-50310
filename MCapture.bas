Attribute VB_Name = "MCapture"
Option Explicit
'This is enough for win2k
'capture screen-works finer on win2k
Private Const VK_SNAPSHOT As Long = &H2C
Private Const KEYEVENTF_KEYUP = &H2
'to press and release the print screen key
Private Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'-----------------------------------------------------------------
Private Declare Function GetVersion Lib "kernel32" () As Long
'-----------------------------------------------------------------
'if win9x, the keybd_event trick to get a printscreen may fail
'thus the following is another way to get a screenshot
'These routines come from Msdn "HowTo Capture screen, a form or any window"
    Private Type PALETTEENTRY
       peRed As Byte
       peGreen As Byte
       peBlue As Byte
       peFlags As Byte
    End Type

    Private Type LOGPALETTE
       palVersion As Integer
       palNumEntries As Integer
       palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
    End Type

    Private Type GUID
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(7) As Byte
    End Type

    #If Win32 Then

       Private Const RASTERCAPS As Long = 38
       Private Const RC_PALETTE As Long = &H100
       Private Const SIZEPALETTE As Long = 104

       Private Type RECT
          Left As Long
          Top As Long
          Right As Long
          Bottom As Long
       End Type

       Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
       Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, _
                ByVal nHeight As Long) As Long
       Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
       Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, _
                ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
       Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
       Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
       Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, _
                ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, _
                ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
       Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
       Private Declare Function GetForegroundWindow Lib "USER32" () As Long
       Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, _
                ByVal bForceBackground As Long) As Long
       Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
       Private Declare Function GetWindowDC Lib "USER32" (ByVal hWnd As Long) As Long
       Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
       Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
       Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
       Private Declare Function GetDesktopWindow Lib "USER32" () As Long

       Private Type PicBmp
          Size As Long
          Type As Long
          hBmp As Long
          hPal As Long
          Reserved As Long
       End Type

       Private Declare Function OleCreatePictureIndirect _
          Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
          ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

    #ElseIf Win16 Then

       Private Const RASTERCAPS As Integer = 38
       Private Const RC_PALETTE As Integer = &H100
       Private Const SIZEPALETTE As Integer = 104

       Private Type RECT
          Left As Integer
          Top As Integer
          Right As Integer
          Bottom As Integer
       End Type

       Private Declare Function CreateCompatibleDC Lib "GDI" ( _
          ByVal hDC As Integer) As Integer
       Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
          ByVal hDC As Integer, ByVal nWidth As Integer, _
          ByVal nHeight As Integer) As Integer
       Private Declare Function GetDeviceCaps Lib "GDI" ( _
          ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer
       Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
          ByVal hDC As Integer, ByVal wStartIndex As Integer, _
          ByVal wNumEntries As Integer, _
          lpPaletteEntries As PALETTEENTRY) As Integer
       Private Declare Function CreatePalette Lib "GDI" ( _
          lpLogPalette As LOGPALETTE) As Integer
       Private Declare Function SelectObject Lib "GDI" ( _
          ByVal hDC As Integer, ByVal hObject As Integer) As Integer
       Private Declare Function BitBlt Lib "GDI" ( _
          ByVal hDCDest As Integer, ByVal XDest As Integer, _
          ByVal YDest As Integer, ByVal nWidth As Integer, _
          ByVal nHeight As Integer, ByVal hDCSrc As Integer, _
          ByVal XSrc As Integer, ByVal YSrc As Integer, _
          ByVal dwRop As Long) As Integer
       Private Declare Function DeleteDC Lib "GDI" ( _
          ByVal hDC As Integer) As Integer
       Private Declare Function GetForegroundWindow Lib "USER" _
          Alias "GetActiveWindow" () As Integer
       Private Declare Function SelectPalette Lib "USER" ( _
          ByVal hDC As Integer, ByVal hPalette As Integer, ByVal _
          bForceBackground As Integer) As Integer
       Private Declare Function RealizePalette Lib "USER" ( _
          ByVal hDC As Integer) As Integer
       Private Declare Function GetWindowDC Lib "USER" ( _
          ByVal hWnd As Integer) As Integer
       Private Declare Function GetDC Lib "USER" ( _
          ByVal hWnd As Integer) As Integer
       Private Declare Function GetWindowRect Lib "USER" ( _
          ByVal hWnd As Integer, lpRect As RECT) As Integer
       Private Declare Function ReleaseDC Lib "USER" ( _
          ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
       Private Declare Function GetDesktopWindow Lib "USER" () As Integer

       Private Type PicBmp
          Size As Integer
          Type As Integer
          hBmp As Integer
          hPal As Integer
          Reserved As Integer
       End Type

       Private Declare Function OleCreatePictureIndirect _
          Lib "oc25.dll" (PictDesc As PicBmp, RefIID As GUID, _
          ByVal fPictureOwnsHandle As Integer, IPic As IPicture) _
          As Integer

    #End If

'------------------------------------------------------------------
'This is enough on win2K
Public Function getBackGround() As StdPicture  'IPictureDisp
        On Error Resume Next
        ' press prntScr key
        Dim picTmp As StdPicture
        'try to clear clipboard
        Do
            DoEvents
            Err.Clear
            Clipboard.Clear
        Loop While Err.Number <> 0
        
        keybd_event VK_SNAPSHOT, 0, 0, 0
        ' release C key
       
        keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
        
        Do
            
            DoEvents
            Err.Clear
            Set picTmp = Clipboard.GetData(vbCFDIB)
                
        Loop While Err.Number <> 0
        DoEvents
        
        Set getBackGround = picTmp
End Function
'------------------------------------------------------------------
'The following is required on win9x
#If Win32 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Long, _
            ByVal hPal As Long) As Picture

            Dim r As Long

#ElseIf Win16 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Integer, _
            ByVal hPal As Integer) As Picture

            Dim r As Integer
      
#End If
         Dim Pic As PicBmp
         ' IPicture requires a reference to "Standard OLE Types."
         Dim IPic As IPicture
         Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID.
         With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With

         ' Fill Pic with necessary parts.
         With Pic
            .Size = Len(Pic)          ' Length of structure.
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
            .hBmp = hBmp              ' Handle to bitmap.
            .hPal = hPal              ' Handle to palette (may be null).
         End With

         ' Create Picture object.
         r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object.
         Set CreateBitmapPicture = IPic
      End Function
#If Win32 Then
         Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim r As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long
      #ElseIf Win16 Then
         Public Function CaptureWindow(ByVal hWndSrc As Integer, _
            ByVal Client As Boolean, ByVal LeftSrc As Integer, _
            ByVal TopSrc As Integer, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Integer
            Dim hBmp As Integer
            Dim hBmpPrev As Integer
            Dim r As Integer
            Dim hDCSrc As Integer
            Dim hPal As Integer
            Dim hPalPrev As Integer
            Dim RasterCapsScrn As Integer
            Dim HasPaletteScrn As Integer
            Dim PaletteSizeScrn As Integer
      #End If
         Dim LogPal As LOGPALETTE

         ' Depending on the value of Client get the proper device context.
         If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
         Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window.
         End If

         ' Create a memory device context for the copy process.
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC.
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)

         ' Get screen properties.
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            ' capabilities.
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              ' support.
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette.

         ' If the screen has a palette make a copy and realize it.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette.
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it.
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
         End If

         ' Copy the on-screen image into the memory DC.
         r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)

      ' Remove the new copy of the  on-screen image.
         hBmp = SelectObject(hDCMemory, hBmpPrev)

         ' If the screen has a palette get back the palette that was
         ' selected in previously.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         End If

         ' Release the device context resources back to the system.
         r = DeleteDC(hDCMemory)
         r = ReleaseDC(hWndSrc, hDCSrc)

         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles. Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
      End Function


      Public Function CaptureScreen() As Picture
         #If Win32 Then
            Dim hWndScreen As Long
         #ElseIf Win16 Then
            Dim hWndScreen As Integer
         #End If

         ' Get a handle to the desktop window.
         hWndScreen = GetDesktopWindow()

         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object.

         Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
             Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
      End Function


'Public Function CaptureScreen(Left As Long, Top As Long, Width As Long, Height As Long) As StdPicture
'    Dim srcDC As Long
'    Dim trgDC As Long
'    Dim BMPHandle As Long
'    Dim dm As DEVMODE
'    Dim picTmp As StdPicture
'    srcDC = CreateDC("DISPLAY", "", "", dm)
'    trgDC = CreateCompatibleDC(srcDC)
'    BMPHandle = CreateCompatibleBitmap(srcDC, Width, Height)
'    SelectObject trgDC, BMPHandle
'    BitBlt trgDC, 0, 0, Width, Height, srcDC, Left, Top, SRCCOPY
'    OpenClipboard Screen.ActiveForm.hWnd
'    EmptyClipboard
'    SetClipboardData 2, BMPHandle
'
'    CloseClipboard
'    DeleteDC trgDC
'    ReleaseDC BMPHandle, srcDC
'    Set picTmp = Clipboard.GetData(vbCFBitmap)
'    Set CaptureScreen = picTmp
'End Function

Public Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

