VERSION 5.00
Begin VB.Form FrmCapture 
   Caption         =   "Capture Screen"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "FrmCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPathName 
      Height          =   285
      Left            =   2820
      TabIndex        =   2
      Top             =   60
      Width           =   4755
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save To disk"
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1320
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3795
      Left            =   7320
      TabIndex        =   3
      Top             =   420
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   7275
   End
   Begin VB.PictureBox PicContainer 
      AutoRedraw      =   -1  'True
      Height          =   3765
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   7245
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   7305
      Begin VB.PictureBox PicCapture 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2070
         Left            =   60
         ScaleHeight     =   2070
         ScaleWidth      =   3645
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   60
         Width           =   3645
      End
   End
   Begin VB.Timer TCapture 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7260
      Top             =   4080
   End
   Begin VB.CommandButton CmdCapture 
      Caption         =   "Capture screen"
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1320
   End
End
Attribute VB_Name = "FrmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type PCURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
'To grab cursor shape -require at least win98 as per Microsoft documentation...
Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long
'To get a Handle to the cursor
Private Declare Function GetCursor Lib "USER32" () As Long
'To draw cursor shape on bitmap
Private Declare Function DrawIcon Lib "USER32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
     
'to get the cursor position
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
'to end a waiting loopp
Dim GotIt As Boolean
'To use the scrollbars
Dim lngVer As Long
Dim lngHor As Long
Const iconSize As Integer = 9

   
Private Sub CmdCapture_Click()
    
    'hide the form
    Me.Visible = False
    
    'start timer
    TCapture.Enabled = True
    'wait
    Do While Not GotIt
        'let windows work
        DoEvents
    Loop
    
    'reset gotit
    GotIt = False
    
    'enable saving
    CmdSave.Enabled = True
    'show form again
    Me.Visible = True
End Sub

Private Sub CmdSave_Click()
   On Error GoTo errHandler
   SavePicture PicCapture.Picture, TxtPathName.Text
   MsgBox "Picture " & TxtPathName.Text & " saved"
   Exit Sub
errHandler:
   MsgBox "Error saving bmp as " & TxtPathName.Text & vbCrLf & "(" & Err.Description & ")"
End Sub

Private Sub Form_Load()
   'do not let save untill somethging has been captured
   CmdSave.Enabled = False
   'size the internal picture to the size of the screen
    With PicCapture
      .Top = 0
      .Left = 0
      .Width = Screen.Width
      .Height = Screen.Height
      'permit persistent drawing
      .AutoRedraw = True
    End With
    'default path and name of bitmap saved
    TxtPathName.Text = AddSlash(App.Path) & "aaScreen.bmp"
    'initialize scrollbars
    Call InitScroll(VScroll1)
    Call InitScroll(HScroll1)
    'to move inside picture when changing scrollbars values
    'lngVer = PicCapture.Height - PicContainer.Height
    'lngHor = PicCapture.Width - PicContainer.Width
End Sub

Private Sub Form_Resize()
   Dim TheHeight As Long
   Dim TheWidth As Long
   
   If Me.WindowState <> vbMinimized Then
      TheHeight = Me.ScaleHeight - (CmdCapture.Top + CmdCapture.Height + 20 + HScroll1.Height)
      TheWidth = Me.ScaleWidth - VScroll1.Width - 20
      'to move inside picture when changing scrollbars values
       With PicContainer
         If TheHeight > 100 Then
            .Height = TheHeight
            HScroll1.Top = Me.ScaleHeight - HScroll1.Height
            VScroll1.Height = TheHeight
            lngVer = PicCapture.Height - .Height
            'make pictresize
            Call VScroll1_Change
         End If
         If TheWidth > 100 Then
            .Width = TheWidth
            VScroll1.Left = TheWidth + 20
            HScroll1.Width = TheWidth
            lngHor = PicCapture.Width - .Width
            Call HScroll1_Change
         End If
       End With
   End If
End Sub

Private Sub TCapture_Timer()
   
   Dim Point As POINTAPI
   'disable timer
   TCapture.Enabled = False
   'capture screen
   If GetWinVersion >= 5 Then
       PicCapture.PaintPicture MCapture.getBackGround, 0, 0
   Else
   
       PicCapture.PaintPicture MCapture.CaptureScreen, 0, 0
   End If
   
   'get cursor position
   GetCursorPos Point
   
   'now to get the icon of mouse and paint on form the mouse
   Dim pcin As PCURSORINFO
   pcin.hCursor = GetCursor
   pcin.cbSize = Len(pcin)
   Dim ret
   ret = GetCursorInfo(pcin)
   DrawIcon PicCapture.hDC, Point.x - iconSize, Point.y - iconSize, pcin.hCursor
   'The following paint only mouse shape for this app
   'DrawIcon PicCapture.hdc, Point.x - iconSize, Point.y - iconSize, CopyIcon(GetCursor)
   'assign to picture the image
   Set PicCapture.Picture = PicCapture.Image
   'clear clipboard here if you can
   On Error Resume Next
   Clipboard.Clear
   'signal you've done to exit the waiting loop
   GotIt = True
   
   
   
End Sub

Private Function AddSlash(ByVal sPath As String) As String
   'be sure a path ends correctly
   sPath = Trim(sPath)
   If Len(sPath) > 0 Then
      If Right$(sPath, 1) <> "/" Then
         If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
         End If
      End If
      AddSlash = sPath
   End If
End Function

Private Sub VScroll1_Change()
   'make piccapture  move on top down
   PicCapture.Top = -(lngVer * VScroll1.Value \ 100)
End Sub

Private Sub HScroll1_Change()
   'make inside picture mofe on left -right
   PicCapture.Left = -(lngHor * HScroll1.Value \ 100)
End Sub

Private Sub InitScroll(ByVal vS As Object)
   With vS
      .Min = 0
      .Max = 100
      .SmallChange = 2
      .LargeChange = 20
   End With
End Sub
