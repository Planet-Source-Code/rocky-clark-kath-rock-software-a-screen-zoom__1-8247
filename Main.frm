VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Screen Zoom"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   Begin VB.TextBox txtZoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1000%"
      Top             =   0
      Width           =   570
   End
   Begin VB.CheckBox chkOnTop 
      DownPicture     =   "Main.frx":1042
      Height          =   270
      Left            =   240
      Picture         =   "Main.frx":1138
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   240
   End
   Begin VB.CheckBox chkGrid 
      Height          =   270
      Left            =   0
      Picture         =   "Main.frx":122E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1635
      Left            =   30
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   0
      Top             =   315
      Width           =   1605
   End
   Begin VB.HScrollBar hsbZoom 
      Height          =   270
      LargeChange     =   10
      Left            =   480
      Max             =   1000
      Min             =   25
      TabIndex        =   1
      Top             =   0
      Value           =   25
      Width           =   1200
   End
   Begin VB.Timer tmrZoom 
      Interval        =   50
      Left            =   1710
      Top             =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'User Defined Types
Private Type PointAPI   'API point structure.
    X   As Long
    Y   As Long
End Type

Private Type SizeRect   'Size structure (uses Width, Height instead of bounds)
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type RectAPI    'Rect structure (uses Right, Bottom bounds instead of Width, Height)
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Windows API Blt (BitBlt, PatBlt, StretchBlt) ROP constants.
Private Const SRCCOPY           As Long = &HCC0020
Private Const PATCOPY           As Long = &HF00021

'SetWindowPos Flags.
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FLAGS         As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

'Module level variables.
Private mfScale As Single   'Scale of Zoom percentage (6 = 600%) (6 x Size = 600% increase)
Private mlOldX  As Long     'Holds Last X-coord of mouse
Private mlOldY  As Long     'Holds Last Y-coord of mouse

'Declare the Windows API functions that are to be used.
'Alphabetical order to ease lookup later.
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Function CreateCheckeredBrush(ByVal hDC As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long

Dim X           As Long
Dim Y           As Long
Dim lRet        As Long
Dim hBitmapDC   As Long
Dim hBitmap     As Long
Dim hOldBitmap  As Long
    
    'Convert System Colors if needed
    If lColor1 < 0 Then
        lColor1 = GetSysColor(lColor1 And &HFF&)
    End If
    If lColor2 < 0 Then
        lColor2 = GetSysColor(lColor2 And &HFF&)
    End If
    
    'Create a new DC and Bitmap to draw the Brush
    hBitmapDC = CreateCompatibleDC(hDC)
    hBitmap = CreateCompatibleBitmap(hDC, 8, 8)
    'Select the Bitmap into the DC for drawing
    hOldBitmap = SelectObject(hBitmapDC, hBitmap)
    
    'Draw the Brush's Bitmap (Checkerboard)
    For Y = 0 To 6 Step 2
        For X = 0 To 6 Step 2
            lRet = SetPixelV(hBitmapDC, X, Y, lColor1)
            lRet = SetPixelV(hBitmapDC, X + 1, Y, lColor2)
            lRet = SetPixelV(hBitmapDC, X, Y + 1, lColor2)
            lRet = SetPixelV(hBitmapDC, X + 1, Y + 1, lColor1)
        Next X
    Next Y
    
    'Get the bitmap back out of the DC
    hBitmap = SelectObject(hBitmapDC, hOldBitmap)
    
    'Create the Brush from the bitmap
    CreateCheckeredBrush = CreatePatternBrush(hBitmap)
    
    'Delete the DC and Bitmap to free memory
    lRet = DeleteDC(hBitmapDC)
    lRet = DeleteObject(hBitmap)

End Function

Private Sub DoZoom(ptMouse As PointAPI)

Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect

    'Get the Desktop DC
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    
    'Setup the Destination size for StretchBlt.
    With sizDest
        .Left = 0
        .Top = 0
        .Width = picZoom.ScaleWidth
        .Height = picZoom.ScaleHeight
    End With
    
    'Setup the Source size for StretchBlt.
    With sizSrce
        .Left = ptMouse.X - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.Y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        'Adjust Source and Destination sizes if they don't match.
        'sizSrce.Size * mfScale must= sizDest.Size for acurate scaling.
        'Destination must always be as large or larger than picZoom.
        'Adjust the Width, if needed.
        lTemp = Int(.Width * mfScale)  '(Source.Width * mfScale must= sizDest.Width)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        'Adjust the Height, if needed.
        lTemp = Int(.Height * mfScale) '(sizSrce.Height * mfScale must= sizDest.Height)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    
    'Clear the current contents.
    picZoom.Cls
    
    'Stretch the Desktop (source) into picZoom (dest)
    lRet = StretchBlt(picZoom.hDC, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    
    'Release the Desktop DC
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    
    'Redraw the grid, if needed
    If chkGrid.Value = vbChecked Then
        Call DrawGrid
    End If
    
    picZoom.Refresh
    
End Sub

Private Sub DrawGrid()

Dim iWidth      As Integer
Dim iHeight     As Integer
Dim lRet        As Long
Dim hBrush      As Long
Dim hOldBrush   As Long
Dim fX          As Single
Dim fY          As Single

    If mfScale >= 3 Then
    
        'Create a Checkered Brush (Dark and Light Grey)...
        hBrush = CreateCheckeredBrush(picZoom.hDC, &H808080, &HC0C0C0)
        '...and Select it into the PictureBox
        hOldBrush = SelectObject(picZoom.hDC, hBrush)
        
        iWidth = picZoom.ScaleWidth
        iHeight = picZoom.ScaleHeight
        
        'Draw the gridlines using the checkered pattern brush.
        For fX = 0 To iWidth Step mfScale
            lRet = PatBlt(picZoom.hDC, Int(fX), 0, 1, iHeight, PATCOPY)
        Next
        For fY = 0 To iHeight Step mfScale
            lRet = PatBlt(picZoom.hDC, 0, Int(fY), iWidth, 1, PATCOPY)
        Next
        
        'Put the old Brush back and Delete the new one to free memory
        hBrush = SelectObject(picZoom.hDC, hOldBrush)
        lRet = DeleteObject(hBrush)
    
    End If
    
End Sub
Private Function ValidScale(ByVal fScale As Single) As Single

    'If the user typed an invalid scale,
    'change it to be within Zoom bounds.
    If fScale * 100 > hsbZoom.Max Then
        fScale = hsbZoom.Max / 100
    ElseIf fScale * 100 < hsbZoom.Min Then
        fScale = hsbZoom.Min / 100
    End If
    
    ValidScale = fScale
    
End Function

Private Sub LoadSettings()

    'Load the saved settings from the init file.
    Call RestoreFormSize(Me)
    hsbZoom.Value = GetInitEntry("Settings", "Zoom", CStr(200))
    hsbZoom_Change
    chkGrid.Value = IIf(LCase$(GetInitEntry("Settings", "Grid", "False")) = "true", vbChecked, vbUnchecked)
    chkGrid_Click
    chkOnTop.Value = IIf(LCase$(GetInitEntry("Settings", "OnTop", "False")) = "true", vbChecked, vbUnchecked)
    chkOnTop_Click

End Sub

Private Sub SaveSettings()

Dim lRet As Long

    'Save the current settings to the init file.
    Call SaveFormSize(Me)
    lRet = SetInitEntry("Settings", "Zoom", hsbZoom.Value)
    lRet = SetInitEntry("Settings", "Grid", CStr(chkGrid.Value = vbChecked))
    lRet = SetInitEntry("Settings", "OnTop", CStr(chkOnTop.Value = vbChecked))

End Sub

Private Sub chkGrid_Click()

    'Force the zoom to update.
    mlOldX = -100
    
    'Remove focus from button so there's no focus rect.
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
End Sub

Private Sub chkOnTop_Click()

Dim lRet    As Long
Dim lWinPos As Long

    'Set Window to its new position in the zorder
    lWinPos = IIf(chkOnTop.Value = vbChecked, HWND_TOPMOST, HWND_NOTOPMOST)
    lRet = SetWindowPos(Me.hWnd, lWinPos, 0, 0, 0, 0, SWP_FLAGS)
    
    'Remove focus from button so there's no focus rect.
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
End Sub

Private Sub Form_Load()

    Call LoadSettings

End Sub


Private Sub Form_Resize()

    If Me.WindowState <> vbMinimized Then
        If Me.Width < 1680 Then
            Me.Width = 1680
        ElseIf Me.Height < 1680 Then
            Me.Height = 1680
        Else
            'Move the controls into position.
            chkGrid.Move 0, 0
            chkOnTop.Move chkGrid.Width, 0
            hsbZoom.Move chkGrid.Width + chkOnTop.Width, 0, Me.ScaleWidth - txtZoom.Width - chkGrid.Width - chkOnTop.Width
            txtZoom.Move Me.ScaleWidth - txtZoom.Width, -1
            picZoom.Move 0, hsbZoom.Height, Me.ScaleWidth, Me.ScaleHeight - hsbZoom.Height
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call SaveSettings
    
End Sub


Private Sub hsbZoom_Change()

    'Update the label
    txtZoom.Text = Format$(hsbZoom.Value / 100, "####%")
    
    'Reset mfScale
    mfScale = CSng(hsbZoom.Value) / 100!
    
    'Remove focus from scrollbar so there's no flashing thumb.
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
    'Force the zoom to update
    mlOldX = -100

End Sub


Private Sub hsbZoom_Scroll()

    hsbZoom_Change
    
End Sub


Private Sub tmrZoom_Timer()

Dim lRet    As Long
Dim ptMouse As PointAPI

Static lElapsed As Long

    If Me.WindowState <> vbMinimized Then
        'This code runs 20 times/second*, while the form is not minimized.
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.X <> mlOldX) Or (.Y <> mlOldY) Or (lElapsed >= 250) Then
                'This code runs runs 4 times/second* if no mousemove,
                'or 20 times/second* when mouse is moving.
                Call DoZoom(ptMouse)
                If lElapsed >= 250 Then
                    'This code only runs 4 times/second*.
                    If chkOnTop.Value = vbChecked Then
                        lRet = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
                    End If
                End If
                lElapsed = 0
            End If
            mlOldX = .X
            mlOldY = .Y
        End With
    End If
    
    '* Times/second depends on processor speed. A slower processor may not
    'finish processing one timer event before the next arrives, in which
    'case the new event will be discarded.
    
End Sub


Private Sub txtZoom_GotFocus()

    With txtZoom
        'Remove the "%"
        .Text = CStr(Val(.Text))
        'Select the entire string
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtZoom_KeyPress(KeyAscii As Integer)

    'Allow Numbers and Edit Keys (Backspace) only.
    'Backspace [Asc(8)] won't be affected by this code.
    'Other Edit Keys (Delete, Home, End, PageUp, etc.) fire only
    'KeyDown/KeyUp events and also won't be affected by this code.
    If KeyAscii > 31 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        'Not a number key.
        Beep
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then  '(Asc(13))
        'Force a zoom update, then reselect the textbox. (see _LostFocus)
        picZoom.SetFocus
        DoEvents
        txtZoom.SetFocus
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtZoom_LostFocus()

    'Reset the scale
    mfScale = ValidScale(Val(txtZoom.Text) / 100)
    
    'Update the scrollbar (only fires change event if value changes).
    hsbZoom.Value = mfScale * 100
    
    'Update the textbox in case the scrollbar change event didn't fire.
    txtZoom.Text = Format$(mfScale, "####%")

End Sub


