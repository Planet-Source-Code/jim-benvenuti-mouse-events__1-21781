VERSION 5.00
Begin VB.UserControl ctlMouseEvents 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   60
   ToolboxBitmap   =   "MouseEvents.ctx":0000
   Begin VB.PictureBox picOnOff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   390
      Picture         =   "MouseEvents.ctx":0312
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   1
      Top             =   60
      Width           =   360
   End
   Begin VB.PictureBox picMouse 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   60
      Picture         =   "MouseEvents.ctx":0A14
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   60
      Width           =   360
   End
End
Attribute VB_Name = "ctlMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mColParentCtls As Collection
Private mMemBmp As cMemoryBmp
Private hDcMem As Long
Private mMemBmpZoom As cMemoryBmp
Private hDcZoom As Long
Private mMemBmpBtnFace As cMemoryBmp
Private hDcBtnFace As Long
Private mMemBmpButtons As cMemoryBmp
Private hDcButtons As Long
Private mMemBmpText As cMemoryBmp
Private hDcText As Long
Private mMemBmpBtnText As cMemoryBmp
Private hDcBtnText As Long
Private Borders As cBorders
Private mMEObj As cMouseEventObj
Private mTrackingEvents As Boolean
Private mFaceOpen As Boolean
Private mpicMouseDown As Boolean
Private mShowButtons As Boolean
Private i As Integer
Private rtn As Long
Private mCtl As Object
Private BtnTextRect As RECT
Private TextRect As RECT
Private sText As String
Private Type mButton
    x As Long
    Text As String
    wParam As Long
End Type
Private mButtons(5) As mButton

Event MouseEvent(x As Integer, y As Integer, Msg As Long, wParam As Long, Ctl As Object)

Private Sub TrackMouse()

    'GetMessage is used instead of subclassing (difficult for all Controls (windows)
    'The problem with GetMessage is it waits for a message. Since messages are the major means
    'of communications with windows the wait isn't long.  The DoEvents at the end of the loop
    'ensures that other events get handled. The only problem is if everything is shut down and
    'we are still wating for a message, the App will hang. The test for Ambient.UserMode
    'handles the problem. In tests I haven't found that this loop interferes with other
    'processes, on the contrary, they interfere with this routine and messages are missed
    'if they have extensive loops without DoEvents. If you need to see all messages then
    'subclassing would be the way to go.
    
     Dim rtn As Long
    Dim aMsg As Msg     'See MouseEvents.txt for info for Mouse Messages
    Dim x As Integer
    Dim y As Integer
    Dim hWnd As Long
    Dim Message As Long

    Do While mTrackingEvents = True
        On Error Resume Next
        rtn = Ambient.UserMode          'Just reference Ambient Object to see if still around or Error is raised
        If Err.Number <> 0 Then         'This will handle any cases where Control is destroyed
            TrackOff                    'and we are still in this loop (Clicking X on form will do it)
            Exit Sub
        End If
        On Error GoTo 0
        If Ambient.UserMode = False Then    'It is important to get out of this loop if control is shutting down
            TrackOff                        'or Control will not destroy and App will hang.
            Exit Sub
        End If
        
        rtn = GetMessage(aMsg, 0, 0, 0)  'Use 0,0,0 to get all Messages for all App Windows
        rtn = TranslateMessage(aMsg)
        rtn = DispatchMessage(aMsg)         'Pass message along to Window the message was sent to
        If aMsg.Message >= WM_MOUSEFIRST And aMsg.Message <= WM_MOUSELAST Then      'Just process mouse events
            hWnd = aMsg.hWnd
            Err.Clear
            On Error Resume Next
            Set mCtl = mColParentCtls(Str(hWnd)).Ctl    'Get stored Control (window) info
            If Err.Number = 0 Then
                x = CInt(aMsg.lParam And &HFFF)                     'Client Co-ordinates
                y = CInt((aMsg.lParam And &HFFFF0000) / &H10000)    'screen co-ords are also in aMsg
                    'The info must be broken out. Vb will not let you send a UserDefined Type
                    'inside an Object. aMsg contains a PointApi type
                RaiseEvent MouseEvent(x, y, aMsg.Message, aMsg.wParam, mCtl)
                If mFaceOpen Then
                    ShowEvent x, y, aMsg.wParam, mCtl
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
        DoEvents
    Loop

End Sub

Private Sub ShowEvent(x As Integer, y As Integer, wParam As Long, Ctl As Object)

        'Get here only if mFaceOpen is true
        
    Dim hDC As Long
    hDC = GetDC(Ctl.hWnd)   'Most Controls do not expose their hDc as a property but api call
    If hDC <> 0 Then        'will get it
        rtn = StretchBlt(hDcZoom, 0, 0, 121, 99, _
                         hDC, CLng(x - 5), CLng(y - 4), 11, 9, vbSrcCopy)     'Get x11 zoom
        rtn = BitBlt(hDcZoom, 57, 49, 7, 1, 0, 0, 0, vbDstInvert)             'Put on crosshairs
        rtn = BitBlt(hDcZoom, 60, 46, 1, 7, 0, 0, 0, vbDstInvert)
        rtn = BitBlt(UserControl.hDC, 10, 37, 121, 99, hDcZoom, 0, 0, vbSrcCopy)  'Copy to Control
        If mShowButtons Then                                                    'Surface
            DoButtons wParam
        End If
        DoText x, y, Ctl.Name, hDC
        UserControl.Refresh
    End If
    rtn = ReleaseDC(Ctl.hDC, hDC)
   
End Sub

Private Sub DoButtons(wParam As Long)

        'Gets here only if mFaceOpen and mShowButtons are true
        'wParam contains Button & Key (Shift & Control) info
        'see MouseEvents.txt for more info
        
    Dim i As Integer
    For i = 1 To 5
        If wParam And mButtons(i).wParam Then
            DrawBtn i, 4    'Button or Key is Down
        Else
            DrawBtn i, 3    'Button or Key is Up
        End If
    Next
    rtn = BitBlt(UserControl.hDC, 12, 187, 118, 18, hDcButtons, 0, 0, vbSrcCopy)

End Sub

Private Sub DrawBtn(Index As Integer, BorderType As Integer)

    Dim sText As String
    sText = mButtons(Index).Text
    Set mMemBmp = New cMemoryBmp
    hDcMem = mMemBmp.Create(18, 18)
    rtn = BitBlt(hDcMem, 0, 0, 18, 18, hDcBtnFace, 0, 0, vbSrcCopy) 'Get clean Face
    Borders.DrawToHdc hDcMem, BorderType, 0, 0, 18, 18      'Draw Border
    If BorderType = 3 Then  'Up
        rtn = SetTextColor(hDcBtnText, vbBlack)
        rtn = DrawText(hDcBtnText, sText, Len(sText), BtnTextRect, DT_CENTER)
        rtn = BitBlt(hDcMem, 4, 2, 9, 11, hDcBtnText, 0, 0, vbSrcAnd)   ' with White Background
    Else                    'Down                                       'Only Letter is put on
        rtn = SetTextColor(hDcBtnText, vbRed)                           'Face when using vbSrcAnd
        rtn = DrawText(hDcBtnText, sText, Len(sText), BtnTextRect, DT_CENTER)
        rtn = BitBlt(hDcMem, 5, 3, 9, 11, hDcBtnText, 0, 0, vbSrcAnd)
    End If
    rtn = BitBlt(hDcButtons, mButtons(Index).x, 0, 18, 18, hDcMem, 0, 0, vbSrcCopy)
    Set mMemBmp = Nothing
    
End Sub

Private Sub DoText(x As Integer, y As Integer, CtlName As String, hDC As Long)

        'Gets here only if mFaceOpen is true
        
    Dim PointColor As Long
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    Dim sText As String
        'Get Color info (x & y are in Client Area co-ordinates
    PointColor = GetPixel(hDC, CLng(x), CLng(y))                'Get Color of Current Point
    Red = PointColor And &HFF&                                  'Get RGB values
    Green = (PointColor And &HFF00&) / &H100&
    Blue = (PointColor And &HFF0000) / &H10000
        'Display Name of the control Mouse is over
    sText = CtlName
    mMemBmpText.Fill UserControl.BackColor 'Clear Mem Bmp
    rtn = DrawText(hDcText, sText, Len(sText), TextRect, DT_CENTER)
    rtn = BitBlt(UserControl.hDC, 5, 141, 130, 14, hDcText, 0, 0, vbSrcCopy)
        'Display x,y values (x & y are in Client Area co-ordinates
    sText = "x: " & Format(x, "0000") & "  y: " & Format(y, "0000")
    mMemBmpText.Fill UserControl.BackColor
    rtn = DrawText(hDcText, sText, Len(sText), TextRect, DT_CENTER)
    rtn = BitBlt(UserControl.hDC, 5, 156, 130, 14, hDcText, 0, 0, vbSrcCopy)
        'Display color values of curent position
    sText = "rgb(" & Format(Red, "000") & ", " & Format(Green, "000") & ", " & _
                     Format(Blue, "000") & ")"
    mMemBmpText.Fill UserControl.BackColor
    rtn = DrawText(hDcText, sText, Len(sText), TextRect, DT_CENTER)
    rtn = BitBlt(UserControl.hDC, 5, 171, 130, 14, hDcText, 0, 0, vbSrcCopy)

End Sub

Private Sub picOnOff_Click()

        'If just a click then only the tracking will be turned on
        'If picOnOff was DblClicked this event will fire first.
    Dim shWnd As String
    Dim i As Integer
    If mTrackingEvents Then
        TrackOff
    Else
        Set mMemBmp = New cMemoryBmp
        hDcMem = mMemBmp.Create(10, 10)
        mMemBmp.Fill vbGreen
        rtn = BitBlt(picOnOff.hDC, 7, 7, 10, 10, hDcMem, 0, 0, vbSrcCopy)
        Set mMemBmp = Nothing
        For i = 1 To mColParentCtls.Count
            mColParentCtls.Remove 1
        Next
        On Error Resume Next
        For i = 0 To UserControl.ParentControls.Count - 1   'The collection must be built here
            Err.Number = 0                                  'If built in Show all controls may
            shWnd = UserControl.ParentControls(i).hWnd      'not be loaded.
                'Labels and Images do not have handles
            If Err.Number = 0 Then
                AddParentCtl UserControl.ParentControls(i), shWnd
            Else
                Debug.Print UserControl.ParentControls(i).Name
            End If
        Next
        On Error GoTo 0
        Borders.DrawBorder picOnOff, 4
        mTrackingEvents = True
        TrackMouse
    End If

End Sub

Private Sub picOnOff_DblClick()

        'The Click Event will trigger before the DblClick Event so
        'Tracking will be true by the time this fires
        
    If mTrackingEvents Then
        OpenFace
    End If
    
End Sub

Public Sub TrackEvents(Optional Track As Boolean = True, Optional FaceOpen As Boolean = False)

    If Track = True Then
        If FaceOpen = True Then
            OpenFace
        Else
            TrackOff
        End If
        If mTrackingEvents = False Then
            picOnOff_Click              'starts tracking
        End If
    Else
        TrackOff
    End If
    
End Sub

Private Sub TrackOff()
    
        'Closes face of control, turns off tracking and
        'picOnOff turned to Off (borders up and Red indicator)
        
    mTrackingEvents = False
    mFaceOpen = False
    UserControl.Height = 480
    Borders.DrawBorder UserControl.Extender, 2
    Set mMemBmp = New cMemoryBmp
    hDcMem = mMemBmp.Create(10, 10)
    mMemBmp.Fill vbRed
    rtn = BitBlt(picOnOff.hDC, 7, 7, 10, 10, hDcMem, 0, 0, vbSrcCopy)
    Borders.DrawBorder picOnOff, 3

End Sub

Private Sub OpenFace()

        'Opens the face of the control so Zoom and MouseEvent info
        'is visible
        
    If mShowButtons Then
        UserControl.Height = 3135
        DoButtons 0
    Else
        UserControl.Height = 2850
    End If
    InitFace
    Borders.DrawBorder UserControl.Extender, 2
    Borders.DrawBorder UserControl.Extender, 4, 8, 35, 125, 103
    mFaceOpen = True
    
End Sub

Private Sub InitFace()
    
    UserControl.Cls
    UserControl.CurrentX = 29
    UserControl.CurrentY = 8
    UserControl.Print "Mouse Events"

End Sub

    'picMouse is used to move the control at run time
    
Private Sub picMouse_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
       
    Borders.DrawBorder picMouse, 4  'down
    mpicMouseDown = True
    ReleaseCapture
        'allows UserControl to be dragged
    SendMessage UserControl.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub picMouse_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
        'Need to do this on the Move event since the MouseUp event
        'does not happen because the Down event released capture
        
    If mpicMouseDown Then
        mpicMouseDown = False
        Borders.DrawBorder picMouse, 3
    End If

End Sub
    
    'Add, remove and get Parent's Ctls info in collection
    
Private Sub AddParentCtl(ParentCtl As Object, rKey As String)

        'key is the hWnd
        
    Set mMEObj = New cMouseEventObj
    Set mMEObj.Ctl = ParentCtl
    mMEObj.Id = rKey
    mColParentCtls.Add mMEObj, Str(ParentCtl.hWnd)

End Sub

Private Function ParentCtl(hWnd As Long) As Control

    On Error Resume Next
    Set ParentCtl = mColParentCtls(Str(hWnd)).Ctl
    If Err.Number <> 0 Then
        Set ParentCtl = Nothing
        On Error GoTo 0
    End If
    
End Function

Private Sub RemoveParentCtl(rKey As String)

    Dim i As Integer
    For i = 1 To mColParentCtls.Count
        If mColParentCtls(i).Item.Id = rKey Then
            mColParentCtls(i).Remove
            Exit Sub
        End If
    Next
    
End Sub

Public Property Let ShowButtons(ByVal vData As Boolean)
    
    mShowButtons = vData
    PropertyChanged "ShowButtons"

End Property

Public Property Get ShowButtons() As Boolean
    
    ShowButtons = mShowButtons

End Property
    
    'These methods & properties need to be public so the
    'Extender Object has them. (You cannot pass UserControl
    'to a class only UserControl.Extender)
    
Public Sub Refresh()

    UserControl.Refresh
    
End Sub

Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd
    
End Property

Public Property Get hDC() As Long

    hDC = UserControl.hDC
    
End Property

Public Property Get AutoReDraw() As Long

    AutoReDraw = UserControl.AutoReDraw
    
End Property

Private Sub UserControl_Initialize()
    
        'Initialize Collections, classes, Memory BitMaps
        'and rectangles used by the control
        
    On Error GoTo 0     'No error handling yet
    Set mColParentCtls = New Collection     'Used to hold Parent's Ctls info
    Set Borders = New cBorders              'Used to Draw Borders
    UserControl.BackColor = RGB(51, 51, 120)
        'Memory BitMap for a zoom if area around current cursor position
    Set mMemBmpZoom = New cMemoryBmp
    hDcZoom = mMemBmpZoom.Create(121, 99)
        'Memory BitMap and Rectangle used to display text
    Set mMemBmpText = New cMemoryBmp
    hDcText = mMemBmpText.Create(130, 14)
    mMemBmpText.SetFont UserControl.hDC
    rtn = SetTextColor(hDcText, vbYellow)
    rtn = SetBkColor(hDcText, UserControl.BackColor)
    
    TextRect.Left = 0
    TextRect.Top = 0
    TextRect.Right = 129
    TextRect.Bottom = 13
        'Memory BitMaps & Rectangle to display Button and Key activity
    Set mMemBmpBtnFace = New cMemoryBmp
    hDcBtnFace = mMemBmpBtnFace.Create(18, 18)
    rtn = BitBlt(hDcBtnFace, 0, 0, 18, 18, picOnOff.hDC, 3, 3, vbSrcCopy)
        'An array of button/key information
    mButtons(1).x = 0: mButtons(1).wParam = MK_LBUTTON: mButtons(1).Text = "L"
    mButtons(2).x = 25: mButtons(2).wParam = MK_MBUTTON: mButtons(2).Text = "M"
    mButtons(3).x = 50: mButtons(3).wParam = MK_RBUTTON: mButtons(3).Text = "R"
    mButtons(4).x = 75: mButtons(4).wParam = MK_SHIFT: mButtons(4).Text = "S"
    mButtons(5).x = 100: mButtons(5).wParam = MK_CONTROL: mButtons(5).Text = "C"
        
    Set mMemBmpButtons = New cMemoryBmp
    hDcButtons = mMemBmpButtons.Create(118, 18)
    mMemBmpButtons.Fill UserControl.BackColor
    
    Set mMemBmpBtnText = New cMemoryBmp
    hDcBtnText = mMemBmpBtnText.Create(10, 12)
    mMemBmpBtnText.SetFont UserControl.hDC
    
    BtnTextRect.Left = 0
    BtnTextRect.Top = 0
    BtnTextRect.Right = 9
    BtnTextRect.Bottom = 11
        'Init start locations
    picMouse.Top = 4
    picMouse.Left = 4
    picOnOff.Top = 4
    picOnOff.Left = 113
    mTrackingEvents = False
    
End Sub

Private Sub UserControl_Show()
    
        'Initialize StartUp image for both design and run time
        'Tracking is initially Off
    UserControl.Width = 2100
    UserControl.Height = 480
    InitFace
    Borders.DrawBorder UserControl.Extender, 2
    Borders.DrawBorder picMouse, 3
    Borders.DrawBorder picOnOff, 3
    TrackOff
    
End Sub

    'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mShowButtons = PropBag.ReadProperty("ShowButtons", True)

End Sub

    'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ShowButtons", mShowButtons, True)

End Sub

Private Sub UserControl_Terminate()

        'Clean up Collection, classes and Memory BitMaps
    Dim i As Integer
    For i = 1 To mColParentCtls.Count
        mColParentCtls.Remove 1
    Next
    Set mColParentCtls = Nothing
    Set mMEObj = Nothing
    Set mCtl = Nothing
    Set mMemBmpZoom = Nothing
    Set mMemBmpBtnFace = Nothing
    Set mMemBmpButtons = Nothing
    Set mMemBmpText = Nothing
    Set mMemBmpBtnText = Nothing
    Set Borders = Nothing
    mTrackingEvents = False
    
End Sub

