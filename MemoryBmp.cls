VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemoryBmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This is a general purpose class to create a memory Dc.
'It makes using memory Dcs easier to use. It has simple
'methods that will let you set the font, fill areas of the Dc with
'a selected color, get and delete a brush or pen, and do a straight
'copy to a passed object.

'The Fill method accepts a fill color,the start position, and a
'width and height. If you use a width or height of 1
'it will draw a line.

Option Explicit

Private hDcMem As Long
Private hbmpMem As Long
Private hOldBmp As Long
Public bmpWidth As Long
Public bmpHeight As Long
Private UseBrush As Long
Private OldBrush As Long
Private UsePen As Long
Private OldPen As Long
Private rtn As Long
Private FillArea As RECT
Private UseFont As Long
Private mFillColor As Long

Public Function Create(WidthIn As Long, HeightIn As Long) As Long
    
    Dim hWndScn As Long
    Dim hDCScn As Long
    bmpWidth = WidthIn
    bmpHeight = HeightIn
    hWndScn = GetDesktopWindow()
    hDCScn = GetDC(hWndScn)
    hDcMem = CreateCompatibleDC(hDCScn)
    hbmpMem = CreateCompatibleBitmap(hDCScn, WidthIn, HeightIn)
    hOldBmp = SelectObject(hDcMem, hbmpMem)
    rtn = BitBlt(hDcMem, 0, 0, WidthIn, HeightIn, hDCScn, 0, 0, vbWhiteness)
    mFillColor = vbWhite
    Call ReleaseDC(hWndScn, hDCScn)
    Create = hDcMem
    
End Function

Public Function GetPen(ByVal PenType As Long, ByVal PenWidth As Long, ByVal Color As Long) As Long

    UsePen = CreatePen(PenType, PenWidth, Color)
    OldPen = SelectObject(hDcMem, UsePen)
    GetPen = UsePen
    
End Function

Public Sub DeletePen()

    rtn = SelectObject(hDcMem, OldPen)
    rtn = DeleteObject(UsePen)

End Sub

Public Sub DeleteBrush()
    
    rtn = SelectObject(hDcMem, OldBrush)
    rtn = DeleteObject(UseBrush)

End Sub

Public Function GetBrush(ByVal Color As Long) As Long

    UseBrush = CreateSolidBrush(Color)
    OldBrush = SelectObject(hDcMem, UseBrush)
    GetBrush = UseBrush
    
End Function

Public Sub Copy(Dest As Object)

    rtn = BitBlt(Dest.hDC, 0, 0, bmpWidth, bmpHeight, hDcMem, 0, 0, vbSrcCopy)

End Sub

Public Sub Fill(Color As Long, Optional StartX As Long = 0, Optional StartY As Long = 0, _
                               Optional FillWidth As Long = 0, Optional FillHeight As Long = 0)

    Dim MyBrush As Long
    MyBrush = GetBrush(Color)
    FillArea.Left = StartX
    FillArea.Top = StartY
    If FillWidth = 0 Then
        FillArea.Right = bmpWidth - StartX
    Else
        FillArea.Right = FillWidth + StartX
    End If
    If FillHeight = 0 Then
        FillArea.Bottom = bmpHeight - StartY
    Else
        FillArea.Bottom = FillHeight + StartY
    End If
    If FillArea.Right > bmpWidth Then FillArea.Right = bmpWidth
    If FillArea.Bottom > bmpHeight Then FillArea.Bottom = bmpHeight
    rtn = FillRect(hDcMem, FillArea, UseBrush)
    DeleteBrush
    mFillColor = Color

End Sub

Public Sub SetFont(UsehDc As Long)

    UseFont = SelectObject(UsehDc, GetStockObject(SYSTEM_FONT))     'get handle to font
    rtn = SelectObject(UsehDc, UseFont)                             'set UsehDc font back
    rtn = SelectObject(hDcMem, UseFont)                             'set font of MemBmp

End Sub

Private Sub Class_Terminate()
        
    rtn = SelectObject(hDcMem, hOldBmp)
    rtn = DeleteDC(hDcMem)
    rtn = DeleteObject(hbmpMem)
    
End Sub
