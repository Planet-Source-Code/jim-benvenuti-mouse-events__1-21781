VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mShadow(2) As Long
Private mHiLite(2) As Long
Private mStartX As Long
Private mStartY As Long
Private mWidth As Long
Private mHeight As Long
Private mCanvas As Object
Private mMemBmp As cMemoryBmp
Private hDcMem As Long
Private rtn As Long

Public Sub DrawToHdc(hDC As Long, BorderType As Integer, _
                      BorderLeft As Integer, _
                      BorderTop As Integer, _
                      BorderWidth As Integer, _
                      BorderHeight As Integer)

    mStartX = BorderLeft
    mStartY = BorderTop
    mWidth = BorderWidth
    mHeight = BorderHeight

    Set mMemBmp = New cMemoryBmp
    hDcMem = mMemBmp.Create(mWidth, mHeight)
    rtn = BitBlt(hDcMem, 0, 0, mWidth, mHeight, hDC, mStartX, mStartY, vbSrcCopy)
    If BorderType = 0 Then
        SingleBorder
    ElseIf BorderType = 1 Then
        StandardColors
        RecessedBorder
    ElseIf BorderType = 2 Then
        StandardColors
        RaisedBorder
    ElseIf BorderType = 3 Then
        UpColors
        RaisedBorder
    ElseIf BorderType = 4 Then
        DownColors
        RecessedBorder
    Else
        MsgBox "BorderType Incorrect" & " - " & BorderType
    End If
    rtn = BitBlt(hDC, mStartX, mStartY, mWidth, mHeight, hDcMem, 0, 0, vbSrcCopy)
    Set mMemBmp = Nothing

End Sub

Public Sub DrawBorder(Canvas As Object, BorderType As Integer, _
                      Optional BorderLeft As Integer = 0, _
                      Optional BorderTop As Integer = 0, _
                      Optional BorderWidth As Integer = 0, _
                      Optional BorderHeight As Integer = 0)

    Set mCanvas = Canvas
    mStartX = BorderLeft
    mStartY = BorderTop
    mWidth = BorderWidth
    mHeight = BorderHeight
    If mWidth = 0 Then
        mWidth = mCanvas.Width - mStartX
    ElseIf mWidth > mCanvas.Width - mStartX Then
        mWidth = mCanvas.Width - mStartX
    End If
    If mHeight = 0 Then
        mHeight = mCanvas.Height - mStartY
    ElseIf mHeight > mCanvas.Height - mStartY Then
        mHeight = mCanvas.Height - mStartY
    End If
    Set mMemBmp = New cMemoryBmp
    hDcMem = mMemBmp.Create(mWidth, mHeight)
    rtn = BitBlt(hDcMem, 0, 0, mWidth, mHeight, mCanvas.hDC, mStartX, mStartY, vbSrcCopy)
    If BorderType = 0 Then
        SingleBorder
    ElseIf BorderType = 1 Then
        StandardColors
        RecessedBorder
    ElseIf BorderType = 2 Then
        StandardColors
        RaisedBorder
    ElseIf BorderType = 3 Then
        UpColors
        RaisedBorder
    ElseIf BorderType = 4 Then
        DownColors
        RecessedBorder
    Else
        MsgBox "BorderType Incorrect" & " - " & BorderType
    End If
    rtn = BitBlt(mCanvas.hDC, mStartX, mStartY, mWidth, mHeight, hDcMem, 0, 0, vbSrcCopy)
    If mCanvas.AutoReDraw = True Then
        mCanvas.Refresh
    End If
    Set mMemBmp = Nothing
    Set mCanvas = Nothing
    
End Sub

Private Sub SingleBorder()

    mMemBmp.Fill vbBlack, 0, 0, mWidth, 1
    mMemBmp.Fill vbBlack, 0, 0, 1, mHeight - 1
    mMemBmp.Fill vbBlack, 0, mHeight - 1, mWidth - 1, 1
    mMemBmp.Fill vbBlack, mWidth - 1, 0, 1, mHeight
    
End Sub

Private Sub RecessedBorder()

    mMemBmp.Fill mHiLite(2), mWidth - 1, 0, 1, mHeight
    mMemBmp.Fill mHiLite(2), 0, mHeight - 1, mWidth - 1, 1
    mMemBmp.Fill mHiLite(1), mWidth - 2, 1, 1, mHeight - 2
    mMemBmp.Fill mHiLite(1), 1, mHeight - 2, mWidth - 2, 1
    mMemBmp.Fill mShadow(1), 0, 0, mWidth - 1, 1
    mMemBmp.Fill mShadow(1), 0, 0, 1, mHeight - 1
    mMemBmp.Fill mShadow(2), 1, 1, mWidth - 3, 1
    mMemBmp.Fill mShadow(2), 1, 1, 1, mHeight - 3

End Sub

Private Sub RaisedBorder()

    mMemBmp.Fill mHiLite(2), 0, 0, mWidth, 1
    mMemBmp.Fill mHiLite(2), 0, 0, 1, mHeight
    mMemBmp.Fill mHiLite(1), 1, 1, mWidth - 1, 1
    mMemBmp.Fill mHiLite(1), 1, 1, 1, mHeight - 1
    mMemBmp.Fill mShadow(2), mWidth - 1, 0, 1, mHeight
    mMemBmp.Fill mShadow(2), 0, mHeight - 1, mWidth - 1, 1
    mMemBmp.Fill mShadow(1), mWidth - 2, 1, 1, mHeight - 2
    mMemBmp.Fill mShadow(1), 1, mHeight - 2, mWidth - 2, 1

End Sub

Private Sub Class_Initialize()

    StandardColors
    
End Sub

Private Sub StandardColors()

    mShadow(1) = GetSysColor(COLOR_BTNSHADOW)
    mShadow(2) = vbBlack 'RGB(23, 18, 42)
    mHiLite(1) = GetSysColor(COLOR_BTNFACE) 'RGB(190, 180, 225) '
    mHiLite(2) = GetSysColor(COLOR_BTNHIGHLIGHT)

End Sub

Private Sub UpColors()

    mShadow(1) = RGB(74, 74, 56)
    mShadow(2) = RGB(34, 34, 56)
    mHiLite(1) = RGB(134, 134, 133)
    mHiLite(2) = RGB(190, 190, 190)

End Sub

Private Sub DownColors()

    mShadow(1) = RGB(74, 74, 56)
    mShadow(2) = RGB(34, 34, 56)
    mHiLite(1) = RGB(82, 82, 102)
    mHiLite(2) = RGB(101, 101, 145)

End Sub

