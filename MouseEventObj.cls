VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMouseEventObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCtl As Object
Private mId As String

Public Property Let Id(ByVal vData As String)
    
    mId = vData

End Property

Public Property Get Id() As String
    
    Id = mId

End Property

Public Property Set Ctl(ByVal vData As Object)
    
    Set mCtl = vData

End Property

Public Property Get Ctl() As Object
    
    Set Ctl = mCtl

End Property



