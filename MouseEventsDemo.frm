VERSION 5.00
Begin VB.Form frmMouseEventsDemo 
   BackColor       =   &H00FF0000&
   Caption         =   "Mouse Events Demo"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStartOpen 
      Caption         =   "Start Open"
      Height          =   360
      Left            =   5715
      TabIndex        =   7
      Top             =   4170
      Width           =   990
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "End"
      Height          =   360
      Left            =   6825
      TabIndex        =   6
      Top             =   4170
      Width           =   870
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   360
      Left            =   4725
      TabIndex        =   5
      Top             =   4170
      Width           =   870
   End
   Begin VB.PictureBox picDemo 
      Height          =   915
      Left            =   6960
      Picture         =   "MouseEventsDemo.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   900
      TabIndex        =   4
      Top             =   810
      Width           =   960
   End
   Begin VB.TextBox txtEventInfo 
      Height          =   360
      Left            =   4740
      TabIndex        =   3
      Top             =   3660
      Width           =   2835
   End
   Begin VB.TextBox txtDemo 
      Height          =   5220
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "MouseEventsDemo.frx":4206
      Top             =   45
      Width           =   4515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   6915
      TabIndex        =   1
      Top             =   4800
      Width           =   1005
   End
   Begin MouseEventsDemo.ctlMouseEvents MouseEvents 
      Height          =   480
      Left            =   4725
      TabIndex        =   0
      Top             =   90
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Or get event info thru raised event"
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   4890
      TabIndex        =   8
      Top             =   3330
      Width           =   2565
   End
End
Attribute VB_Name = "frmMouseEventsDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()

    Unload Me
    End
    
End Sub

Private Sub cmdStart_Click()

    MouseEvents.TrackEvents True
    
End Sub

Private Sub cmdStartOpen_Click()

    MouseEvents.TrackEvents True, True

End Sub

Private Sub CmdEnd_Click()

    MouseEvents.TrackEvents False

End Sub

Private Sub MouseEvents_MouseEvent(x As Integer, _
                            y As Integer, _
                            Msg As Long, wParam As Long, _
                            Ctl As Object)

    txtEventInfo = x & "," & y & " - " & Msg & _
                       " - " & wParam & " - " & Ctl.Name
    
End Sub

