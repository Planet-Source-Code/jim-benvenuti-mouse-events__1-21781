ctlMouseEvents

A UserControl that will capture and report mouse events on form and child controls.  I've found it useful for finding position and color info at runtime when working with picure objects and when doing bitmap manipulation.


Events are reported on the open face of the control. To open 
the face DblClick on the OnOff button on the right or with the 
method
        MouseEvents.TrackEvents True,True

Mouse Events can also be obtained from a raised event

Private Sub MouseEvents_MouseEvent(x As Integer, _
                            y As Integer, _
                            Msg As Long, wParam As Long, _
                            Ctl As Object)

    txtEventInfo = x & "," & y & " - " & Msg & _
                       " - " & wParam & " - " & Ctl.Name    
End Sub

To just get the event info thru the raised event Click on the 
OnOff button or use the method
        MouseEvents.TrackEvents True

To move the control Click on Mouse Button on left and drag.

To end the tracking Click on OnOff button or
         MouseEvents.TrackEvents False

There is a boolean property that allows the button status to be displayed when face is open


Jim Benvenuti
benj@sympatico.ca