'# date: 2022/05/06
'# version: 1.0
'# description: Power State Checker
'# license: MIT
'# github: https://github.com/bakiun/psc
'# author: bakiun

Sub Play(SoundFile)
Dim Sound
Set Sound = CreateObject("WMPlayer.OCX")
Sound.URL = SoundFile
'Sound.settings.volume = 100  # if you want to set volume
Sound.Controls.play
do while Sound.currentmedia.duration = 0
    wscript.sleep 100
loop
wscript.sleep(int(Sound.currentmedia.duration)+1)*1000
End Sub

Play("C:\PSC\hostonline.WAV") '# Play when start the app

Function batteryStatus
    Dim oWMI, items, item
    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set items = oWMI.ExecQuery("Select * from Win32_Battery",,48)
    For Each item in items
        If item.BatteryStatus = 1 Then
            batteryStatus=1
            Exit Function
        ElseIf item.BatteryStatus = 2 then
            batteryStatus=2
            Exit Function
        End If
    Next
End Function

Dim battery_status, prev_status, PathSound
prev_status = batteryStatus
pwLost = "C:\PSC\pwlost.WAV" 'Play when power is lost
pwOnline = "C:\PSC\pwonline.WAV" 'play when power is online
Set colMonitoredEvents = GetObject("winmgmts:\\.\root\cimv2")._
    ExecNotificationQuery("Select * from Win32_PowerManagementEvent")
Do
    Set strLatestEvent = colMonitoredEvents.NextEvent
    If strLatestEvent.EventType = 10 Then
        battery_status = batteryStatus
        If battery_status <> prev_status Then
            If battery_status = 1 Then              
                Call Play(pwlost)
            ElseIf battery_status = 2 Then
                Call Play(pwOnline)
            End If
        End If
    End If
    prev_status = battery_status
Loop



