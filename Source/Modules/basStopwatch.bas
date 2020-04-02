Attribute VB_Name = "basStopwatch"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private mdblFrequency As Double
Private mcurStart As Currency

Public Sub StopwatchInit()
    Dim curFrequency As Currency
    
    QueryPerformanceFrequency curFrequency
    mdblFrequency = CDbl(curFrequency)
End Sub

Public Sub StopwatchStart()
    QueryPerformanceCounter mcurStart
End Sub

Public Function StopwatchStop() As Double
    Dim curStop As Currency
    
    QueryPerformanceCounter curStop
    StopwatchStop = CDbl((curStop - mcurStart) / mdblFrequency)
End Function

Public Function StopwatchStopFormatted() As String
    Dim curStop As Currency
    Dim dblSeconds As Double
    Dim strTime As String
    Dim lngMinutes As Long
    Dim lngSeconds As Long
    
    QueryPerformanceCounter curStop
    dblSeconds = CDbl((curStop - mcurStart) / mdblFrequency)
    Select Case dblSeconds
        Case Is < 0.1
            strTime = Format(dblSeconds, "0.0000000") & " seconds"
        Case Is < 2
            strTime = Format(dblSeconds, "0.000") & " seconds"
        Case Is < 10
            strTime = Format(dblSeconds, "0.00") & " seconds"
        Case Is < 60
            strTime = Format(dblSeconds, "0.0") & " seconds"
        Case Else
            lngSeconds = dblSeconds
            lngMinutes = lngSeconds \ 60
            lngSeconds = lngSeconds Mod 60
            strTime = lngMinutes & ":" & Format(lngSeconds, "00")
    End Select
    StopwatchStopFormatted = strTime
End Function

Public Function StopwatchStopTime() As String
    Dim curStop As Currency
    Dim dblSeconds As Double
    Dim strTime As String
    Dim lngMinutes As Long
    Dim lngSeconds As Long
    Dim lngHours As Long
    
    QueryPerformanceCounter curStop
    lngSeconds = (curStop - mcurStart) / mdblFrequency
    lngMinutes = lngSeconds \ 60
    lngSeconds = lngSeconds Mod 60
    If lngMinutes > 59 Then
        lngHours = lngMinutes \ 60
        lngMinutes = lngMinutes Mod 60
        StopwatchStopTime = lngHours & ":" & Format(lngMinutes, "00") & ":" & Format(lngSeconds, "00")
    Else
        StopwatchStopTime = lngMinutes & ":" & Format(lngSeconds, "00")
    End If
End Function

