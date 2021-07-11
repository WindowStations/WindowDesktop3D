Attribute VB_Name = "modTimer"
Option Explicit
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Timercollection  As New Collection
Public CTimersCol       As New Collection
Private mTimersColCount As Integer
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
   Dim t As Timer
   Dim c As Timers
   On Error Resume Next
   Set t = Timercollection("id:" & idEvent)
   If t Is Nothing Then
      Call apiKillTimer(0, idEvent)
   Else
      If t.ParentsColKey > 0 Then
         Set c = CTimersCol("key:" & t.ParentsColKey)
         If c Is Nothing Then
            Call apiKillTimer(0, idEvent)
         Else
            c.RaiseTimer_Event t.index
         End If
      Else
         t.RaiseTimer_Event
      End If
   End If
   Set t = Nothing
End Sub
Public Function RegisterTimerCollection(ByRef c As Timers) As Integer
   Dim key As String
   mTimersColCount = mTimersColCount + 1
   key = "key:" & mTimersColCount
   CTimersCol.add c, key
   RegisterTimerCollection = mTimersColCount
End Function
