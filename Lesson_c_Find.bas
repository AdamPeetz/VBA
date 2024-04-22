Attribute VB_Name = "Lesson_c_Find"
Option Explicit

Sub One_Find()
    Dim CompId As Range
    Range("C3").ClearContents
    Set CompId = Range("A:A").Find(what:=Range("B3").Value, _
    LookIn:=xlValues, lookat:=xlWhole)
    If Not CompId Is Nothing Then
        Range("C3").Value = CompId.Offset(, 4).Value
    Else
    MsgBox "Company not found!"
    End If
End Sub

Sub Many_Finds()
    Dim CompId As Range
    Dim i As Byte
    Dim FirstMatch As Variant
    
    Range("D3:D6").ClearContents
    i = 3
    Dim Start
    Start = VBA.Timer
    
    Set CompId = Range("A:A").Find(what:=Range("B3").Value, _
    LookIn:=xlValues, lookat:=xlWhole)
    If Not CompId Is Nothing Then
        Range("D" & i).Value = CompId.Offset(, 4).Value
        FirstMatch = CompId.Address
        Do
            Set CompId = Range("A:A").FindNext(CompId)
            If CompId.Address = FirstMatch Then Exit Do
            i = i + 1
            Range("D" & i).Value = CompId.Offset(, 4).Value
        Loop
    Else
    MsgBox "Company not found!"
    End If
    Debug.Print Round(Timer - Start, 3)
    
    
    'Application.Speech.Speak "Well Done. " & i - 2 & " matches were found."
    
End Sub

Sub Counter_Looping_for_Timer()
'for tab Find (for Timer testing)

    Dim r As Long
    Dim i As Byte
    i = 3
    Dim Start
    Start = VBA.Timer
    Range("D3:D6").ClearContents
    'this part is hardcoded for the purpose of testing the Timer Function
    For r = 8 To 200000
        If UCase(Range("A" & r).Value) = UCase(Range("B3").Value) Then
            Range("D" & i).Value = Range("E" & r).Value
            i = i + 1
        End If
    Next r
    
    Debug.Print Round(Timer - Start, 3)
    
End Sub

