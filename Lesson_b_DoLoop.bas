Attribute VB_Name = "Lesson_b_DoLoop"
Option Explicit
Dim StartCell As Integer

Sub Simple_Do_Until_V1()
    StartCell = 8
    Do Until Range("A" & StartCell).Value = ""
        Range("B" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop
End Sub

Sub Simple_Do_Until_V2()
    StartCell = 8
    Do Until StartCell = 14
        Range("B" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop

End Sub

Sub Simple_Do_While()
    StartCell = 8
    Do While Range("A" & StartCell).Value <> ""
        Range("C" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop

End Sub

Sub Simple_Do_Until_Conditional()
    StartCell = 8
    Do Until StartCell = 14
        If Range("A" & StartCell).Value = 0 Then Exit Do
        Range("D" & StartCell).Value = Range("A" & StartCell).Value + 10
        StartCell = StartCell + 1
    Loop
End Sub

Sub Input_Number_Only()

    Dim myAnswer As String
    Do While IsNumeric(myAnswer) = False
    
        myAnswer = VBA.InputBox("Please input Quantity")
        If IsNumeric(myAnswer) Then MsgBox "Well Done!"
    Loop
    
End Sub
