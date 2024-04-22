Attribute VB_Name = "Lesson_a_For_Next"
Option Explicit
Const StartRow As Byte = 10
Dim LastRow As Long

Sub Simple_For()
    Dim i As Long

    Dim myValue As Double
    
    LastRow = Range("A" & StartRow).End(xlDown).Row
    For i = StartRow To LastRow
        myValue = Range("F" & i).Value
        If myValue > 400 Then Range("F" & i).Value = myValue + 10
        If myValue < 0 Then Exit For
        
    Next i
    
End Sub

Sub For_Next_Loop_in_Text()
    Dim i As Long 'for looping inside each cell
    Dim myValue As String
    Dim NumFound As Long
    Dim TxtFound As String
    Dim r As Long 'for looping through rows
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    For r = StartRow To LastRow
        myValue = Range("A" & r).Value
        For i = 1 To VBA.Len(myValue)
            If IsNumeric(VBA.Mid(myValue, i, 1)) Then
                NumFound = NumFound & Mid(myValue, i, 1)
            ElseIf Not IsNumeric(Mid(myValue, i, 1)) Then
                TxtFound = TxtFound & Mid(myValue, i, 1)
            End If
            
        Next i
        Range("H" & r).Value = TxtFound
        Range("I" & r).Value = NumFound
        NumFound = 0
        TxtFound = ""
        
    Next r
    
End Sub

Sub Clear_Values_For_Text_Loop()
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("H" & StartRow, "I" & LastRow).ClearContents
End Sub

Sub Delete_Hidden_filtered_Rows()
    Dim r As Long
    
    LastRow = Range("A" & StartRow).CurrentRegion.Rows.Count + StartRow - 2
    For r = LastRow To StartRow Step -1
        If Rows(r).Hidden = True Then
            'Range("H" & r).Value = "X"
            Rows(r).Delete
            
        End If
    Next r
End Sub

Sub copy_filtered_list()
'Bonus
    ActiveSheet.AutoFilter.Range.Copy
    Worksheets.Add
    Range("A1").PasteSpecial
    
End Sub
