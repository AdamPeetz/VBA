Attribute VB_Name = "ProjectActivity"
Option Explicit

Sub Document_All_Comments()
    Dim Sh As Worksheet
    Dim Cmt As Comment
    Dim r As Long 'counting rows
    Dim w As Byte ' counting sheets
    Application.ScreenUpdating = False
    ' Create a new sheet
    Set Sh = Worksheets.Add
    With Sh
        'Put header "Comment", "Address" & "Author" in A1, B1 & C1 respectively.
        .Cells(1, 1).Value = "Comment"
        .Cells(1, 2).Value = "Address"
        .Cells(1, 3).Value = "Author"
        
        'Put the content of each comment in the workbook in a separate cell - starting from A2.
        'Put the address of each comment - such as the worksheet name and cell address in a separate cell starting from B2 and author from C2.
        r = 2
        For w = 1 To Worksheets.Count
        
            For Each Cmt In Worksheets(w).Comments
                .Cells(r, 1).Value = Cmt.Text
                .Cells(r, 2).Value = Worksheets(w).Name & "! " & Cmt.Parent.Address
                .Cells(r, 3).Value = Cmt.Author
                r = r + 1
            Next Cmt
        Next w
        .Columns.AutoFit
        
    End With
      Application.ScreenUpdating = True
      
End Sub
