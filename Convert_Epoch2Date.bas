Attribute VB_Name = "Module1"
Sub Convert2Date()
'
' Convert to Hyman readible date format
'

'
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.NumberFormat = "[$-en-US]mm/dd/yy hh:mm AM/PM;@"
  
        Dim cellRange As Range
        Set cellRange = Range("A:A")
        Dim i As Long
        For i = 1 To 100000
            
            given_value = cellRange.Cells(i, 1).Value
      
            
            If Not IsEmpty(given_value) Then
                Dim new_cell As Range
                Set new_cell = Range("B:B")
                If IsNumeric(given_value) Then
                    true_value = DateAdd("s", given_value / 1000, "1/1/1970")
                    new_cell.Cells(i, 1).Value = true_value
                Else
                    new_cell.Cells(i, 1).Value = given_value
                End If
            Else
            End If
          
        Next i
    
    
    
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.NumberFormat = "[$-en-US]mm/dd/yy hh:mm AM/PM;@"
  
        Dim end_time As Range
        Set end_time = Range("C:C")
        Dim m As Long
        For m = 1 To 100000
            
            end_given_value = end_time.Cells(m, 1).Value
      
            
            If Not IsEmpty(end_given_value) Then
                Dim end_new_cell As Range
                Set end_new_cell = Range("D:D")
                If IsNumeric(end_given_value) Then
                    end_true_value = DateAdd("s", end_given_value / 1000, "1/1/1970")
                    end_new_cell.Cells(m, 1).Value = end_true_value
                Else
                    end_new_cell.Cells(m, 1).Value = end_given_value
                End If
            Else
            End If
          
        Next m
        
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Columns.AutoFit
    Columns("B:B").Select
    Selection.Columns.AutoFit
    
    MsgBox "Script was coded by Rafail Sharifov"
    
    
End Sub





