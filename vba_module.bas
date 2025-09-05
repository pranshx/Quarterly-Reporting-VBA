Attribute VB_Name = "ALargeMacroExample"
Sub MoveCopyPrintDivisionQuarter()

Call aMovebyInput
Call aCreateQuarterlyLog
Call aDuplicateMonthlyLog
Call aPrintTeamQuarter



End Sub

Sub aMovebyInput()
'
' PrintDivision Macro
'
  
  
  Dim Message As String
  Dim Region, YNAnswer As Integer


  Message = "Enter the region of your choice:" & vbCrLf & _
            "1 - Southeast" & vbCrLf & _
            "2 - Northeast" & vbCrLf & _
            "3 - Mid-west" & vbCrLf & _
            "4 - Southwest" & vbCrLf & _
            "5 - Northwest" & vbCrLf & _
            "6 - Far-west"
  Region = InputBox(Message, "Region", "Enter 1, 2, 3, 4, 5, or 6")
  
Select Case Region
    Case 1
    ' SouthEast Macro
    ' Moves SouthEast worksheets to beginning of worksheet list.
        Worksheets("SE Sales").Move Before:=Worksheets(1)
        Worksheets("SE Marketing").Move Before:=Worksheets(1)
        Worksheets("SE Clients").Move Before:=Worksheets(1)
        Worksheets("SE Team").Move Before:=Worksheets(1)
    Case 2
    ' NEMove Macro
    ' Moves Northeast worksheets to beginning of worksheet list.
        Worksheets("NE Sales").Move Before:=Worksheets(1)
        Worksheets("NE Marketing").Move Before:=Worksheets(1)
        Worksheets("NE Clients").Move Before:=Worksheets(1)
        Worksheets("NE Team").Move Before:=Worksheets(1)
    Case 3
    ' Mid-West Macro
    ' Moves Mid-West worksheets to beginning of worksheet list.
        Worksheets("MW Sales").Move Before:=Worksheets(1)
        Worksheets("MW Marketing").Move Before:=Worksheets(1)
        Worksheets("MW Clients").Move Before:=Worksheets(1)
        Worksheets("MW Team").Move Before:=Worksheets(1)
    Case 4
    ' Southwest Macro
    ' Moves Southwest worksheets to beginning of worksheet list.
        Worksheets("SW Sales").Move Before:=Worksheets(1)
        Worksheets("SW Marketing").Move Before:=Worksheets(1)
        Worksheets("SW Clients").Move Before:=Worksheets(1)
        Worksheets("SW Team").Move Before:=Worksheets(1)
    Case 5
    ' Northwest Macro
    ' Moves Northwest worksheets to beginning of worksheet list.
        Worksheets("NW Sales").Move Before:=Worksheets(1)
        Worksheets("NW Marketing").Move Before:=Worksheets(1)
        Worksheets("NW Clients").Move Before:=Worksheets(1)
        Worksheets("NW Team").Move Before:=Worksheets(1)
    Case 6
    ' Far-west Macro
    ' Moves Far-west worksheets to beginning of worksheet list.
        Worksheets("FW Sales").Move Before:=Worksheets(1)
        Worksheets("FW Marketing").Move Before:=Worksheets(1)
        Worksheets("FW Clients").Move Before:=Worksheets(1)
        Worksheets("FW Team").Move Before:=Worksheets(1)
    Case Else
        YNAnswer = MsgBox("You didn't type a number between 1 and 6. Try Again?", vbYesNo)
        If YNAnswer = vbYes Then
        Call aMovebyInput
        End If
    End Select
  
End Sub
Sub aCreateQuarterlyLog()
'
'
'
Range("A20000").End(xlUp).Select
Selection.Offset(3, 0).Select

Selection.Value = "Log"
With Selection.Font
    .Size = 16
    .Bold = True
    
End With
Selection.Offset(1, 0).Select
Selection.Value = "Client Name"
Selection.Font.Bold = True

Selection.Offset(1, 0).Select
Selection.Value = "Contact Name"
Selection.Font.Bold = True

Selection.Offset(0, 1).Select
Selection.Value = "Date"
Selection.Font.Bold = True

Selection.Offset(0, 1).Select
Selection.Value = "Duration"
Selection.Font.Bold = True

Selection.Offset(0, 1).Select
Selection.Value = "Notes:"
Selection.Font.Bold = True

End Sub

Sub aDuplicateMonthlyLog()

Dim quarterSelect, YNAnswer, y As Integer
quarterSelect = InputBox("Which Quarter is this for?")

Select Case quarterSelect
    Case 1
        quarterSelect = 1
    Case 2
        quarterSelect = 4
    Case 3
        quarterSelect = 7
    Case 4
        quarterSelect = 10
    Case Else
        YNAnswer = MsgBox("You didn't enter a valid quarter number. Try Again?", vbYesNo)
            If YNAnswer = vbYes Then
                Call aDuplicateMonthlyLog
            End If
    End Select
        
For y = 1 To 3

Worksheets(1).Copy After:=Worksheets(y)
Worksheets(y + 1).Name = Format(DateSerial(1, quarterSelect + y - 1, 1), "MMMM")

Next y

'
End Sub

Sub aPrintTeamQuarter()

Worksheets(2).PrintOut
Worksheets(3).PrintOut
Worksheets(4).PrintOut

End Sub


