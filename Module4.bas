Attribute VB_Name = "Module4"
Sub CalculateButton()

    Dim Calculation_Button As Button
    Dim wkbook As Workbook
    Dim wksheet As Worksheet
    
    Set wkbook = Excel.Application.ThisWorkbook
    Set wksheet = wkbook.Worksheets("Sheet2")
    
    Application.ScreenUpdating = False
    
    Set Calculation_Button = wksheet.Buttons.Add(48.75, 30, 192, 60)
    
    With Calculation_Button
        .Caption = "Calculate"
        .Name = "Calculate"
    End With
    
    Calculation_Button.Select
    Selection.OnAction = "Button_Click"
    
End Sub


