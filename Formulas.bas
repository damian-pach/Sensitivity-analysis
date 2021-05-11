Attribute VB_Name = "Formulas"
Const mainWorksheetName As String = "Sensitivity Analysis"

Public Sub ClickMe()

    Call MoveButtonToOtherSheet

End Sub

Sub PerformSensitivityAnalysis(verticalIndex As Integer, horizontalIndex As Integer, rangeToPlace As String, sheetName As String, multiplicationValue As Double)
    
    If multiplicationValue = 0 Then
        Cells(verticalIndex, horizontalIndex) = GetNPV()
        Exit Sub
    End If

    Application.Calculation = xlCalculationManual
    
    Dim OriginalCellFormulas() As Variant
    Dim ModifiedCellValues() As Variant
      
    Dim defaultRng As Range
    Set defaultRng = ThisWorkbook.Sheets(sheetName).Range(rangeToPlace)
    
    On Error GoTo PassOneValue_Formulas
    OriginalCellFormulas = defaultRng.Formula
    If False Then
PassOneValue_Formulas:
        ReDim OriginalCellFormulas(0)
        OriginalCellFormulas(0) = defaultRng.Formula
        GoTo PassOneValue_Values
    End If
    Resume
    
    On Error GoTo PassOneValue_Values
    ModifiedCellValues = defaultRng.Value2
    If False Then
PassOneValue_Values:
        ReDim ModifiedCellValues(0)
        ModifiedCellValues(0) = defaultRng.Value2
    End If
        
    Dim i As Integer: i = 0
    For i = LBound(ModifiedCellValues) To UBound(ModifiedCellValues)
        ModifiedCellValues(i) = ModifiedCellValues(i) * (1 + multiplicationValue)
    Next
    
    defaultRng.Value = ModifiedCellValues
    
    Application.Calculation = xlCalculationAutomatic
    
    Worksheets(mainWorksheetName).Activate
    
    Cells(verticalIndex, horizontalIndex) = GetNPV()
    
    Application.Calculation = xlCalculationManual

    defaultRng.Formula = OriginalCellFormulas
    
    Application.Calculation = xlCalculationAutomatic

End Sub

Public Function GetNPV()
    
    GetNPV = ThisWorkbook.Sheets(Cells(1, 2).Value).Range(Cells(1, 3).Value)
    
End Function

Sub CreateSheet(sheetName As String)

    If WorksheetExists(sheetName) = True Then
        ActiveWorkbook.Worksheets(sheetName).Activate
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    
    sheet.Name = sheetName
    sheet.Activate

End Sub

Private Function WorksheetExists(sheetName As String) As Boolean
   
    For Each sheet In Worksheets
        If sheetName = sheet.Name Then
            WorksheetExists = True
            Exit Function
        Else
            WorksheetExists = False
        End If
    Next sheet

End Function

Private Sub MoveButtonToOtherSheet()
    
    Dim ButtonText As String
    ButtonText = Application.Caller
    'ActiveSheet.Shapes(ButtonText).Delete
    
    ActiveSheet.Shapes.Range(Array(ButtonText)).Select
    Selection.Delete
    
    Call CreateSheet(mainWorksheetName)
    
    Dim m_button As Variant
    Set m_button = ActiveSheet.Buttons.Add(500, 50, 120, 40)
    m_button.OnAction = "DisplayUserform"
    m_button.Characters.Text = "Display Userform"
    
    With m_button.Font
        .Name = "Calibri"
        .Size = 14
    End With

    UserForm1.Show
    
End Sub

Private Function DisplayUserform()

    UserForm1.Show

End Function
