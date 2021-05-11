VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sensitivty analysis tool"
   ClientHeight    =   4632
   ClientLeft      =   -24
   ClientTop       =   144
   ClientWidth     =   6312
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'------------------------------------------------------------------'
'---User changes should be made only for the following constants---'
'------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------'
'---References to the starting location of data display, row have to be be greater than 2.---'
'---First row in the sheet is used to save the data------------------------------------------'
'--------------------------------------------------------------------------------------------'
Const initialRowIndex As Integer = 4
Const initialColumnIndex As Integer = 2
Const mainWorksheetName As String = "Sensitivity Analysis"

'------------------------------------------------------------------------------------------'
'---If there are named ranges that contain following characters, they should be replaced---'
'------------------------------------------------------------------------------------------'
Const dataEntrySplitSign As String = "@"
Const dataEntryParamSplitSign As String = "#"

'-------------------------------------------------------------------------------------'
'---Cells range from which the sensitivity amount (percentage of the default value)---'
'-------------------------------------------------------------------------------------'
Public valuesChangeArray As Range

Private Sub UserForm_Initialize()
    
    If Cells(initialRowIndex, initialColumnIndex) = "" Then
    
        Call SetDefaultDataForSheet
        Dim rng As Range
        Set rng = Range(Cells(initialRowIndex, initialColumnIndex), Cells(initialRowIndex, initialColumnIndex))
        rng.Columns.AutoFit
        
    End If
    
    Call LoadParameters(isInitialized:=True)
    
End Sub

Private Sub ExecuteProgram_Click()
     
    Call SetNPV
    
    If Cells(1, 2) = "" Or Cells(1, 3) = "" Then
        Exit Sub
    End If
    
    Worksheets(mainWorksheetName).Activate
    
    Application.ScreenUpdating = False
    
    Set valuesChangeArray = Range(Cells(initialRowIndex, initialColumnIndex + 1), Cells(initialRowIndex, initialColumnIndex + 5))
    
    If List_ParameterName.ListCount = 0 Then
        MsgBox ("No data available for analysis.")
        Exit Sub
    End If
    
'   Call ClearExistingSolution      'to include
    Call SetNamesOfParameters       'may be included in ExecuteAnalysis function
    Call ExecuteAnalysis
    Call ResizeColumnsWidth
    Call GraphCreator               'works properly
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub AddValues_Click()

    UserForm2.Show
    Call UserForm2.ClearData

End Sub

Private Sub RemoveFromListButton_Click()

    Dim i As Integer
    
    For i = List_ParameterName.ListCount - 1 To 0 Step -1
    
        If List_ParameterName.Selected(i) = True Then
        
            List_ParameterName.RemoveItem (i)
            List_RangeName.RemoveItem (i)
            List_SheetName.RemoveItem (i)
            Exit Sub
            
        End If
        
    Next
    
End Sub

Private Sub ClearDataButton_Click()
    
    Call ClearData
        
End Sub

Private Sub LoadButton_Click()

    Call LoadParameters(isInitialized:=False)
    Call ResizeColumnsWidth

End Sub

Public Sub SaveButton_Click()
    
    Call SaveParameters
       
End Sub

Sub SetNamesOfParameters()

    Dim i As Integer
    
    For i = 0 To List_ParameterName.ListCount - 1
        Cells(initialRowIndex + 1 + i, initialColumnIndex) = List_ParameterName.List(i)
    Next

End Sub

Sub ExecuteAnalysis()

    Worksheets(mainWorksheetName).Activate
    Dim cell As Range
 
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To List_ParameterName.ListCount - 1
        
        j = 1
        
        For Each cell In valuesChangeArray.Cells
            
            If cell.Value = "" Then GoTo nextCell
            
            Call Formulas.PerformSensitivityAnalysis(initialRowIndex + i + 1, _
                                                     initialColumnIndex + j, _
                                                     List_RangeName.List(i), _
                                                     List_SheetName.List(i), _
                                                     cell.Value)
nextCell:
            j = j + 1

        Next cell
            
    Next
    
    Dim rng As Range
    Set rng = Range(Cells(initialRowIndex + 1, initialColumnIndex), _
                    Cells(initialRowIndex + List_ParameterName.ListCount, initialColumnIndex + 5))
    
    Call SetBordersForRange(rng)
    
    Dim valuesRange As Range
    Set valuesRange = Range(Cells(initialRowIndex + 1, initialColumnIndex + 1), _
                            Cells(initialRowIndex + List_ParameterName.ListCount, initialColumnIndex + 5))
    
    Call SetCurrencyFormatting(valuesRange)
    Call ResizeColumnsWidth
    
End Sub

Sub SaveParameters()
    
    Worksheets(mainWorksheetName).Activate
    
    Dim saveData As String
    Dim i As Integer
    
    If Not Cells(1, 1).Value = "" Then
        
        Dim isDataExisting As Variant: isDataExisting = MsgBox("There is already saved data, overwrite it?", vbOKCancel)
        
        If isDataExisting = vbCancel Then Exit Sub
    
    End If
   
    For i = 0 To List_ParameterName.ListCount - 1
    
        saveData = saveData + List_ParameterName.List(i) & "#" & List_RangeName.List(i) & "#" & List_SheetName.List(i) & "@"
    
    Next
    
    Cells(1, 1) = saveData
    
End Sub


Sub LoadParameters(isInitialized As Boolean)
    
    Worksheets(mainWorksheetName).Activate
    
    If Cells(1, 1).Value = "" Then
    
        If isInitialized = False Then
        
            MsgBox ("No data to load")
            
        End If
        
        Exit Sub
        
    End If
    
    If List_ParameterName.ListCount > 0 Then
    
        Dim reaction As Variant: reaction = MsgBox("You are trying to replace the entered data with the saved data. Continue?", vbOKCancel)
        
        If reaction = vbOK Then
            Call ClearData
        Else
            If reaction = vbCancel Then
                Exit Sub
            End If
        End If
        
    End If
        
    Dim saveData As String: saveData = Cells(1, 1)
    Dim numOfEntries As Integer: numOfEntries = NumberOfEntries(saveData, dataEntrySplitSign)
    
    ReDim dataToLoad(numOfEntries, 2) As String
    
    Dim dataEntries() As String: dataEntries = Split(saveData, dataEntrySplitSign)
    
    Dim i As Integer
    
    For i = 0 To UBound(dataEntries) - LBound(dataEntries) - 1
    
        Dim dataEntryParams() As String
        dataEntryParams = Split(dataEntries(i), dataEntryParamSplitSign)
        dataToLoad(i, 0) = dataEntryParams(0)
        dataToLoad(i, 1) = dataEntryParams(1)
        dataToLoad(i, 2) = dataEntryParams(2)
    
    Next i
    
    For i = 0 To numOfEntries
    
        List_ParameterName.AddItem dataToLoad(i, 0)
        List_RangeName.AddItem dataToLoad(i, 1)
        List_SheetName.AddItem dataToLoad(i, 2)
    
    Next
    
End Sub

Function NumberOfEntries(str As String, chr As String) As Integer

    NumberOfEntries = Len(str) - Len(Replace(str, chr, "")) - 1
    
End Function

Sub ClearData()

    List_ParameterName.Clear
    List_RangeName.Clear
    List_SheetName.Clear

End Sub

Sub ClearExistingSolution()

    If Cells(initialRowIndex + 1, initialColumnIndex).Value <> "" Then
        
        Dim i As Integer: i = 1
        
        Do While Cells(initialRowIndex + i, initialColumnIndex).Value <> ""
            i = i + 1
        Loop
        
        Dim cell As Variant
        
        For Each cell In Range(Cells(initialRowIndex + 1, initialColumnIndex), Cells(initialRowIndex + i, initialColumnIndex + 5))
            cell.Value = ""
        Next cell
        
    End If

End Sub

Sub SetNPV()
    
    If Cells(1, 2) = "" Or Cells(1, 3) = "" Then
        
        Dim rng As Range
        On Error GoTo inputFailed
        Set rng = Application.InputBox("Select cell containing NPV", Type:=8)
       
        Cells(1, 2) = rng.Worksheet.Name
        Cells(1, 3) = rng.address
        
    End If
    
    If False = True Then
inputFailed:
        MsgBox ("Select cell that outputs NPV.")
        SetNPV
    End If
        
End Sub

Sub SetDefaultDataForSheet()

    Application.ScreenUpdating = False

    With Cells(initialRowIndex, initialColumnIndex)
        .Value = "Sensitivity analysis variables"
        .Font.Size = 12
        .Font.Name = "Arial"
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = -0.349986266670736
    End With
    
    Dim upRng As Range
    Set upRng = Range(Cells(initialRowIndex, initialColumnIndex + 1), Cells(initialRowIndex, initialColumnIndex + 5))
    Dim valuesChangeArray(1 To 5) As Double
        valuesChangeArray(1) = -0.5
        valuesChangeArray(2) = -0.15
        valuesChangeArray(3) = 0
        valuesChangeArray(4) = 0.15
        valuesChangeArray(5) = 0.5
    
    Dim cell As Variant
    Dim i As Integer: i = 1
    
    For Each cell In upRng
        With cell
            .Value = valuesChangeArray(i)
            .Style = "Percent"
            .HorizontalAlignment = xlCenter
        End With
        i = i + 1
    Next cell
    
    Dim rng As Range
    Set rng = Range(Cells(initialRowIndex, initialColumnIndex), Cells(initialRowIndex, initialColumnIndex + 5))
    
    Call SetBordersForRange(rng)
    
    For i = 1 To 3
        Cells(1, i).Font.ThemeColor = xlThemeColorDark1
        Cells(1, 1).Font.Size = 1
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Sub SetBordersForRange(rng As Range)

    Application.ScreenUpdating = False

    Dim cell As Variant
    Dim cellSide As Variant
 
    Dim cellSides(0 To 3) As Variant
        cellSides(0) = xlEdgeLeft
        cellSides(1) = xlEdgeRight
        cellSides(2) = xlEdgeTop
        cellSides(3) = xlEdgeBottom
    
    For Each cell In rng.Cells
        For Each cellSide In cellSides
            cell.Borders(cellSide).LineStyle = xlContinuous
            cell.Borders(cellSide).ColorIndex = 0
            cell.Borders(cellSide).TintAndShade = 0
            cell.Borders(cellSide).Weight = xlThin
        Next cellSide
    Next cell

    Application.ScreenUpdating = True

End Sub

Sub SetCurrencyFormatting(rng As Range)

    Application.ScreenUpdating = False

    Dim cell As Variant
    
    For Each cell In rng.Cells
    
        With cell
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Style = "Comma"
            .NumberFormat = "_-* #,##0 _z_³_-;-* #,##0 _z_³_-;_-* ""-""?? _z_³_-;_-@_-"
        End With
        
    Next cell
    
    Application.ScreenUpdating = True
    
End Sub

Sub ResizeColumnsWidth()

    Range(Cells(initialRowIndex, initialColumnIndex), Cells(initialRowIndex + List_ParameterName.ListCount, initialColumnIndex + 5)).Columns.AutoFit

End Sub

'---------------------------------------------------------------------'
'---Chart size has been set by default to fit the size of A4 sheet.---'
'---Chart remains legible when inserting it as an image and setting---'
'---the image width to 16 cm------------------------------------------'
'---------------------------------------------------------------------'

Sub GraphCreator()
    
    Dim reaction As Variant
    reaction = MsgBox("Include chart?", vbYesNo, "Chart wizard")
    If reaction = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim rgbVal As Integer: rgbVal = 215
    
    On Error Resume Next
    ActiveSheet.ChartObjects(mainWorksheetName).Delete
    
    Dim ch As ChartObject
    Set ch = ActiveSheet.ChartObjects.Add(Left:=250, Top:=200, Width:=750, Height:=450)
    ch.Name = mainWorksheetName
    ActiveSheet.ChartObjects(mainWorksheetName).Activate
    
    With ActiveChart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = True
        .ChartTitle.Text = mainWorksheetName
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 25
        .HasLegend = True
        .Legend.Format.TextFrame2.TextRange.Font.Size = 15
        .Legend.Position = xlLegendPositionRight
        .Axes(xlCategory).Border.Weight = 1.5
        .Axes(xlCategory).HasMinorGridlines = True
        .Axes(xlCategory).MinorGridlines.Border.Color = RGB(rgbVal, rgbVal, rgbVal)
        .Axes(xlCategory).TickLabels.Font.Size = 15
        .Axes(xlValue).Border.Weight = 1.5
        .Axes(xlValue).HasMinorGridlines = True
        .Axes(xlValue).MinorGridlines.Border.Color = RGB(rgbVal, rgbVal, rgbVal)
        .Axes(xlValue).TickLabels.Font.Size = 15
        
        Select Case Math.Round(Log(Abs(Formulas.GetNPV)) / Log(10), 0)
        Case 1 To 3
            .Axes(xlValue).DisplayUnit = xlDisplayUnitNone
        Case 4 To 6
            .Axes(xlValue).DisplayUnit = xlThousands
        Case 7 To 9
            .Axes(xlValue).DisplayUnit = xlMillions
        Case Else
            .Axes(xlValue).DisplayUnit = xlDisplayUnitNone
        End Select
            
        .Axes(xlValue).DisplayUnitLabel.Format.TextFrame2.TextRange.Font.Size = 15

        Dim i As Integer
    
        For i = 1 To List_ParameterName.ListCount
            Dim rangeOfData As Range
            Set rangeOfData = Range(Cells(initialRowIndex + i, initialColumnIndex + 1), Cells(initialRowIndex + i, initialColumnIndex + 5))
            .SeriesCollection.NewSeries
            .SeriesCollection(i).XValues = valuesChangeArray
            .SeriesCollection(i).Values = rangeOfData
            .SeriesCollection(i).Name = Cells(initialRowIndex + i, initialColumnIndex).Value
        Next
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub
