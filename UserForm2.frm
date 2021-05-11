VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Data input window"
   ClientHeight    =   1980
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3888
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub AddButton_Click()
    
    If Not RangeNameBox.Value = "" Then
        
        If Not ParameterName.Value = "" Then
        
            UserForm1.List_ParameterName.AddItem ParameterName.Text, UserForm1.List_ParameterName.ListCount
            UserForm1.List_RangeName.AddItem RangeNameBox.Text, UserForm1.List_RangeName.ListCount
            UserForm1.List_SheetName.AddItem SheetNameBox.Text, UserForm1.List_SheetName.ListCount
            
            UserForm2.Hide
        
        Else
        
            MsgBox ("Variable name not entered")
         
        End If
                        
    Else
    
        MsgBox ("Range not selected")
        
    End If
       
End Sub

Private Sub SelectRangeButton_Click()
    
    On Error Resume Next
    
    Dim rng As Range: Set rng = Application.InputBox("Select the data range in which you want the values to change.", Type:=8)
    RangeNameBox.Text = rng.address
    SheetNameBox.Text = rng.Worksheet.Name
   
End Sub

Sub ClearData()

    RangeNameBox.Text = ""
    SheetNameBox.Text = ""
    ParameterName.Text = ""

End Sub

