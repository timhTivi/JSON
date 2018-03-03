Attribute VB_Name = "modTest"
Option Explicit

Public Sub Test()

Dim oInterest As clsInterest
Dim oCategory As clsCategory

Set oCategory = New clsCategory
With oCategory

    .controlType = 0
    .id = "12"
    .title = "Species"

End With

Set oInterest = New clsInterest
With oInterest

    .id = "123"
    .label = "Sheep"
    .Value = True
    
End With
oCategory.interests.List.Add oInterest

Set oInterest = New clsInterest
With oInterest

    .id = "124"
    .label = "Cattle"
    .Value = False
    
End With
oCategory.interests.List.Add oInterest

Debug.Print oCategory.JSON

End Sub
