VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public interests As New clsInterests

Private colNameValue As New Collection

Public Sub AddString(strName As String, Value As String)
Dim oNameValue As New clsNameValue

    oNameValue.Name = strName
    oNameValue.Value = Value
    oNameValue.Delim = """"
    colNameValue.Add
    
End Sub

Public Sub AddBool(strName As String, Value As Boolean)
Dim oNameValue As New clsNameValue

    oNameValue.Name = strName
    oNameValue.Value = IIf(Value, "true", "false")
    oNameValue.Delim = ""
    colNameValue.Add
    
End Sub

Public Sub AddVar(strName As String, Value As Variant)
Dim oNameValue As New clsNameValue

    oNameValue.Name = strName
    oNameValue.Value = CStr(Value)
    oNameValue.Delim = ""
    colNameValue.Add
    
End Sub


Public Function JSON() As String
Dim strJson As String

    strJson = StringFormat("id:""{0}"",title:""{1}"",controlType:{2}, interests:{3}", id, title, controlType, interests.JSON)
    strJson = WrapBr(strJson)
    JSON = strJson
    
End Function

