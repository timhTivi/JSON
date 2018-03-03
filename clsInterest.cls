VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public label As String
Public Value As Boolean

Public Function JSON() As String
Dim strJson As String

    strJson = StringFormat("id:""{0}"",label:""{1}"",value:{2}", id, label, Value)
    strJson = WrapBr(strJson)
    JSON = strJson
    
End Function
