VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public List As New Collection

Public Function JSON() As String
Dim interest As clsInterest
Dim strDelim As String
Dim strOut As String

    For Each interest In List
    
        strOut = strOut & strDelim & interest.JSON
        strDelim = ","
    
    Next

    JSON = WrapSq(strOut)
    
End Function
