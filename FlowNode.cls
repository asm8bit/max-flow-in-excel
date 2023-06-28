VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlowNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************
' FlowNode.CLS
'
' Node class for maximum flow example program.
' ************************************************
' Copyright (C) 1997 John Wiley & Sons, Inc.
' All rights reserved. See additional copyright
' information in RIGHTS.TXT.
' ************************************************
Option Explicit

Public id As Integer
Public x As Single
Public y As Single
Public geo_long As Single
Public geo_lat As Single
Public name As String
Public Count As Long
Public Dist As Long
Public excelrow As Integer
Public mysql_id As Long
Public InLink2 As FlowLink
Public Links As Collection
Public NodeStatus As Integer ' Path tree status.
Public InLink As FlowLink    ' The link into the node.

' ************************************************
' Create a new Links collection. We do this here
' instead of simply declaring Links as:
'       Dim Links As New Collection
' so we can quickly empty the collection with:
'       Set Links = Nothing
' ************************************************
Private Sub Class_Initialize()
    Set Links = New Collection
End Sub


