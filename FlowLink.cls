VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlowLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************
' FlowLink.CLS
'
' Link class for maximum flow example program.
' ************************************************
' Copyright (C) 1997 John Wiley & Sons, Inc.
' All rights reserved. See additional copyright
' information in RIGHTS.TXT.
' ************************************************
Option Explicit

Public node1 As FlowNode
Public node2 As FlowNode
Public capacity As Long
Public flow As Long
Public excelrow As Integer
Public reversal As Boolean
Public distance As Integer
Public mysql_id As Long

