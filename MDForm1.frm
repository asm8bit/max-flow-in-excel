VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MDForm1 
   Caption         =   "����� ����������"
   ClientHeight    =   9150.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   OleObjectBlob   =   "MDForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MDForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDrawY_Click()
'��������� ����� ��������� ������������ ����� � �������������� ��
If ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value <> "" Then
    ListY1 = "����" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
    ListY2 = "����������" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
    Call General.PaintScheme(11, 0)
    Call General.DrawLabel
End If
End Sub

Private Sub cmdSaveSchema_Click()
    Dim L As VbMsgBoxResult
    L = MsgBox("��������� ���������� �����?", vbYesNo)
    If L = vbYes Then
    If ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value <> "" Then
        If CInt(ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value) > 0 Then
            
            ListY1 = "����" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
            ListY2 = "����������" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
            Call General.SaveSchema(ListY1)
            
        End If
    End If
    End If
End Sub

Private Sub cmdStart_Click()
Dim R As Integer

Factor = CInt(cbxFactor.Value)
FixRevers = True
bStop = False
cmdStart.Enabled = False
cmdDrawY.Enabled = False
cmdStop.Enabled = True
Excel.Application.ScreenUpdating = False
     
R = 11
    ListY1 = "����" & cmbNumSch.Value
    ListY2 = "����������" & cmbNumSch.Value
    MDForm1.Caption = "������ ����� �" & cmbNumSch.Value
    
    Call General.Export1(R, R, 0, ListY1, 1)
        'If optAlg1.Value = True Then Call FFB.FindMaxFlowsDist2011(True, 0)
        'If optAlg2.Value = True Then Call FFB.FindMaxFlowsDist2011(True, 1)
    Call General.FindMaxFlowMinDist
    
    Call Clear1(R, R, 2, 2)
    Call General.SaveSolution1(R)
    
faststop:
Excel.Application.ScreenUpdating = True
cmdDrawY.Enabled = True
cmdStop.Enabled = False

Call General.PaintScheme(11, 0)
Call General.DrawLabel
'��������� ����� ����� ��� ������� ����������
ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value = MDForm1.cmbNumSch.Value

MsgBox "��� ������", vbInformation, "������"
End
End Sub

Private Sub Clear1(S As Integer, R As Integer, row1 As Integer, row2 As Integer)
Do While Sheets(ListY1).Cells(row1, 1) <> ""
    Sheets(ListY1).Cells(row1, S + 1).Value = ""
    row1 = row1 + 1
Loop
Do While Sheets(ListY2).Cells(row2, 1) <> ""
    Sheets(ListY2).Cells(row2, R + 1).Value = ""
    row2 = row2 + 1
Loop
End Sub

Private Sub cmdStop_Click()
    bStop = True
End Sub

Private Sub UserForm_Initialize()
'���������� ������

'��� ������� ��������
With Me.cbxFactor
    .Clear
    .AddItem "1"
    .AddItem "10"
    .AddItem "100"
    .AddItem "1000"
    .Value = "1"
End With

cbxColorTheme.AddItem "������-�����"
cbxColorTheme.AddItem "������-�������"
cbxColorTheme.AddItem "����"
cbxColorTheme.AddItem "�/�"
cbxColorTheme.Value = "������-�����"
cbxTrans.AddItem "0%"
cbxTrans.AddItem "30%"
cbxTrans.Value = "0%"

cbxArray.Clear
cbxArray.AddItem "������"
cbxArray.AddItem "������"
cbxArray.Value = "������"

With Me.cmbNumSch
    .Clear
    For R = 1 To 5
        .AddItem CStr(R)
    Next R
    .Value = "1"
End With

With Me.cbxFont
    .Clear
    For R = 8 To 16
        .AddItem CStr(R)
    Next R
    .Value = "10"
End With

With Me.cbxRusEng
    .Clear
    .AddItem "���"
    .AddItem "Eng"
    .Value = "���"
End With

End Sub
