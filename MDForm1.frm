VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MDForm1 
   Caption         =   "Форма управления"
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
'Подбираем номер последней отрисованной схемы и перерисовываем ее
If ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value <> "" Then
    ListY1 = "Узлы" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
    ListY2 = "Магистрали" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
    Call General.PaintScheme(11, 0)
    Call General.DrawLabel
End If
End Sub

Private Sub cmdSaveSchema_Click()
    Dim L As VbMsgBoxResult
    L = MsgBox("Сохранить координаты узлов?", vbYesNo)
    If L = vbYes Then
    If ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value <> "" Then
        If CInt(ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value) > 0 Then
            
            ListY1 = "Узлы" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
            ListY2 = "Магистрали" & ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value
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
    ListY1 = "Узлы" & cmbNumSch.Value
    ListY2 = "Магистрали" & cmbNumSch.Value
    MDForm1.Caption = "Расчет схемы №" & cmbNumSch.Value
    
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
'Сохраняем номер схемы для функции сохранения
ThisWorkbook.Worksheets(ListN1).Cells(1, 10).Value = MDForm1.cmbNumSch.Value

MsgBox "Все готово", vbInformation, "Модель"
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
'Показываем версию

'для дробных значений
With Me.cbxFactor
    .Clear
    .AddItem "1"
    .AddItem "10"
    .AddItem "100"
    .AddItem "1000"
    .Value = "1"
End With

cbxColorTheme.AddItem "Красно-Синяя"
cbxColorTheme.AddItem "Красно-Зеленая"
cbxColorTheme.AddItem "Эрта"
cbxColorTheme.AddItem "Ч/Б"
cbxColorTheme.Value = "Красно-Синяя"
cbxTrans.AddItem "0%"
cbxTrans.AddItem "30%"
cbxTrans.Value = "0%"

cbxArray.Clear
cbxArray.AddItem "Кривые"
cbxArray.AddItem "Прямые"
cbxArray.Value = "Прямые"

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
    .AddItem "Рус"
    .AddItem "Eng"
    .Value = "Рус"
End With

End Sub
