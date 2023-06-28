Attribute VB_Name = "General"
Option Explicit
Public strVert1() As String
Public intVertCount() As Integer
Public strVertTupik() As String
Public strW1() As String
Public Nodes As Collection
Public Links As Collection
Public MaxId As Integer
Public Modified As Boolean
Public node_1 As FlowNode
Public node_2 As FlowNode
Public Const RADIUS = 14
Public Const NOT_IN_LIST = 0
Public Const NOW_IN_LIST = 1
Public Const WAS_IN_LIST = 2
Public Const END_CYCLE = 3
Public Const INFINITY = 1000000
Public SourceNode As FlowNode
Public SinkNode As FlowNode
Public node As FlowNode
Public link As FlowLink
Public link_capacity As Integer
Public id1 As String
Public id2 As String
Public TotalFlow As Long
Public row As Integer
Public Factor As Integer
Public Const MAX_W = 3200
Public file_name As String

Global intSizeV As Integer
Global intSizeA As Integer
Global intPrir As Integer
Global FixRevers As Boolean
Global intDR As Integer
Global dDate1 As Date
Global dDate3 As Date
Global bGreenControl1 As Boolean
Global bDateOn As Boolean
Global bShowCapNods As Byte
Global R As Long
Global M As Long
Global intCountTask As Integer


Public bStop As Boolean

Public ListY1 As String 'Узлы
Public ListY2 As String 'Магистрали

Public Const ListN1 = "Схема"
Public Const Scope = 0.5
Public Const constVer = "1.1 от 30-07-2015"


'Сохранить новые координаты узлов
Public Sub SaveSchema(ListY1 As String)
Dim sname As String
Dim label As String
Dim iShape As Shape
Dim row As Integer
Dim Count As Integer
Dim CountR As Integer
Dim x1 As Single
Dim y1 As Single
    Sheets(ListN1).Activate
    Count = Sheets(ListN1).Shapes.Count
    For Each iShape In Sheets(ListN1).Shapes
        With iShape
            sname = Left(.name, 4)
            Select Case sname
                Case "Node"
                    row = 2
                    label = Right(.name, Len(.name) - 5)
                    Do While Sheets(ListY1).Cells(row, 1).Value <> ""
                        If Str(Sheets(ListY1).Cells(row, 1).Value) = Str(label) Then
                            'x1 = .Top + (.Height / 2) 'Работало на Офисе 2003-2007
                            'y1 = .Left + (.Width / 2)
                            x1 = .Top + (10 * Scope) 'Почему поправка 10 в Офисе 2010? Потому-что гладиолус. Ответ в недрах Microsoft...
                            y1 = .Left + (10 * Scope)
                            Sheets(ListY1).Cells(row, 7).Value = x1 / Scope
                            Sheets(ListY1).Cells(row, 6).Value = y1 / Scope
                            CountR = CountR + 1
                            Exit Do
                        End If
                        row = row + 1
                    Loop
            End Select
        End With
    Next iShape
    MsgBox "Сохранение новых положений " & CStr(CountR) & " узлов завершено"
End Sub

'Рисуем этикетку для схемы
Public Sub DrawLabel()
Dim oShape1 As Variant
Dim txt As String
                Set oShape1 = Sheets(ListN1).Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 200, 100)
                With oShape1.DrawingObject
                    .ShapeRange.Fill.Visible = msoFalse
                    .ShapeRange.Line.Visible = msoFalse
                    .ShapeRange.name = "LabelY"
                    txt = "Тестовая схема " & MDForm1.cmbNumSch.Value & Chr(10)
                    txt = txt & "Алгоритм Динца" & Chr(10)
                    If MDForm1.CheckBoxBF.Value Then txt = txt & "Алгоритм Беллмана-Форда" & Chr(10)
                    txt = txt & Chr(12) & "Создано: " & CStr(Now)
                    .Characters.Text = txt
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
End Sub

'Рисуем схему
Public Sub PaintScheme(col As Integer, col_freepower As Integer)
Dim dx As Single
Dim dy As Single
Dim Dist As Single
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim i As Integer

Dim varFontColor1 As Variant
Dim varShape As Variant

Dim sPower As Single
Dim sStream As Single
Dim sngFlowD As Single
Dim iFlow As Integer

Dim bOffice2007 As Boolean
Dim oShape1 As Variant
Dim sLabel As String
Dim vTrans1 As Variant
Dim bGFont As Variant
Dim sFree As Single
'Очистка схемы
Dim sname As String
Dim iShape As Shape

Dim varRGBColors(1 To 8) As Variant 'набор RGB-цветов для схемы
Dim varConn As Variant
Dim varFillColor As Variant
Dim tmpCaption As String

'Dim txt As String
'Dim bSkip As Boolean
'Dim sngPower As Single
'Dim sChartName As String
'Dim tmpNameList As String
'Dim label As String

'Наполняем массив координатами цветов
If MDForm1.cbxColorTheme.Value = "Красно-Синяя" Then
    varRGBColors(1) = RGB(255, 0, 0)
    varRGBColors(2) = RGB(255, 102, 0)
    varRGBColors(3) = RGB(255, 153, 102)
    varRGBColors(4) = RGB(255, 204, 153)
    varRGBColors(5) = RGB(102, 204, 255)
    varRGBColors(6) = RGB(0, 102, 255)
    varRGBColors(7) = RGB(51, 153, 255)
    varRGBColors(8) = RGB(0, 0, 255)
ElseIf MDForm1.cbxColorTheme.Value = "Эрта" Then
    varRGBColors(1) = RGB(0, 64, 116)
    varRGBColors(2) = RGB(0, 97, 178)
    varRGBColors(3) = RGB(0, 97, 178)
    varRGBColors(4) = RGB(42, 121, 208)
    varRGBColors(5) = RGB(42, 121, 208)
    varRGBColors(6) = RGB(157, 194, 235)
    varRGBColors(7) = RGB(157, 194, 235)
    varRGBColors(8) = RGB(157, 194, 235)
ElseIf MDForm1.cbxColorTheme.Value = "Ч/Б" Then
    varRGBColors(1) = RGB(60, 60, 60)
    varRGBColors(2) = RGB(80, 80, 80)
    varRGBColors(3) = RGB(100, 100, 100)
    varRGBColors(4) = RGB(120, 120, 120)
    varRGBColors(5) = RGB(140, 140, 140)
    varRGBColors(6) = RGB(160, 160, 160)
    varRGBColors(7) = RGB(180, 180, 180)
    varRGBColors(8) = RGB(200, 200, 200)
ElseIf MDForm1.cbxColorTheme.Value = "Красно-Зеленая" Then
    varRGBColors(1) = RGB(244, 4, 4)
    varRGBColors(2) = RGB(244, 54, 4)
    varRGBColors(3) = RGB(246, 109, 2)
    varRGBColors(4) = RGB(249, 163, 0)
    varRGBColors(5) = RGB(253, 218, 0)
    varRGBColors(6) = RGB(236, 254, 0)
    varRGBColors(7) = RGB(182, 255, 3)
    varRGBColors(8) = RGB(127, 254, 5)
End If

'Легенда
ThisWorkbook.Sheets(ListN1).Activate
ThisWorkbook.Sheets(ListN1).Cells(10, 2).Value = "Заполнение, %"
For R = 1 To 8
ThisWorkbook.Sheets(ListN1).Cells(10 + R, 2).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = varRGBColors(R)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Next R

'Значения прозрачности, шрифты по умолчанию
vTrans1 = 0#
If MDForm1.cbxTrans.Value = "30%" Then vTrans1 = 0.3
bGFont = CInt(MDForm1.cbxFont.Value)

'Версия офиса, специальные настройки
'ThisWorkbook.Sheets(ListN1).Cells(18, 2).Value = Excel.Application.Version
'If Excel.Application.Version = "12.0" And MDForm1.cbxOffice2003.Value = 0 Then bOffice2007 = True
'If Excel.Application.Version = "14.0" And MDForm1.cbxOffice2003.Value = 0 Then bOffice2007 = True

'Поддерживаются только версии Офиса 2007 и выше
bOffice2007 = True

ThisWorkbook.Worksheets(ListN1).Activate
'Очистка листа Схема
    For Each iShape In Sheets(ListN1).Shapes
            sname = Left(iShape.name, 4)
        Select Case sname
            Case "Node", "Arry", "Text", "Char", "Labe", "Free", "AddN", "Rout"
                iShape.Delete
        End Select
    Next

'Рисуем узлы сети
row = 2
    Do While Sheets(ListY1).Cells(row, 1).Value <> ""
        x1 = Sheets(ListY1).Cells(row, 6).Value * Scope
        y1 = Sheets(ListY1).Cells(row, 7).Value * Scope
        If x1 > 0 And y1 > 0 Then
                'русские/английские названия узлов
                If MDForm1.cbxRusEng.Value = "Рус" Then
                    sLabel = Sheets(ListY1).Cells(row, 3).Value
                Else
                    sLabel = Sheets(ListY1).Cells(row, 3).Value
                End If
                'показывать номера узлов на схеме
                If MDForm1.CheckBox6.Value = True Then sLabel = "№" & Sheets(ListY1).Cells(row, 1).Value & " " & sLabel
                varFillColor = RGB(255, 255, 255)
                'показывать значения в узлах
                If IsNumeric(ThisWorkbook.Sheets(ListY1).Cells(row, col).Value) Then
                    sPower = Sheets(ListY1).Cells(row, col).Value
                End If
                If MDForm1.CheckBox8.Value = True And sPower <> 0 Then
                            sStream = Sheets(ListY1).Cells(row, col + 1).Value
                            sLabel = sLabel & Chr(12) & Format(sStream, "###0.0") & "/" & Format(sPower, "###0.0")
                            If ((Abs(sPower) - Abs(sStream)) > 0.01) Then varFillColor = RGB(255, 235, 235)
                End If
                
            varFontColor1 = 1
            varShape = msoShapeRoundedRectangle

            Set oShape1 = Sheets(ListN1).Shapes.AddShape(varShape, x1 - (Scope * 10), y1 - (Scope * 10), Scope * 20, Scope * 20)

            With oShape1.DrawingObject
                    .ShapeRange.Line.Weight = 1.5
                    .ShapeRange.name = "Node-" & Sheets(ListY1).Cells(row, 1).Value
                    .ShapeRange.Line.Transparency = vTrans1
                    .ShapeRange.Fill.Transparency = vTrans1
                    .ShapeRange.Fill.ForeColor.RGB = varFillColor
                    If sPower = 0 Then
                        .ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)
                    ElseIf sPower > 0 Then
                        .ShapeRange.Line.ForeColor.RGB = RGB(51, 102, 255) 'синий источник
                    Else
                        .ShapeRange.Line.ForeColor.RGB = RGB(255, 0, 0) 'красный потребитель ресурса
                    End If
                    .Characters.Text = sLabel
                    .Characters.Font.Size = bGFont
                    .Characters.Font.Bold = True
                    .Characters.Font.ColorIndex = varFontColor1
                    .AutoSize = True
            End With
        End If
    row = row + 1
Loop


row = 2
    Do While row < 300
        x1 = 0
        y1 = 0
        x2 = 0
        y2 = 0
                x1 = FindX2(Sheets(ListY2).Cells(row, 1).Value) * Scope
                y1 = FindY2(Sheets(ListY2).Cells(row, 1).Value) * Scope
                x2 = FindX2(Sheets(ListY2).Cells(row, 2).Value) * Scope
                y2 = FindY2(Sheets(ListY2).Cells(row, 2).Value) * Scope
                dx = x2 - x1
                dy = y2 - y1
                Dist = Sqr(dx * dx + dy * dy)
        If x1 * y1 * x2 * y2 = 0 Then Dist = 0
        If Sheets(ListY2).Cells(row, 1).Value = 34 Or Sheets(ListY2).Cells(row, 2).Value = 34 Then Dist = 0
        If Dist > 0 Then
            dx = dx * Scope * 75 / Dist
            dy = dy * Scope * 35 / Dist
            'If bOffice2007 = True Then
            '    dx = dx * Scope * 75 / Dist
            '    dy = dy * Scope * 35 / Dist
            'Else
            '    dx = dx * Scope * 20 / Dist
            '    dy = dy * Scope * 20 / Dist
            'End If
            x1 = Int(x1 + dx)
            y1 = Int(y1 + dy)
            x2 = Int(x2 - dx)
            y2 = Int(y2 - dy)
            
              
                    'Получаем значение мощности магистрали
                    sPower = 0
                    If IsNumeric(ThisWorkbook.Sheets(ListY2).Cells(row, col).Value) Then
                            sPower = ThisWorkbook.Sheets(ListY2).Cells(row, 11).Value
                    End If
                    
                    'Получаем значение потока магистрали
                    sngFlowD = 0
                    If IsNumeric(ThisWorkbook.Sheets(ListY2).Cells(row, 12).Value) Then
                        sngFlowD = Sheets(ListY2).Cells(row, 12).Value
                    End If
                
                'If bOffice2007 = True Then
                    varConn = msoConnectorCurve
                    If MDForm1.cbxArray.Value = "Прямые" Then varConn = msoConnectorStraight
                    Set oShape1 = ThisWorkbook.Sheets(ListN1).Shapes.AddConnector(varConn, 1, 1, 1, 1)
                    'oShape1.DrawingObject.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOval
                'Else
                                
                'If sngFlowD >= 0 Then
                '    Set oShape1 = ThisWorkbook.Sheets(ListN1).Shapes.AddLine(x1, y1, x2, y2)
                'Else
                 '   Set oShape1 = ThisWorkbook.Sheets(ListN1).Shapes.AddLine(x2, y2, x1, y1)
                'End If
                '    sPower = ThisWorkbook.Sheets(ListY2).Cells(row, col).Value
                '    With oShape1.DrawingObject.ShapeRange
                '        If sPower = 0 And ThisWorkbook.Sheets(ListY2).Cells(row, 5).Value = 1 Then
                '            .Line.BeginArrowheadStyle = msoArrowheadTriangle
                '        End If
                '        .Line.EndArrowheadStyle = msoArrowheadTriangle
                '        .Line.Weight = Int(sPower / 50) + 1.25
                '        .Line.DashStyle = msoLineSolid
                '        If sPower = 0 Then .Line.DashStyle = 4
                '        .Line.Style = msoLineSingle
                '        .Line.Transparency = vTrans1
                '        .Line.Visible = msoTrue
                '        .name = "Arry-" & Sheets(ListY2).Cells(row, 1).Value & "-" & Sheets(ListY2).Cells(row, 2).Value
                '        .ZOrder msoSendToBack
                '    End With
                   
               'End If

                    
                    With oShape1.DrawingObject.ShapeRange
                    If sPower = 0 And ThisWorkbook.Sheets(ListY2).Cells(row, 5).Value = 1 Then
                        .Line.BeginArrowheadStyle = msoArrowheadTriangle
                    End If
                    .Line.EndArrowheadStyle = msoArrowheadTriangle
                        
                        If sPower > 0 Then
                            .Line.Weight = Int(Log(sPower + 1) * 1) + 3
                        Else
                            .Line.Weight = 3.25
                        End If

                    .Line.DashStyle = msoLineSolid
                    If sPower = 0 Then .Line.DashStyle = 4
                    .Line.Style = msoLineSingle
                    .Line.Transparency = vTrans1
                    .Line.Visible = msoTrue
                    'If bOffice2007 = True Then
                    If sngFlowD >= 0 Then
                        .ConnectorFormat.BeginConnect Sheets(ListN1).Shapes("Node-" & CStr(Sheets(ListY2).Cells(row, 1).Value)), 1
                        .ConnectorFormat.EndConnect Sheets(ListN1).Shapes("Node-" & CStr(Sheets(ListY2).Cells(row, 2).Value)), 1
                    Else
                        .ConnectorFormat.BeginConnect Sheets(ListN1).Shapes("Node-" & CStr(Sheets(ListY2).Cells(row, 2).Value)), 1
                        .ConnectorFormat.EndConnect Sheets(ListN1).Shapes("Node-" & CStr(Sheets(ListY2).Cells(row, 1).Value)), 1
                    End If
                    .RerouteConnections
                        '.BeginDisconnect
                        '.EndDisconnect
                    'End If
                    .name = "Arry-" & Sheets(ListY2).Cells(row, 1).Value & "-" & Sheets(ListY2).Cells(row, 2).Value

            If Sheets(ListY2).Cells(row, col).Value <> 0 Then iFlow = Abs(Int((sngFlowD / sPower) * 100))
            Select Case iFlow
                    Case Is > 90
                        .Line.ForeColor.RGB = varRGBColors(1)
                    Case Is > 80
                        .Line.ForeColor.RGB = varRGBColors(2)
                    Case Is > 70
                        .Line.ForeColor.RGB = varRGBColors(3)
                    Case Is > 60
                        .Line.ForeColor.RGB = varRGBColors(4)
                    Case Is > 50
                        .Line.ForeColor.RGB = varRGBColors(5)
                    Case Is > 30
                        .Line.ForeColor.RGB = varRGBColors(6)
                    Case Is > 10
                        .Line.ForeColor.RGB = varRGBColors(7)
                    Case Else
                        .Line.ForeColor.RGB = varRGBColors(8)
            End Select
            If sPower = 0 Then .Line.ForeColor.RGB = varRGBColors(8) 'если мощность =0 то пунктир синий)
            End With
                
            'Подпись значений на магистрали
            If MDForm1.OptionButton1.Value = False And sPower <> 0 Then
                x1 = (x1 + x2) / 2
                y1 = (y1 + y2) / 2
                    sFree = Abs(Sheets(ListY2).Cells(row, col).Value) - Abs(sngFlowD)
                    
                    If sFree < 0 Then sFree = 0
                    If MDForm1.OptionButton2.Value = True Then
                        If sFree > 0.09 Then
                            tmpCaption = Format((Abs(sPower) - sngFlowD), "###0.00")
                        Else
                            tmpCaption = "Нет"
                        End If
                    ElseIf MDForm1.OptionButton3.Value = True Then
                        tmpCaption = CStr(iFlow) & "%"
                    ElseIf MDForm1.OptionButton4.Value = True Then
                        Select Case Factor
                        Case 100
                            tmpCaption = Format(sngFlowD, "###0.00") & "/" & Format(sPower, "###0.00")
                        Case 10
                            tmpCaption = Format(sngFlowD, "###0.0") & "/" & Format(sPower, "###0.0")
                        Case 1
                            tmpCaption = Format(sngFlowD, "###0") & "/" & Format(sPower, "###0")
                        End Select
                        If Sheets(ListY2).Cells(row, 5).Value = 1 Then
                            tmpCaption = tmpCaption & "R" 'Реверсная магистраль
                        End If
                        
                        tmpCaption = tmpCaption & " D" & Sheets(ListY2).Cells(row, 6).Value
                        
                    End If
                Set oShape1 = Sheets(ListN1).Shapes.AddTextbox(msoTextOrientationHorizontal, x1 - 20, y1 - 10, 60, 20)
                With oShape1.DrawingObject
                    .ShapeRange.name = "Text-" & Sheets(ListY2).Cells(row, 1).Value & "-" & Sheets(ListY2).Cells(row, 2).Value
                    .Characters.Font.Size = bGFont
                    .Characters.Font.Bold = True
                    .Characters.Text = tmpCaption
                    .ShapeRange.Line.Visible = msoFalse
                    .ShapeRange.Fill.Visible = msoTrue
                    .ShapeRange.Fill.Solid
                    .ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .ShapeRange.Fill.Transparency = 0.2
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .AutoSize = msoTrue
                End With
            End If
        End If
row = row + 1
Loop
End Sub

Function FindX2(id As String) As Single
Dim row As Integer
row = 2
    Do While Sheets(ListY1).Cells(row, 1).Value <> ""
        If Sheets(ListY1).Cells(row, 1).Value = id Then
            FindX2 = Sheets(ListY1).Cells(row, 6).Value
            Exit Do
        End If
        row = row + 1
    Loop
End Function

Function FindY2(id As String) As Single
Dim row As Integer
row = 2
    Do While Sheets(ListY1).Cells(row, 1).Value <> ""
        If Sheets(ListY1).Cells(row, 1).Value = id Then
            FindY2 = Sheets(ListY1).Cells(row, 7).Value
            Exit Do
        End If
        row = row + 1
    Loop
End Function

Public Sub Export1(col_n As Integer, col_l As Integer, row_start As Integer, list_name1 As String, fact1 As Single)
Dim intW1 As Long 'Вес ребра (поток)
Dim sgW As Single 'Вес ребра (поток)
Dim num As Integer 'Индекс массивов
Dim bRev As Boolean 'Возможность реверса потока
Dim Dist As Integer 'Расстояние
Dim intS1 As Long
Dim intT1 As Long
Dim i As Integer

' В перспективе, для ускорения работы, проверка на:
'    * повторяющиеся идентификаторы узлов, связей
'    * допустимые значения
'    * небаланс добычи/потребления и максимальный поток сети
'    * тупики и изолированные участки сети (под вопросом)

'Ищем очередной номер файла журнала
i = 1
file_name = ThisWorkbook.Path & "\debug-log_" & CStr(i) & ".txt"
Do While Dir(file_name) <> ""
    i = i + 1
    file_name = ThisWorkbook.Path & "\debug-log_" & CStr(i) & ".txt"
Loop

Open file_name For Append As #1
    Print #1, "Excel version " & Application.Version
    Print #1, "Macros version " & constVer
    Print #1, "Export procedure begin " & Now

On Error GoTo mErr3
NewNetwork
'Исток + Сток
AddNode 1, "Исток", 0
AddNode 2, "Сток", 0
Set SourceNode = Nodes(Str$(1))
Set SinkNode = Nodes(Str$(2))
'Узлы
num = 3
row = 2
bRev = False
Do While Sheets(ListY1).Cells(row, 1).Value <> ""
On Error GoTo mErr1
               
                AddNode Str$(num), Sheets(ListY1).Cells(row, 3).Value, row
                'Проверяем значение мощности узла
                'Integer from -32768 to 32767
                'Long    from - 2 147 483 648 to 2 147 483 647
                If Not IsNumeric(Sheets(list_name1).Cells(row + row_start, col_n).Value) Then GoTo mErr1
                If Abs(Sheets(list_name1).Cells(row + row_start, col_n).Value * Factor) > 2147483647 Then GoTo mErr1
                
                sgW = CSng(Sheets(list_name1).Cells(row + row_start, col_n).Value)
                intW1 = Abs(sgW) * Factor
                
                If sgW > 0 Then
                    AddLink Str$(1), Str$(num), intW1, row + row_start, bRev, 0
                ElseIf sgW < 0 Then
                    AddLink Str$(num), Str$(2), intW1, row + row_start, bRev, 0
                End If
GoTo noErr1:
mErr1:
         Print #1, "ERROR: Skip export row " & CStr(row + row_start) & " list " & list_name1
noErr1:
                num = num + 1
                row = row + 1
Loop

'Магистрали
row = 2
Do While Sheets(ListY2).Cells(row, 1).Value <> ""
On Error GoTo mErr2
    intS1 = id_node(Sheets(ListY2).Cells(row, 1).Value)
    intT1 = id_node(Sheets(ListY2).Cells(row, 2).Value)
    Dist = CInt(Sheets(ListY2).Cells(row, 6).Value)
    bRev = False
    If Sheets(ListY2).Cells(row, 5).Value = 1 Then bRev = True
    intW1 = 0
    'Проверяем значение мощности магистрали
    If Not IsNumeric(Sheets(ListY2).Cells(row, col_l).Value) Then GoTo mErr2
    If Abs(Sheets(ListY2).Cells(row, col_l).Value * Factor) > 2147483647 Then GoTo mErr2
    If Abs(Sheets(ListY2).Cells(row, col_l).Value) > 0 Then
        intW1 = Abs(CSng(Sheets(ListY2).Cells(row, col_l).Value)) * Factor
        AddLink Str$(intS1), Str$(intT1), intW1, row, bRev, Dist
    End If
GoTo noErr2:
mErr2:
         Print #1, "ERROR: Skip export row " & CStr(row + row_start) & " list " & ListY2
noErr2:
    row = row + 1
Loop

If Nodes.Count * Links.Count = 0 Then GoTo mErr3
Print #1, "Export procedure finished " & Now
Close #1
Exit Sub
mErr3:
Print #1, "Export procedure failed"
Close #1
MsgBox "Невозможно создать коллекцию линков и нодов", vbExclamation, "Ваш Excel"
End Sub

Function id_node(excel_id As Integer) As Integer
Dim rowf As Integer
rowf = 2
Do While Sheets(ListY1).Cells(rowf, 1).Value <> ""
            If excel_id = Sheets(ListY1).Cells(rowf, 1).Value Then
                For Each node In Nodes
                    If node.excelrow = rowf Then
                        id_node = node.id
                        Exit Do
                    End If
                Next node
            End If
            rowf = rowf + 1
Loop
End Function

Public Sub AddLink(id1 As String, id2 As String, link_capacity As Long, excel_row As Integer, rev As Boolean, Dist As Integer)
On Error GoTo Skip
                Set link = New FlowLink
                Links.Add link, id1 & "-" & id2
                Set node_1 = Nodes(id1)
                Set node_2 = Nodes(id2)
                Set link.node1 = node_1
                Set link.node2 = node_2
                link.capacity = link_capacity
                link.excelrow = excel_row
                link.reversal = rev
                link.distance = Dist
                link.flow = 0
                node_1.Links.Add link, id2
                node_2.Links.Add link, id1
Exit Sub
Skip:
    Debug.Print "Ошибка добавления связи в строке: " & excel_row
End Sub

Public Sub AddNode(num As Integer, name As String, excel_row As Integer)
On Error GoTo Skip
                Set node = New FlowNode
                node.id = num
                node.name = name
                node.excelrow = excel_row
                Nodes.Add node, Str$(node.id)
Exit Sub
Skip:
    Debug.Print "Ошибка добавления узла в строке: " & excel_row
End Sub


Sub DrawAugm1(node1 As FlowNode, node2 As FlowNode, count1 As Long, min_residual As Long, varColorAugm As Variant)
'varColorRandom = RGB(255, 200, 200)
                x1 = node1.x
                y1 = node1.y
                x2 = node2.x
                y2 = node2.y
                If (x1 * y1 * x2 * y2 > 0) Then
                    x1 = x1 + count1 * 2
                    y1 = y1 + count1 * 2
                    x2 = x2 + count1 * 2
                    y2 = y2 + count1 * 2
                    Set oShape1 = ThisWorkbook.Sheets(ListN1).Shapes.AddLine(x1, y1, x2, y2)
                    With oShape1.DrawingObject.ShapeRange
                        '.Line.BeginArrowheadStyle = msoArrowheadTriangle
                        .Line.EndArrowheadStyle = msoArrowheadTriangle
                        .Line.Weight = 3
                        .Line.DashStyle = msoLineSolid
                        'If sPower = 0 Then .Line.DashStyle = 4
                        .Line.Style = msoLineSingle
                        '.Line.Transparency = 0.3
                        .Line.DashStyle = msoLineSolid
                        .Line.Visible = msoTrue
                        .Line.ForeColor.RGB = varColorAugm
                        .name = "Arry2-" & node1.id & "-" & node2.id & " cnt: " & CStr(count1) & " st:" & CStr(min_residual / Factor)
                        '.ZOrder msoSendToBack
                    End With
                End If
End Sub

Public Sub SaveSolution1(col As Integer)
Dim capacity As Long
Dim flow As Long
Dim excel_row As Integer
        
        For Each link In Links
            id1 = link.node1.id
            id2 = link.node2.id
            capacity = link.capacity
            flow = link.flow
            excel_row = link.excelrow
            'If flow / Factor <> Int(flow / Factor) Then Debug.Print CStr(flow / Factor)
    'If flow <> 0 Then
        If id2 = "2" Then
            Sheets(ListY1).Cells(excel_row, col + 1).Value = -flow / Factor
            If Abs(capacity) - Abs(flow) > 0 Then
                Sheets(ListY1).Cells(excel_row, col + 1).Font.ColorIndex = 3
            Else
                Sheets(ListY1).Cells(excel_row, col + 1).Font.ColorIndex = xlAutomatic
            End If
        ElseIf id1 = "1" Then
            Sheets(ListY1).Cells(excel_row, col + 1).Value = flow / Factor
            If Abs(capacity) - Abs(flow) > 0 Then
                'Sheets(ListY1).Cells(excel_row, col + 1).Font.ColorIndex = 5 'Синий
                Sheets(ListY1).Cells(excel_row, col + 1).Font.ColorIndex = 3 'Красный
            Else
                Sheets(ListY1).Cells(excel_row, col + 1).Font.ColorIndex = xlAutomatic
            End If
        Else
            Sheets(ListY2).Cells(excel_row, col + 1).Value = flow / Factor
        End If
    'End If
    Next link
End Sub
     
Public Sub NewNetwork()
Dim node As FlowNode
    Set Links = New Collection
    Set Nodes = New Collection
    Set SourceNode = Nothing
    Set SinkNode = Nothing
    Set node_1 = Nothing
    Set node_2 = Nothing
    TotalFlow = 0
End Sub

Private Sub ResetFlows()
Dim link As FlowLink
    For Each link In Links
        link.flow = 0
    Next link
    Set SourceNode = Nothing
    Set SinkNode = Nothing
End Sub

Sub FindMaxFlowMinDist()

Dim candidates As New Collection
Dim cycles As New Collection
Dim Residual() As Long
Dim num_nodes As Integer
Dim id1 As Integer
Dim id2 As Integer
Dim node As FlowNode

Dim c_node As FlowNode
Dim to_node As FlowNode
Dim from_node As FlowNode

Dim link As FlowLink

Dim min_residual As Long
Dim Revers As Boolean
'Оптимизация по расстоянию 2010
Dim i As Long
Dim best_i As Integer
Dim best_dist As Long
Dim new_dist As Long

Dim lngSourceCapacity As Long
Dim lngSinkCapacity As Long
'Ведем журнал для отладки
Dim count1 As Long
Dim timestart As Date
'Алгоритм BF
Dim Root As FlowNode 'узел начала поиска отрицательных циклов
Dim link_old As FlowLink 'текущий входящий линк узла
Dim capacity As Long 'значение мощности магистрали
Dim flow As Long 'значение потока
Dim link_count As Long 'для принудительного ограничения обхода схемы
Dim free_push As Long 'можно протолкнуть по отр циклу
Dim need_push As Long 'нужно протолкнуть по отр циклу
Dim It As Long 'порядковый номер итерации алгоритма БФ
Dim optim_val As Long 'произведение потоков на расстояния - результат оптимизации сети
Dim node_dist As Long
Dim min_p As Long

Dim rv As Boolean
Dim opt As Integer
rv = True 'Подтверждаем использование реверсных потоков
opt = 0

timestart = Now

Open file_name For Append As #5
    Print #5, "Начало расчёта " & Now
    Print #5, "Схема №" & MDForm1.cmbNumSch.Value
    Print #5, "Настройки: rv=" & rv & " opt=" & opt & " factor=" & Factor

    If SourceNode Is Nothing Or SinkNode Is Nothing _
        Then Exit Sub
        
        For Each link In SourceNode.Links
            lngSourceCapacity = lngSourceCapacity + link.capacity
        Next link
        
Print #5, "Сумма мощностей истока: " & CStr(lngSourceCapacity / Factor)

        For Each link In SinkNode.Links
            lngSinkCapacity = lngSinkCapacity + link.capacity
        Next link
        
Print #5, "Сумма мощностей стока: " & CStr(lngSinkCapacity / Factor)
   
    ' Dimension the Residual array.
    num_nodes = Nodes.Count
    ReDim Residual(1 To num_nodes, 1 To num_nodes)
    ' Initially the residual values are the same
    ' as the capacities.
    For Each node In Nodes
        For Each link In node.Links
            If link.node1 Is node Then
                Set to_node = link.node2
                id1 = node.id
                id2 = to_node.id
                If rv = False Then Residual(id2, id1) = link.capacity
                Residual(id1, id2) = link.capacity
            Else
                If link.reversal = True Or rv = False Then
                    Set to_node = link.node1
                    id2 = node.id
                    id1 = to_node.id
                    Residual(id1, id2) = link.capacity
                    Residual(id2, id1) = link.capacity
                End If
            End If
        Next link
    Next node
    ' Repeat until we can find no more
    ' augmenting paths.
    Do
        ' Find an augmenting path in the residual
        ' network.
        ' Reset the nodes' NodeStatus and InLink values.
        'DoEvents
        For Each node In Nodes
            node.NodeStatus = NOT_IN_LIST
            Set node.InLink = Nothing
        Next node
        ' Start with an empty candidate list.
        Set candidates = New Collection
        ' Put the source on the candidate list.
        SourceNode.Dist = 0
        candidates.Add SourceNode
        SourceNode.NodeStatus = NOW_IN_LIST
        ' Repeat until the candidate list is empty.
        Do While candidates.Count > 0
                
                best_i = 1
                If opt > 0 Then 'только в случае оптимизации по расстоянию ***OPT***
                best_dist = INFINITY
                    For i = 1 To candidates.Count
                        new_dist = candidates(i).Dist
                        If new_dist < best_dist Then
                            best_i = i 'вкл.\выкл. оптимизации по расстоянию
                            best_dist = new_dist
                        End If
                    Next i
                End If
                
            Set node = candidates(best_i)
            candidates.Remove best_i
            node.NodeStatus = WAS_IN_LIST
            id1 = node.id
            ' Examine the links out of this node.
        For Each link In node.Links
            Revers = False
                    If link.node1 Is node Then
                        Set to_node = link.node2
                    Else
                        Set to_node = link.node1
                    End If
                    id2 = to_node.id
            If Residual(id1, id2) > 0 Then 'And revers = False
                If to_node.NodeStatus = NOT_IN_LIST Then
                    ' The node has not been in the
                    ' candidate list. Add it.
                    If Not to_node Is SinkNode Then candidates.Add to_node
                    to_node.NodeStatus = NOW_IN_LIST
                    to_node.Dist = best_dist + link.distance
                    Set to_node.InLink = link
                ElseIf to_node.NodeStatus = NOW_IN_LIST And opt > 0 Then 'только в случае оптимизации по расстоянию ***OPT***
                    ' The node is in the candidate
                    ' list. Update its Dist and inlink
                    ' values if necessary.
                    new_dist = best_dist + link.distance
                    If new_dist < to_node.Dist Then
                        to_node.Dist = new_dist
                        Set to_node.InLink = link
                    End If
                End If
            End If
        Next link
            ' Stop if the sink has been labeled.
        Loop
        ' Stop if we found no augmenting path.
        If SinkNode.InLink Is Nothing Then
            Exit Do
        End If
        'Print #5, "Dist augm. way: " & SinkNode.Dist
        'Find the smallest residual along the
        'augmenting path.
        min_residual = INFINITY
        Set node = SinkNode
        Do
            If node Is SourceNode Then Exit Do
            id2 = node.id
            Set link = node.InLink
            If link.node1 Is node Then
                Set from_node = link.node2
            Else
                Set from_node = link.node1
            End If
            id1 = from_node.id
            If min_residual > Residual(id1, id2) Then
                min_residual = Residual(id1, id2)
            End If
            Set node = from_node
        Loop
        ' Update the residuals using the
        ' augmenting path.
        Set node = SinkNode
        Do
            If node Is SourceNode Then Exit Do
            id2 = node.id
            Set link = node.InLink
            If link.node1 Is node Then
                Set from_node = link.node2
            Else
                Set from_node = link.node1
            End If
            id1 = from_node.id
            Residual(id1, id2) = Residual(id1, id2) _
                - min_residual
            Residual(id2, id1) = Residual(id2, id1) _
                + min_residual
            Set node = from_node
        Loop
        count1 = count1 + 1 'считаем кол-во найденных маршрутов
        
    Loop ' Repeat until there are no more augmenting paths.
    ' Calculate the flows from the residuals.
        For Each link In Links
            id1 = link.node1.id
            id2 = link.node2.id
            If link.capacity > Residual(id1, id2) Then
                link.flow = link.capacity - Residual(id1, id2)
            ElseIf link.reversal = True Or rv = False Then
                link.flow = Residual(id2, id1) - link.capacity
            End If
        Next link
    ' Find the total flow.
    TotalFlow = 0
    For Each link In SourceNode.Links
        TotalFlow = TotalFlow + Abs(link.flow)
    Next link

    Print #5, "Аугм. путей: " & CStr(count1)
    Print #5, "Максимальный поток: " & TotalFlow / Factor


    If MDForm1.CheckBoxBF.Value = True Then
'*******Bellman-Ford*******
'Готовимся к работе алгоритма
    Print #5, "Start Bellman-Ford:" & Now
    It = 1 'Первая итерация
skip0:
    optim_val = 0 'считаем произведение потоков на расстояния
    For Each link In Links
        optim_val = optim_val + Abs(link.distance * link.flow)
    Next link
    Print #5, "Optimal flows by distance: " & CStr(optim_val / Factor)
    Print #5, "Iteration: " & It
    It = It + 1
    link_count = Links.Count * Nodes.Count 'считаем ограничение для ребер, по которым мы пройдем за одну итерацию
    Set candidates = New Collection
    For Each node In Nodes
        node.Dist = INFINITY
        node.NodeStatus = NOT_IN_LIST
        Set node.InLink2 = Nothing
    Next node
    Set Root = SinkNode
    Root.Dist = 0
    Root.Count = 1 ' Для отрисовки процесса
    Set Root.InLink2 = Nothing
    Root.NodeStatus = NOW_IN_LIST
    candidates.Add Root
    
    Do While candidates.Count > 0
    
        Set node = candidates(1)
        candidates.Remove 1
        node_dist = node.Dist
        node.NodeStatus = WAS_IN_LIST
        Print #5, "View node: " & node.name
        For Each link In node.Links
            'идем по свободным мощностям или по потоку ему навстречу
            link_count = link_count - 1
            
            If link_count < 0 Or It > 100 Then 'не даем алгоритму "зациклиться" или аварийный выход
                Print #5, "***принудительное завершение BF*** " & CStr(Now)
                Exit Do
            End If
            
            min_p = 0 'чтобы предыдущее значение не влияло на поиск
            If node Is link.node1 Then
                Set from_node = link.node1
                Set to_node = link.node2
                    'идем как бы по стрелке
                    If link.flow >= 0 Then
                        min_p = link.capacity - link.flow 'идем по свободной мощности
                        new_dist = node_dist - link.distance 'уменьшаем расстояние
                    Else
                        min_p = -link.flow 'если есть встречный поток
                        new_dist = node_dist + link.distance 'увеличиваем расстояние
                    End If
            Else
                Set to_node = link.node1
                Set from_node = link.node2
                    'идем как бы против стрелки
                    If link.flow > 0 Then
                        min_p = link.flow 'если есть встречный поток
                        new_dist = node_dist + link.distance 'увеличиваем расстояние
                    ElseIf link.flow <= 0 And link.reversal = True Then
                        min_p = link.capacity + link.flow 'если есть свободная мощность в реверсном ребре
                        new_dist = node_dist - link.distance 'уменьшаем расстояние
                    End If
            End If
            Print #5, "View link: " & from_node.name & " " & to_node.name & " min_p: " & min_p & " new_dist: " & new_dist
            If min_p > 0 Then 'Если маршрут(и отрицательный цикл) невозможен в нужном для нас направлении, то прекращаем его обработку
               
                If to_node.NodeStatus = NOT_IN_LIST Then 'этот узел еще не посещали
                    to_node.NodeStatus = NOW_IN_LIST
                    candidates.Add to_node
                    Set to_node.InLink2 = link
                    Print #5, to_node.name & " dist: " & to_node.Dist & " new dist: " & new_dist
                    to_node.Dist = new_dist
                Else
                    Set link_old = to_node.InLink2 'сохраняем предыдущий входящий линк для восстановления узла
                    If to_node.Dist < new_dist Then 'проверяем, нужно ли обновлять расстояние в узле(нашли отрицательный цикл)
                    '############## Отрицательный цикл ##################
                    to_node.NodeStatus = END_CYCLE
                    
                    Print #5, "Negative cycle start: " & to_node.name & " old dist: " & to_node.Dist & " new dist: " & new_dist
                    
                    Set to_node.InLink2 = link
                    Set c_node = to_node
                    
                    free_push = INFINITY
                    need_push = INFINITY
                    i = 0
                    Do
                        'Ищем следующий узел цикла
                        id1 = c_node.id
                        If c_node.InLink2 Is Nothing Then
                            'Цикл проходит через сток?
                            Print #5, "Negative cycle error #1: " & c_node.name
                            to_node.NodeStatus = WAS_IN_LIST
                            candidates.Add to_node
                            Print #5, "(2) Node name: " & to_node.name & " dist: " & to_node.Dist & " new dist: " & new_dist
                            to_node.Dist = new_dist
                            GoTo skip1
                            Exit Do
                        End If
                        'ищем максимальную величину потока в цикле для проталкивания
                        capacity = c_node.InLink2.capacity
                        flow = c_node.InLink2.flow
                        
                        If id1 = c_node.InLink2.node1.id Then
                            id2 = c_node.InLink2.node2.id
                            If c_node.InLink2.reversal = False Then
                                If free_push > flow Then free_push = flow 'это возможность
                                If need_push > flow Then need_push = flow 'это необходимость
                            Else
                                If free_push > (capacity + flow) Then free_push = capacity + flow 'это возможность
                                If flow > 0 And need_push > flow Then need_push = flow 'это необходимость (если поток отрицательный, то необходимость не меняется)
                            End If
                        Else
                            id2 = c_node.InLink2.node1.id
                            'идем против свободной мощности, но как бы по ней
                            If free_push > (capacity - flow) Then free_push = capacity - flow
                            If flow < 0 And need_push > -flow Then need_push = -flow
                        End If
                        
                        Print #5, "In cycle add id1: " & c_node.InLink2.node1.name & " id2: " & c_node.InLink2.node2.name
                        Set c_node = Nodes(id2)
                        i = i + 1
                        
                        If i > Links.Count * Nodes.Count Then
                            Print #5, "Принудительное завершение цикла"
                            free_push = 0
                            need_push = 0
                            Exit Do
                        End If
                        If c_node.NodeStatus = END_CYCLE Then
                            Print #5, "Negative cycle links: " & CStr(i) & " Negative cycle end id:", id1
                            Exit Do
                        End If
                    Loop
                    
                    If need_push > 0 And need_push < INFINITY Then
                    If free_push > 0 And free_push < INFINITY Then
                    If need_push > free_push Then need_push = free_push
                    
                    i = 0
                    Do
                        'Ищем следующий узел цикла
                        id1 = c_node.id
                        
                        If c_node.InLink2 Is Nothing Then
                            'Цикл проходит через сток или другая ошибка
                            Print #5, "Negative cycle error #2: " & c_node.name
                            Exit Do
                        End If
                        
                        'проталкиваем поток в цикле на найденную величину
                        If id1 = c_node.InLink2.node1.id Then
                            id2 = c_node.InLink2.node2.id
                            c_node.InLink2.flow = c_node.InLink2.flow - need_push
                        Else
                            id2 = c_node.InLink2.node1.id
                            c_node.InLink2.flow = c_node.InLink2.flow + need_push
                        End If
                        
                        Print #5, "Push id1: " & c_node.InLink2.node1.name & " id2: " & c_node.InLink2.node2.name & " need_push: " & CStr(need_push / Factor)
                        
                        Set c_node = Nodes(id2)
                        i = i + 1
                        If c_node.NodeStatus = END_CYCLE Then
                            Print #5, "Negative cycle good end. Links: " & CStr(i)
                            GoTo skip0
                            Exit Do
                        End If
                    Loop
                    Else
                            Print #5, "Отбрасываем цикл т.к. нечего проталкивать"
                            'продолжаем просмотр
                            to_node.NodeStatus = WAS_IN_LIST
                            candidates.Add to_node
                            to_node.Dist = new_dist
                    End If
                    End If
                End If
skip1:
                End If 'узел уже посещали
            End If
        Next link
    Loop
        Print #5, "Stop Bellman-Ford:" & Now
    End If
    
    optim_val = 0
    For Each link In Links
        optim_val = optim_val + Abs(link.distance * link.flow)
    Next link
    
    Print #5, "Itog Optimal flows by distance: " & CStr(optim_val / Factor)
    Print #5, "Время расчета, секунд: " & DateDiff("s", timestart, Now)
    Print #5, "------"
    Close #5
    'If MDForm1.chbDrawAW.Value = True Then
    'L = MsgBox("Завершить расчет и визуализацию?", vbYesNo)
    'If L = vbYes Then End
    'End If
End Sub

