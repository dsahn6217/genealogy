Option Explicit

' =============================================================================
' 광주안씨 분파도(廣州安氏 分派圖) 자동 작도 매크로
' - 박스 내 함자 가로·세로 정중앙 정렬 적용 버전
' =============================================================================

Public shtData As Worksheet, shtPedigree As Worksheet
Public shp As Shape, arrow As Shape
Public rngAll As Range
Public intBottom As Long
Public leftPos As Single, topPos As Single
Public intRows As Integer, intCols As Integer, intRow As Integer, intCol As Integer
Public intBoxW As Integer, intboxH As Integer, intLeft As Integer, intInter As Integer
Public intLineX As Integer, intLineY As Integer
Public startPL1 As Long, startPT1 As Long   ' 텍스트박스의 좌측(Left), 상단(Top) 좌표
Public endPL1 As Long, endPT1 As Long       ' 텍스트박스의 우측(Right), 하단(Bottom) 좌표
Public startPL2 As Long, startPT2 As Long   ' 연결선(LineX)의 시작 좌표
Public endPL2 As Long, endPT2 As Long       ' 연결선(LineX)의 끝 좌표
Public lngLW As Single, lngTLW As Single
Public intM As Integer, intN As Integer, intL As Integer, intT As Integer, intR As Integer, intS As Integer, intColor As Integer
Public varLineX() As Variant
Public strColor As Long, TextOrient As MsoTextOrientation, Textx As String
Public blnEnd As Boolean, blnStart As Boolean

' -----------------------------------------------------------------------------
' [메인 실행 프로시저] 분파도 전체 생성
' -----------------------------------------------------------------------------
Public Sub PedigreeA()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set shtData = ThisWorkbook.Worksheets("data")
    Set rngAll = shtData.Range("A1").CurrentRegion
    Set shtPedigree = ThisWorkbook.Worksheets("Pedigree")

    ' 기존 Pedigree 시트의 모든 도형 고속 삭제
    Dim i As Long
    For i = shtPedigree.Shapes.Count To 1 Step -1
        shtPedigree.Shapes(i).Delete
    Next i

    ' 기본 치수 및 스타일 초기화
    intRows = rngAll.Rows.Count
    intCols = rngAll.Columns.Count
    intBoxW = 24
    intboxH = 26
    intLeft = 36
    intInter = 12
    intBottom = 570
    lngLW = 0.6
    lngTLW = 0.6
    strColor = vbBlack
    blnStart = False
    blnEnd = False
    intN = 0

    ' 1단계: 하단 '후(后)'(13행) 및 '파(派)'(12행) 박스 작도
    For intRow = intRows To intRows - 1 Step -1
        intBoxW = 24
        If intRow = intRows - 1 Then
            intboxH = 90
            Textx = "파"
            intS = 10
            intColor = 41
        ElseIf intRow = intRows Then
            intboxH = 25
            Textx = "후"
            intS = 10
            intColor = 44
        Else
            intboxH = 20
        End If

        intBottom = intBottom - intboxH - 2
        startPL1 = 0
        startPT1 = intBottom

        If intRow = 2 Then
            TextOrient = msoTextOrientationVerticalFarEast
        Else
            TextOrient = msoTextOrientationHorizontal
        End If

        ' 좌측 '후' / '파' 헤더 배지 생성 (정중앙 정렬)
        With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, intBoxW, intboxH)
            .Line.Weight = lngTLW
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame.Characters.Text = Textx
            .TextFrame.Characters.Font.Size = intS
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame2.Orientation = TextOrient
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.MarginLeft = 0
            .TextFrame2.MarginRight = 0
            .TextFrame2.MarginBottom = 0
            .TextFrame2.MarginTop = 0
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Spacing = 0
            .Fill.ForeColor.RGB = RGB(51, 153, 255)
            .TextFrame.Characters.Font.Color = vbWhite
            .Adjustments.Item(1) = 0
        End With

        ' 각 열(1~24열) 순회하며 '후' 병합 박스 및 '파' 세로 박스 생성
        For intCol = 1 To intCols
            intBoxW = 16
            If shtData.Cells(intRow, intCol).Value <> "X" Then
                If Not blnStart And Not blnEnd Then
                    startPL1 = intLeft + (intCol - 1) * (intBoxW + intInter)
                    startPT1 = intBottom
                    blnStart = True
                End If

                If intCol = 1 Then
                    If intRow = intRows Then
                        If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                            endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                            endPT1 = intBottom
                            blnEnd = True
                        End If
                    Else
                        endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                        endPT1 = intBottom + intboxH
                        blnEnd = True
                    End If
                Else
                    If intRow = intRows Then
                        If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol - 1).Value Then
                            If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                                endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                                endPT1 = intBottom
                                blnEnd = True
                            End If
                        Else
                            If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                                endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                                endPT1 = intBottom
                                blnEnd = True
                            End If
                        End If
                    Else
                        endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                        endPT1 = intBottom
                        blnEnd = True
                    End If
                End If

                If intRow = intRows - 2 Then
                    startPT1 = intBottom
                End If

                If blnEnd Then
                    If intRow = intRows Then
                        TextOrient = msoTextOrientationHorizontal
                    ElseIf intRow = intRows - 1 Then
                        TextOrient = msoTextOrientationVerticalFarEast
                    End If

                    With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, endPL1 - startPL1, intboxH)
                        .Line.Weight = lngTLW
                        .Line.ForeColor.RGB = RGB(0, 0, 0)
                        .TextFrame.Characters.Text = CStr(shtData.Cells(intRow, intCol).Value)
                        .TextFrame.Characters.Font.Size = 9
                        .TextFrame.Characters.Font.Color = vbBlack
                        .TextFrame.MarginLeft = 0#
                        .TextFrame.MarginRight = 0#
                        .TextFrame.MarginTop = 0#
                        .TextFrame.MarginBottom = 0#
                        .TextFrame.HorizontalAlignment = xlHAlignCenter
                        .TextFrame.VerticalAlignment = xlVAlignCenter
                        .TextFrame2.Orientation = TextOrient
                        .TextFrame2.HorizontalAnchor = msoAnchorCenter
                        .TextFrame2.VerticalAnchor = msoAnchorMiddle
                        .TextFrame2.MarginLeft = 0
                        .TextFrame2.MarginRight = 0
                        .TextFrame2.MarginBottom = 0
                        .TextFrame2.MarginTop = 0
                        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                        .Adjustments.Item(1) = 0
                        .TextFrame2.TextRange.Font.Spacing = 0
                        .Fill.ForeColor.RGB = RGB(51, 153, 255)
                        .TextFrame.Characters.Font.Color = vbBlack
                        .TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 0.7
                    End With
                    blnEnd = False
                    blnStart = False
                End If
            End If
        Next intCol
    Next intRow

    ' 2단계: 24세(11행) ~ 14세 중시조(1행) 계보 박스 및 연결선 작도
    intBottom = startPT1
    Call PedigreeB(intBottom)

    ' 3단계: 파별·후별 색상 입히기
    Call RectangleColor

    ' 4단계: 세거지 표 및 하단 주석 복사·배치
    Call TextBoxCopy

    ' 5단계: 우측 끝 '위위주부공(휘창)후' 세로 병합 박스 추가
    Call AddBox

    ' 6단계: 중시조 및 4개 파 별칭(설서공파·휴명공파·동지공파·밀양사포파) + 빨간 화살표 배치
    Call AddFiveTextBoxesWithUpArrows

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "분파도 생성 중 오류가 발생했습니다: " & Err.Description, vbExclamation, "오류"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------
' [2단계] 14세(중시조 휘 수) ~ 24세(파시조) 계보 트리 및 상하좌우 연결선 작도
' -----------------------------------------------------------------------------
Public Sub PedigreeB(ByVal intBottom As Long)
    Dim shpPasted As Shape

    intBottom = intBottom - 2
    intRows = rngAll.Rows.Count - 2   ' 11행(24세) ~ 1행(14세)
    intCols = rngAll.Columns.Count
    intBoxW = 24
    intboxH = 24
    intLineX = 3
    intLineY = 3
    lngTLW = 0.5
    ReDim varLineX(0 To intCols + 5)

    For intRow = intRows To 1 Step -1
        intBoxW = 26
        intL = 0
        intT = 0
        startPL1 = 0
        startPT1 = intBottom - (intRows - intRow + 1) * (intLineX + intLineY + intboxH) + intLineX + intLineY
        Textx = (24 - (11 - intRow)) & "세"
        TextOrient = msoTextOrientationVerticalFarEast
        intS = 9

        ' 좌측 세대 표시 박스 (14세 ~ 24세, 정중앙 정렬)
        With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, intBoxW, intboxH)
            .Line.Weight = lngTLW
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame.Characters.Text = Textx
            .TextFrame.Characters.Font.Name = "한양해서"
            .TextFrame.Characters.Font.Size = intS
            .TextFrame.MarginTop = 0#
            .TextFrame.MarginBottom = 0#
            .TextFrame.MarginLeft = 0#
            .TextFrame.MarginRight = 0#
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame2.MarginLeft = 0#
            .TextFrame2.MarginRight = 0#
            .TextFrame2.MarginTop = 0#
            .TextFrame2.MarginBottom = 0#
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1.25
            .TextFrame2.Orientation = msoTextOrientationHorizontal
            .Fill.ForeColor.RGB = RGB(51, 153, 255)
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Color = vbWhite
            .TextFrame2.TextRange.Font.Spacing = 0
            .Adjustments.Item(1) = 0.4
        End With

        For intCol = 1 To intCols
            intBoxW = 16
            If intRow = 1 Then
                startPT1 = intBottom - (intRows - intRow) * (intLineX + intLineY) - (intRows + 1 - intRow) * intboxH
                GoTo DrawRow1
            End If

            ' 셀 값이 "X"인 경우 수직 연장선만 작도
            If shtData.Cells(intRow, intCol).Value = "X" Then
                startPL1 = intLeft + (intCol - 1) * (intBoxW + intInter) + intBoxW / 2
                startPT1 = intBottom - (intRows - intRow) * (intLineX + intLineY + intboxH)
                With shtPedigree.Shapes.AddLine(startPL1, startPT1, startPL1, startPT1 - intLineX - intLineY - intboxH)
                    .Line.Weight = lngLW
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    intT = intT + 1
                    intL = intL + 1
                    If intRow = intRows Then ReDim Preserve varLineX(intL)
                    varLineX(intL) = .Left
                End With
            End If

            ' 셀 값이 "X"가 아닌 경우 인물 함자 박스 및 연결선 작도
            If shtData.Cells(intRow, intCol).Value <> "X" Then
                If intCol = 1 Then
                    startPL1 = intLeft + (intCol - 1) * (intInter + intBoxW)
                    startPT1 = intBottom - (intRows - intRow) * (intLineX + intLineY + intboxH) - intboxH

                    If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                        endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                        endPT1 = startPT1
                        If intRow <> intRows And shtData.Cells(intRow + 1, intCol).Value <> "X" Then
                            intT = intT + 1
                            startPL1 = varLineX(intT) - intBoxW / 2
                        End If

                        ' 함자 박스 생성 (정중앙 정렬)
                        Call AddPersonNode(startPL1, startPT1, intBoxW, intboxH, CStr(shtData.Cells(intRow, intCol).Value), TextOrient)

                        intN = intN + 1
                        With shtPedigree.Shapes.AddLine(startPL1 + (endPL1 - startPL1) / 2, startPT1 - intLineY, _
                                                        startPL1 + (endPL1 - startPL1) / 2, startPT1)
                            .Line.Weight = lngLW
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            startPL2 = .Left
                            startPT2 = .Top
                        End With
                    End If

                    If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol + 1).Value Then
                        endPL2 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                        endPT2 = intBottom - (intRow - 1) * (intLineX + intLineY) + intRow * intboxH + intLineY
                        With shtPedigree.Shapes.AddLine(startPL2, startPT2, startPL2, startPT2 - intLineY)
                            .Line.Weight = lngLW
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            If intRow = intRows Then
                                intL = intL + 1
                                ReDim Preserve varLineX(intL)
                                varLineX(intL) = .Left
                            Else
                                intT = intT + 1
                                varLineX(intT) = .Left
                            End If
                        End With
                    End If
                Else
                    ' 첫 칸이 아닌 경우
                    If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol - 1).Value Then
                        startPL1 = intLeft + (intCol - 1) * (intBoxW + intInter)
                        startPT1 = intBottom - (intRows - intRow) * (intLineX + intLineY + intboxH) - intboxH

                        If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                            endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                            endPT1 = startPT1
                            If intRow <> intRows And shtData.Cells(intRow + 1, intCol).Value <> "" Then
                                intT = intT + 1
                                startPL1 = varLineX(intT) - intBoxW / 2
                            End If

                            ' 함자 박스 생성 (정중앙 정렬)
                            Call AddPersonNode(startPL1, startPT1, intBoxW, intboxH, CStr(shtData.Cells(intRow, intCol).Value), TextOrient)

                            intN = intN + 1
                            With shtPedigree.Shapes.AddLine(startPL1 + (endPL1 - startPL1) / 2, startPT1 - intLineY, _
                                                            startPL1 + (endPL1 - startPL1) / 2, startPT1)
                                .Line.Weight = lngLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                                If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol + 1).Value Then
                                    endPL2 = .Left
                                    endPT2 = .Top
                                End If
                                If intN = 1 Then
                                    startPL2 = .Left
                                    startPT2 = .Top
                                End If
                            End With
                        End If
                    Else
                        ' 앞 칸과 내용이 같고 뒤 칸과 내용이 다를 때 (형제 그룹의 마지막 열)
                        If shtData.Cells(intRow, intCol).Value <> shtData.Cells(intRow, intCol + 1).Value Then
                            endPL1 = intLeft + (intCol - 1) * (intInter + intBoxW) + intBoxW
                            endPT1 = intBottom - (intRows - intRow) * (intLineX + intLineY) - (intRows - intRow) * intboxH
DrawRow1:
                            If intRow = 1 Then
                                intBoxW = intBoxW * 1.5
                                lngTLW = 1.2
                            End If
                            If intRow <> intRows And shtData.Cells(intRow + 1, intCol).Value <> "X" Then
                                intT = intT + 1
                                startPL1 = varLineX(intT) - intBoxW / 2
                            End If

                            ' 함자 박스 생성 (정중앙 정렬: MarginLeft=0, msoAnchorCenter, msoAlignCenter)
                            With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, intBoxW, intboxH)
                                .Line.Weight = lngTLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                                .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent4
                                .Fill.ForeColor.Brightness = 0.8000000119
                                .TextFrame.Characters.Text = CStr(shtData.Cells(intRow, intCol).Value)
                                .TextFrame.Characters.Font.Size = 9
                                .TextFrame.MarginTop = 0#
                                .TextFrame.MarginBottom = 0#
                                .TextFrame.MarginLeft = 0#
                                .TextFrame.MarginRight = 0#
                                .TextFrame.HorizontalAlignment = xlHAlignCenter
                                .TextFrame.VerticalAlignment = xlVAlignCenter
                                .TextFrame.Characters.Font.Color = vbBlack
                                .TextFrame2.MarginTop = 0#
                                .TextFrame2.MarginBottom = 0#
                                .TextFrame2.MarginLeft = 0#
                                .TextFrame2.MarginRight = 0#
                                .TextFrame2.HorizontalAnchor = msoAnchorCenter
                                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                .TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1.25
                                .TextFrame2.Orientation = TextOrient
                                .Fill.Transparency = 0#
                                .Adjustments.Item(1) = 0.4

                                ' 14세 중시조(綏) 박스인 경우 (정중앙 정렬 유지)
                                If intRow = 1 Then
                                    .TextFrame.Characters.Font.Size = 13
                                    .TextFrame.Characters.Font.Bold = True
                                    .TextFrame.Characters.Font.Color = vbWhite
                                    .Fill.ForeColor.RGB = RGB(0, 0, 255)
                                    .TextFrame.MarginLeft = 0#
                                    .TextFrame2.MarginLeft = 0#
                                    .TextFrame.HorizontalAlignment = xlHAlignCenter
                                    .TextFrame.VerticalAlignment = xlVAlignCenter
                                    .TextFrame2.HorizontalAnchor = msoAnchorCenter
                                    .TextFrame2.VerticalAnchor = msoAnchorMiddle
                                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

                                    Dim rootLeft As Single, rootTop As Single, rootW As Single, rootH As Single
                                    rootLeft = .Left
                                    rootTop = .Top
                                    rootW = .Width
                                    rootH = .Height

                                    ' 상단 타이틀(廣州安氏 分派圖) 복사
                                    shtData.Shapes("PicTitle").Copy
                                    DoEvents
                                    shtPedigree.Paste
                                    Set shpPasted = shtPedigree.Shapes(shtPedigree.Shapes.Count)
                                    shpPasted.Name = "PicTitle"
                                    shpPasted.Left = rootLeft - rootW * 3.35
                                    shpPasted.Top = 8

                                    ' 1세~14세 상단 선대 계보(PicPre) 복사
                                    shtData.Shapes("PicPre").Copy
                                    DoEvents
                                    shtPedigree.Paste
                                    Set shpPasted = shtPedigree.Shapes(shtPedigree.Shapes.Count)
                                    shpPasted.Name = "PicPre"
                                    shpPasted.Left = rootLeft - 218
                                    shpPasted.Top = rootTop - rootH * 2.55
                                    Application.CutCopyMode = False
                                    Exit Sub
                                End If
                            End With

                            intN = intN + 1
                            With shtPedigree.Shapes.AddLine(startPL1 + intBoxW / 2, startPT1 - intLineY, _
                                                            startPL1 + intBoxW / 2, startPT1)
                                .Line.Weight = lngLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                                If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol + 1).Value Then
                                    endPL2 = .Left
                                    endPT2 = .Top
                                End If
                                If intN = 1 Then
                                    startPL2 = .Left
                                    startPT2 = .Top
                                End If
                            End With
                        End If
                    End If

                    ' 상위 행(부친 세대)과의 가로·세로 연결선 작도
                    If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol - 1).Value Then
                        If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol + 1).Value Then
                            With shtPedigree.Shapes.AddLine(startPL1 + intBoxW / 2, endPT2, _
                                                            startPL1 + intBoxW / 2, startPT2 - intLineX)
                                .Line.Weight = lngLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                                intL = intL + 1
                                If intRow = intRows Then ReDim Preserve varLineX(intL)
                                varLineX(intL) = .Left
                            End With
                            intN = 0
                        End If
                    Else
                        If shtData.Cells(intRow - 1, intCol).Value <> shtData.Cells(intRow - 1, intCol + 1).Value Then
                            With shtPedigree.Shapes.AddLine(startPL2 + (endPL2 - startPL2) / 2, startPT2, _
                                                            startPL2 + (endPL2 - startPL2) / 2, endPT2 - intLineY)
                                .Line.Weight = lngLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                                intL = intL + 1
                                If intRow = intRows Then ReDim Preserve varLineX(intL)
                                varLineX(intL) = .Left
                            End With
                            With shtPedigree.Shapes.AddLine(startPL2, startPT2, endPL2, endPT2)
                                .Line.Weight = lngLW
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
                            End With
                            intN = 0
                        End If
                    End If
                End If
            End If
        Next intCol
    Next intRow
End Sub

' -----------------------------------------------------------------------------
' [보조 함수] 개별 인물 둥근 사각형 노드 생성 (함자 가로·세로 정중앙 정렬)
' -----------------------------------------------------------------------------
Private Sub AddPersonNode(ByVal pLeft As Single, ByVal pTop As Single, ByVal pW As Single, ByVal pH As Single, _
                          ByVal pText As String, ByVal pOrient As MsoTextOrientation)
    With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, pLeft, pTop, pW, pH)
        .Line.Weight = lngTLW
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .Fill.ForeColor.Brightness = 0.8000000119
        .TextFrame.Characters.Text = pText
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.MarginTop = 0#
        .TextFrame.MarginBottom = 0#
        .TextFrame.MarginLeft = 0#
        .TextFrame.MarginRight = 0#
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Color = vbBlack
        .TextFrame2.MarginTop = 0#
        .TextFrame2.MarginBottom = 0#
        .TextFrame2.MarginLeft = 0#
        .TextFrame2.MarginRight = 0#
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1.25
        .TextFrame2.Orientation = pOrient
        .Fill.Transparency = 0#
        .Adjustments.Item(1) = 0.4
    End With
End Sub

' -----------------------------------------------------------------------------
' [3단계] 후(后)·파(派)·파시조 박스 색상 구분
' -----------------------------------------------------------------------------
Public Sub RectangleColor()
    Dim ShapeE As Shape
    Dim N As Integer
    Dim shapeTxt As String

    Set shtPedigree = ThisWorkbook.Worksheets("Pedigree")

    ' 특정 중시조·파시조 함자별 강조 색상 지정
    For Each ShapeE In shtPedigree.Shapes
        If ShapeE.AutoShapeType = msoShapeRoundedRectangle Then
            shapeTxt = ShapeE.TextFrame2.TextRange.Characters.Text
            Select Case shapeTxt
                Case "鐵關", "鐵山", "從生", "淑仁", "淑信"
                    ShapeE.Fill.ForeColor.RGB = RGB(250, 236, 197)
                Case "禎", "省", "冲"
                    ShapeE.Fill.ForeColor.RGB = RGB(242, 203, 97)
                Case "國柱", "昌", "夢得"
                    ShapeE.Fill.ForeColor.RGB = RGB(188, 229, 92)
                Case "宙"
                    ShapeE.Fill.ForeColor.RGB = RGB(228, 247, 186)
            End Select
        End If
    Next ShapeE

    ' 생성 순서에 따른 하단 후(后)·파(派)·24세 파시조 박스 배색
    N = 0
    For Each ShapeE In shtPedigree.Shapes
        If ShapeE.AutoShapeType = msoShapeRoundedRectangle Then
            shapeTxt = ShapeE.TextFrame2.TextRange.Characters.Text
            If shapeTxt <> "후" And shapeTxt <> "파" And shapeTxt <> "24세" Then
                N = N + 1
                If N < 7 Then
                    If N Mod 2 = 0 Then
                        ShapeE.Fill.ForeColor.RGB = RGB(188, 229, 92)
                    Else
                        ShapeE.Fill.ForeColor.RGB = RGB(242, 203, 97)
                    End If
                ElseIf N < 10 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(250, 236, 197)
                ElseIf N < 20 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(228, 247, 186)
                ElseIf N < 23 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(250, 236, 197)
                ElseIf N < 28 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(228, 247, 186)
                ElseIf N < 30 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(250, 236, 197)
                ElseIf N < 31 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(228, 247, 186)
                ElseIf N < 34 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(250, 236, 197)
                ElseIf N < 48 Then
                    ShapeE.Fill.ForeColor.RGB = RGB(228, 247, 186)
                End If
            End If
        End If
    Next ShapeE
End Sub

' -----------------------------------------------------------------------------
' [4단계] 하단 '분파도 세거지' 표 이미지 및 '분파도 주' 텍스트박스 복사·배치
' -----------------------------------------------------------------------------
Public Sub TextBoxCopy()
    Dim shp As Shape

    Set shtData = ThisWorkbook.Worksheets("data")
    Set shtPedigree = ThisWorkbook.Worksheets("Pedigree")

    ' 1. 분파도 세거지 복사
    shtData.Shapes("분파도 세거지").Copy
    DoEvents
    shtPedigree.Paste
    Set shp = shtPedigree.Shapes(shtPedigree.Shapes.Count)
    shp.Name = "분파도 세거지_복사본"
    shp.Left = 30
    shp.Top = 610
    shp.Height = 440
    Application.CutCopyMode = False

    ' 2. 분파도 주(해설) 복사
    shtData.Shapes("분파도 주").Copy
    DoEvents
    shtPedigree.Paste
    Set shp = shtPedigree.Shapes(shtPedigree.Shapes.Count)
    shp.Name = "분파도 주_복사본"
    shp.Left = 38
    shp.Top = 1059
    shp.Height = 170
    Application.CutCopyMode = False
End Sub

' -----------------------------------------------------------------------------
' [5단계] 우측 끝 24열 '위위주부공(휘창)후' 세로 병합 박스 추가 (중앙 정렬)
' -----------------------------------------------------------------------------
Public Sub AddBox()
    Dim ShapeE As Shape

    Set shtPedigree = ThisWorkbook.Worksheets("Pedigree")
    intM = 0
    intN = 0
    intboxH = 0

    For Each ShapeE In shtPedigree.Shapes
        If ShapeE.AutoShapeType = msoShapeRoundedRectangle Then
            intM = intM + 1
            If intM = 7 Then
                ShapeE.TextFrame.Characters.Font.Color = RGB(242, 203, 97)
                intN = ShapeE.Top
                intboxH = ShapeE.Height - 1
                startPL1 = ShapeE.Left + ShapeE.Width - 1
                startPT1 = ShapeE.Top + 0.8
                intBoxW = 4
                With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, intBoxW, intboxH)
                    .Line.Visible = msoFalse
                    .Fill.ForeColor.RGB = RGB(188, 229, 92)
                    .TextFrame.Characters.Font.Size = 9
                    .ZOrder msoBringToFront
                    .Adjustments.Item(1) = 0
                End With
            End If

            If intM = 32 Then
                intBoxW = 16
                startPL1 = ShapeE.Left + intBoxW
                startPT1 = ShapeE.Top
                intboxH = intN - ShapeE.Top + intboxH + 1

                With shtPedigree.Shapes.AddShape(msoShapeRoundedRectangle, startPL1, startPT1, intBoxW, intboxH)
                    .Line.Weight = 0.3
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .Fill.ForeColor.RGB = RGB(188, 229, 92)
                    .TextFrame.Characters.Text = "위위주부공(휘창)후"
                    .TextFrame.Characters.Font.Size = 9
                    .TextFrame.MarginTop = 0#
                    .TextFrame.MarginBottom = 0#
                    .TextFrame.MarginLeft = 0#
                    .TextFrame.MarginRight = 0#
                    .TextFrame.HorizontalAlignment = xlHAlignCenter
                    .TextFrame.VerticalAlignment = xlVAlignCenter
                    .TextFrame.Characters.Font.Color = vbBlack
                    .TextFrame2.MarginTop = 0#
                    .TextFrame2.MarginBottom = 0#
                    .TextFrame2.MarginLeft = 0#
                    .TextFrame2.MarginRight = 0#
                    .TextFrame2.HorizontalAnchor = msoAnchorCenter
                    .TextFrame2.VerticalAnchor = msoAnchorMiddle
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    .TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1
                    .TextFrame2.Orientation = msoTextOrientationVerticalFarEast
                    .ZOrder msoSendToBack
                    .Fill.Transparency = 0#
                    .Adjustments.Item(1) = 0
                End With
                Exit For
            End If
        End If
    Next ShapeE
End Sub

' -----------------------------------------------------------------------------
' [6단계] 중시조 표기 및 4개 파 별칭 + 빨간 화살표 배치
' -----------------------------------------------------------------------------
Public Sub AddFiveTextBoxesWithUpArrows()
    Set shtPedigree = ThisWorkbook.Worksheets("Pedigree")

    ' 1. 14세 휘 수(綏) 하단: 中始祖
    Call CreateLabelAndArrow("중시조_박스", "中始祖", 319#, 148.2, 120#, 30#, False, 0#, 0#, 0#)

    ' 2. 6열 완귀공파(玩龜公派) 하단: 一名 : 說書公派
    Call CreateLabelAndArrow("설서공파_박스", "一名 : 說書公派", 134.3, 563.8, 120#, 30#, True, 183#, 569.5, 532.5)

    ' 3. 20열 사헌공파(使憲公派) 하단: 一名 : 休命公派
    Call CreateLabelAndArrow("휴명공파_박스", "一名 : 休命公派", 509.9, 563.1, 120#, 30#, True, 572.5, 571.9, 534.9)

    ' 4. 21열 교리공파(校理公派) 하단: 一名 : 同旨公派
    Call CreateLabelAndArrow("동지공파_박스", "一名 : 同旨公派", 556.3, 585.6, 120#, 30#, True, 604.4, 586.6, 532.8)

    ' 5. 24열 위위주부공파(衛尉注簿公派) 하단: 一名 : 密陽沙浦派
    Call CreateLabelAndArrow("밀양사포파_박스", "一名 : 密陽沙浦派", 619.5, 564.2, 120#, 30#, True, 687.7, 573.6, 536.6)

    ' 6. 우측 상단 제작자 표기 텍스트박스
    With shtPedigree.Shapes.AddTextbox(msoTextOrientationHorizontal, 494#, 20.7, 216.6, 25.7)
        .Name = "작성자_박스"
        .TextFrame.Characters.Text = "by 安東洙 翫龜公派 010 5097 6403"
        .TextFrame.Characters.Font.Size = 9
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With
End Sub

' -----------------------------------------------------------------------------
' [보조 함수] 투명 라벨 텍스트박스 및 상향 빨간색 화살표 생성
' -----------------------------------------------------------------------------
Private Sub CreateLabelAndArrow(ByVal boxName As String, ByVal labelText As String, _
                                ByVal boxLeft As Single, ByVal boxTop As Single, _
                                ByVal boxWidth As Single, ByVal boxHeight As Single, _
                                ByVal hasArrow As Boolean, ByVal arrowX As Single, _
                                ByVal arrowBeginY As Single, ByVal arrowEndY As Single)
    Set shp = shtPedigree.Shapes.AddTextbox(msoTextOrientationHorizontal, boxLeft, boxTop, boxWidth, boxHeight)
    With shp.TextFrame.Characters
        .Text = labelText
        .Font.Size = 11
    End With
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    shp.Name = boxName

    If hasArrow Then
        Set arrow = shtPedigree.Shapes.AddLine(BeginX:=arrowX, BeginY:=arrowBeginY, EndX:=arrowX, EndY:=arrowEndY)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)
        arrow.Line.EndArrowheadStyle = msoArrowheadTriangle
    End If
End Sub
