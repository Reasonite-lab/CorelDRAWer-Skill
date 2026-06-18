' =============================================================================
' 钻孔柱状图 - 专业版 (CorelDRAW VBA)
' 生成日期：2026-06-18
' 说明：数据驱动模板，修改 DATA DEFINITION 区域的数组即可改变地层
' =============================================================================

Option Explicit

' =============================================================================
' ★ DATA DEFINITION — 修改这里即可改变地层！
'   每行格式：名称, 顶深(m), 底深(m), C, M, Y, K, 花纹类型, 文字颜色(1=黑,0=白)
'   花纹类型: "pure"纯色 "dots"粗点 "hdots"细密点 "hlines"横线
'             "diag"斜线 "cross"交叉线 "brick"砖格 "chaos"杂乱线
' =============================================================================
Public Type LayerData
    Name       As String   ' 地层名称
    TopDepth   As Double   ' 层顶深度 (m)
    BottomDepth As Double  ' 层底深度 (m)
    C          As Long     ' 填充色 C (0-100)
    M          As Long     ' 填充色 M
    Y_         As Long     ' 填充色 Y (Y 是保留字，用 Y_)
    K          As Long     ' 填充色 K
    Pattern    As String   ' 花纹类型
    TextWhite  As Boolean  ' 文字用白色（深底色用）
End Type

Public Function GetLayers() As LayerData()
    Dim arr(1 To 6) As LayerData
    
    ' --- 第 1 层：杂填土，0-2m ---
    arr(1).Name = "① 杂填土"
    arr(1).TopDepth = 0
    arr(1).BottomDepth = 2
    arr(1).C = 0
    arr(1).M = 10
    arr(1).Y_ = 30
    arr(1).K = 35
    arr(1).Pattern = "chaos"
    arr(1).TextWhite = True
    
    ' --- 第 2 层：粉质粘土，2-6m ---
    arr(2).Name = "② 粉质粘土"
    arr(2).TopDepth = 2
    arr(2).BottomDepth = 6
    arr(2).C = 0
    arr(2).M = 15
    arr(2).Y_ = 50
    arr(2).K = 15
    arr(2).Pattern = "pure"
    arr(2).TextWhite = False
    
    ' --- 第 3 层：粉细砂，6-12m ---
    arr(3).Name = "③ 粉细砂"
    arr(3).TopDepth = 6
    arr(3).BottomDepth = 12
    arr(3).C = 8
    arr(3).M = 0
    arr(3).Y_ = 15
    arr(3).K = 5
    arr(3).Pattern = "hdots"
    arr(3).TextWhite = False
    
    ' --- 第 4 层：砂质泥岩，12-19m ---
    arr(4).Name = "④ 砂质泥岩"
    arr(4).TopDepth = 12
    arr(4).BottomDepth = 19
    arr(4).C = 10
    arr(4).M = 50
    arr(4).Y_ = 40
    arr(4).K = 20
    arr(4).Pattern = "hlines"
    arr(4).TextWhite = True
    
    ' --- 第 5 层：砂岩，19-28m ---
    arr(5).Name = "⑤ 砂岩"
    arr(5).TopDepth = 19
    arr(5).BottomDepth = 28
    arr(5).C = 20
    arr(5).M = 0
    arr(5).Y_ = 30
    arr(5).K = 10
    arr(5).Pattern = "dots"
    arr(5).TextWhite = False
    
    ' --- 第 6 层：石灰岩，28-35m ---
    arr(6).Name = "⑥ 石灰岩"
    arr(6).TopDepth = 28
    arr(6).BottomDepth = 35
    arr(6).C = 0
    arr(6).M = 0
    arr(6).Y_ = 5
    arr(6).K = 8
    arr(6).Pattern = "brick"
    arr(6).TextWhite = False
    
    GetLayers = arr
End Function

' =============================================================================
' ★ LAYOUT CONSTANTS — 调整布局
' =============================================================================
Private Const COL_LEFT      As Double = 45    ' 柱体左 X (mm)
Private Const COL_WIDTH     As Double = 42    ' 柱体宽度
Private Const COL_TOP       As Double = 268   ' 柱体顶 Y (mm，页面顶部)
Private Const COL_BOTTOM    As Double = 25    ' 柱体底 Y
Private Const SCALE_LEFT    As Double = 18    ' 深度标尺左 X
Private Const SCALE_WIDTH   As Double = 22    ' 标尺宽度
Private Const LABEL_LEFT    As Double = 92    ' 标签左 X
Private Const LABEL_WIDTH   As Double = 55    ' 标签区宽
Private Const LEGEND_LEFT   As Double = 155   ' 图例左 X
Private Const LEGEND_WIDTH  As Double = 40    ' 图例框宽
Private Const LEGEND_TOP    As Double = 268   ' 图例顶 Y
Private Const TITLE_Y       As Double = 280   ' 标题 Y
Private Const LEGEND_ITEM_H As Double = 7     ' 每个图例项高
Private Const PAT_SPACING   As Double = 3.5   ' 花纹间距 (mm)

' =============================================================================
' ★ 辅助：创建花纹并 PowerClip 到容器
' =============================================================================
Private Sub ApplyPattern(ByRef container As Shape, ByVal patType As String, _
                         ByVal leftX As Double, ByVal bottomY As Double, _
                         ByVal w As Double, ByVal h As Double)
    If patType = "pure" Then Exit Sub
    
    ' 保存用户当前选中，避免干扰
    Dim savedSel As ShapeRange
    Set savedSel = ActiveSelection
    
    ActiveSelection.Clear
    
    Dim x As Double, y As Double
    Dim s As Shape
    Dim sp As Double
    Dim patternCount As Long
    
    sp = PAT_SPACING  ' 间距
    patternCount = 0
    
    Select Case patType
        ' ---------- 粗点 ----------
        Case "dots"
            x = leftX + sp / 2
            Do While x < leftX + w
                y = bottomY + sp / 2
                Do While y < bottomY + h
                    ' 用 CreateEllipse(外接矩形) 替代 CreateEllipse2（兼容旧版）
                    Set s = ActiveLayer.CreateEllipse(x - 1.2, y - 1.2, x + 1.2, y + 1.2)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 30
                    s.Outline.Type = cdrNoOutline
                    patternCount = patternCount + 1
                    y = y + sp * 1.2
                Loop
                x = x + sp * 1.2
            Loop
            
        ' ---------- 细密点 ----------
        Case "hdots"
            sp = sp * 0.6
            x = leftX + sp / 2
            Do While x < leftX + w
                y = bottomY + sp / 2
                Do While y < bottomY + h
                    Set s = ActiveLayer.CreateEllipse(x - 0.6, y - 0.6, x + 0.6, y + 0.6)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 25
                    s.Outline.Type = cdrNoOutline
                    patternCount = patternCount + 1
                    y = y + sp
                Loop
                x = x + sp
            Loop
            
        ' ---------- 横线 ----------
        Case "hlines"
            y = bottomY + sp * 0.6
            Do While y < bottomY + h
                Set s = ActiveLayer.CreateLineSegment(leftX + 1.5, y, leftX + w - 1.5, y)
                s.Outline.Color.CMYKAssign 0, 0, 0, 35
                s.Outline.Width = 0.25
                s.Outline.ScaleWithShape = False
                patternCount = patternCount + 1
                y = y + sp
            Loop
            
        ' ---------- 斜线 (45°) ----------
        Case "diag"
            Dim startX As Double, startY As Double, endX As Double, endY As Double
            Dim d As Double
            d = (w + h) / 2
            startX = leftX - d
            Do While startX < leftX + w + d
                ' 从下到上的斜线 (SW→NE)
                If startX < leftX Then
                    startY = bottomY + (leftX - startX)
                Else
                    startY = bottomY
                End If
                endX = startX + d * 1.414
                endY = startY + d * 1.414
                If endY > bottomY + h Then
                    endX = startX + (bottomY + h - startY)
                    endY = bottomY + h
                End If
                If endX > leftX + w Then
                    endX = leftX + w
                    endY = startY + (leftX + w - startX)
                End If
                If endY > bottomY And startY < bottomY + h Then
                    Set s = ActiveLayer.CreateLineSegment(startX, startY, endX, endY)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 30
                    s.Outline.Width = 0.22
                    s.Outline.ScaleWithShape = False
                    patternCount = patternCount + 1
                End If
                startX = startX + sp
            Loop
            
        ' ---------- 交叉线 ----------
        Case "cross"
            Dim startX2 As Double
            ' 45° 线 (NE)
            startX = leftX - (w + h)
            Do While startX < leftX + w + h
                Set s = ActiveLayer.CreateLineSegment(startX, bottomY, _
                    startX + h, bottomY + h)
                s.Outline.Color.CMYKAssign 0, 0, 0, 25
                s.Outline.Width = 0.2
                s.Outline.ScaleWithShape = False
                patternCount = patternCount + 1
                startX = startX + sp * 1.5
            Loop
            ' -45° 线 (SE)
            startX2 = leftX - (w + h)
            Do While startX2 < leftX + w + h
                Set s = ActiveLayer.CreateLineSegment(startX2, bottomY + h, _
                    startX2 + h, bottomY)
                s.Outline.Color.CMYKAssign 0, 0, 0, 25
                s.Outline.Width = 0.2
                s.Outline.ScaleWithShape = False
                patternCount = patternCount + 1
                startX2 = startX2 + sp * 1.5
            Loop
            
        ' ---------- 砖格 ----------
        Case "brick"
            Dim row As Long
            Dim bw As Double, bh As Double
            bw = sp * 1.8: bh = sp * 0.7
            row = 0
            y = bottomY + 0.5
            Do While y < bottomY + h
                x = leftX + 0.5
                If row Mod 2 = 1 Then x = x + bw / 2  ' 错缝
                Do While x < leftX + w
                    Set s = ActiveLayer.CreateRectangle(x, y, x + bw - 0.5, y + bh)
                    s.Fill.Type = cdrNoFill
                    s.Outline.Color.CMYKAssign 0, 0, 0, 20
                    s.Outline.Width = 0.18
                    s.Outline.ScaleWithShape = False
                    patternCount = patternCount + 1
                    x = x + bw
                Loop
                y = y + bh + 0.3
                row = row + 1
            Loop
            
        ' ---------- 杂乱线（填土）----------
        Case "chaos"
            Randomize 42  ' 固定随机种子，保证每次运行结果一致
            Dim i As Long, n As Long
            n = Int((w * h) / (sp * sp) * 1.5)
            For i = 1 To n
                x = leftX + Rnd() * w
                y = bottomY + Rnd() * h
                Dim ex As Double, ey As Double
                ex = x + (Rnd() - 0.5) * sp * 1.8
                ey = y + (Rnd() - 0.5) * sp * 1.8
                Set s = ActiveLayer.CreateLineSegment(x, y, ex, ey)
                s.Outline.Color.CMYKAssign 0, 0, 0, 18
                s.Outline.Width = 0.15 * (0.5 + Rnd())
                s.Outline.ScaleWithShape = False
                patternCount = patternCount + 1
            Next i
    End Select
    
    ' --- PowerClip: 将所有花纹群组后切入容器 ---
    If patternCount > 0 Then
        Dim grp As Shape
        Set grp = ActiveSelection.Group  ' 群组所有花纹图形
        container.PowerClip.PlaceInside grp  ' 将群组切入选中的容器
    End If
    
    ' 恢复用户原来的选中（如有）
    On Error Resume Next
    If Not savedSel Is Nothing Then
        If savedSel.Count > 0 Then
            savedSel.CreateSelection
        End If
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' ★ 主绘图宏
' =============================================================================
Public Sub DrawBorehole()
    
    On Error GoTo ErrHandler
    
    If ActiveDocument Is Nothing Then
        MsgBox "请先打开一个 CorelDRAW 文档", vbExclamation, "钻孔柱状图"
        Exit Sub
    End If
    
    ActiveDocument.BeginCommandGroup "绘制钻孔柱状图"
    
    ' ========== 初始化 ==========
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    Dim layers() As LayerData
    layers = GetLayers()
    
    Dim numLayers As Long
    numLayers = UBound(layers) - LBound(layers) + 1
    
    ' 总深度 & 比例系数
    Dim totalDepth As Double
    totalDepth = layers(numLayers).BottomDepth
    
    Dim colH As Double
    colH = COL_TOP - COL_BOTTOM  ' 可用柱高 (mm)
    
    Dim scaleFactor As Double
    scaleFactor = colH / totalDepth  ' mm per meter
    
    Dim colRight As Double
    colRight = COL_LEFT + COL_WIDTH
    
    ' ========== 1. 标题 ==========
    Dim title As Shape
    Set title = ActiveLayer.CreateArtisticText(COL_LEFT, TITLE_Y, "钻 孔 柱 状 图")
    title.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    ' 设字体大小
    With title.Text.Story
        .Font = "黑体"
        .Size = 14
        .Bold = True
    End With
    ' 居中于柱体
    title.SetPosition COL_LEFT + COL_WIDTH / 2, TITLE_Y
    
    ' 底部分隔线
    Dim underline As Shape
    Set underline = ActiveLayer.CreateLineSegment(COL_LEFT, TITLE_Y - 5, _
        COL_LEFT + COL_WIDTH, TITLE_Y - 5)
    underline.Outline.Color.CMYKAssign 0, 0, 0, 100
    underline.Outline.Width = 0.5
    underline.Outline.ScaleWithShape = False
    
    ' ========== 2. 深度标尺 ==========
    Dim scaleRight As Double
    scaleRight = SCALE_LEFT + SCALE_WIDTH
    
    ' 标尺主体竖线
    Dim scaleLine As Shape
    Set scaleLine = ActiveLayer.CreateLineSegment(scaleRight, COL_BOTTOM, _
        scaleRight, COL_TOP)
    scaleLine.Outline.Color.CMYKAssign 0, 0, 0, 100
    scaleLine.Outline.Width = 0.5
    scaleLine.Outline.ScaleWithShape = False
    
    ' 标尺刻度和数字
    Dim depth As Double, tickY As Double
    Dim tickStep As Double
    tickStep = 5  ' 每 5 米一个刻度
    
    depth = 0
    Do While depth <= totalDepth
        tickY = COL_TOP - depth * scaleFactor
        
        ' 刻度线
        Dim tick As Shape
        Set tick = ActiveLayer.CreateLineSegment(scaleRight - 4, tickY, scaleRight, tickY)
        tick.Outline.Color.CMYKAssign 0, 0, 0, 100
        tick.Outline.Width = 0.3
        tick.Outline.ScaleWithShape = False
        
        ' 深度数字
        Dim tickNum As Shape
        Set tickNum = ActiveLayer.CreateArtisticText(scaleRight - 6, tickY - 1.5, _
            Format(depth, "0"))
        With tickNum.Text.Story
            .Font = "Arial"
            .Size = 7
        End With
        tickNum.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        
        ' 辅助细线（延伸到柱体）
        If depth > 0 Then
            Dim guideLine As Shape
            Set guideLine = ActiveLayer.CreateLineSegment(scaleRight, tickY, COL_LEFT, tickY)
            guideLine.Outline.Color.CMYKAssign 0, 0, 0, 10
            guideLine.Outline.Width = 0.1
            guideLine.Outline.ScaleWithShape = False
        End If
        
        depth = depth + tickStep
    Loop
    
    ' 标尺标签
    Dim scaleLabel As Shape
    Set scaleLabel = ActiveLayer.CreateArtisticText(SCALE_LEFT + 1, _
        COL_TOP + 3, "深度(m)")
    With scaleLabel.Text.Story
        .Font = "黑体"
        .Size = 7
    End With
    scaleLabel.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    
    ' ========== 3. 地层柱体 + 花纹 ==========
    
    ' 左侧边界线
    Dim leftBound As Shape
    Set leftBound = ActiveLayer.CreateLineSegment(COL_LEFT, COL_BOTTOM, COL_LEFT, COL_TOP)
    leftBound.Outline.Color.CMYKAssign 0, 0, 0, 100
    leftBound.Outline.Width = 0.5
    leftBound.Outline.ScaleWithShape = False
    
    ' 右侧边界线
    Dim rightBound As Shape
    Set rightBound = ActiveLayer.CreateLineSegment(colRight, COL_BOTTOM, colRight, COL_TOP)
    rightBound.Outline.Color.CMYKAssign 0, 0, 0, 100
    rightBound.Outline.Width = 0.5
    rightBound.Outline.ScaleWithShape = False
    
    Dim li As Long
    For li = LBound(layers) To UBound(layers)
        Dim lay As LayerData
        lay = layers(li)
        
        ' 计算层的位置
        Dim layTop As Double, layBottom As Double
        layTop = COL_TOP - lay.TopDepth * scaleFactor
        layBottom = COL_TOP - lay.BottomDepth * scaleFactor
        Dim layH As Double
        layH = layTop - layBottom
        
        ' 如果层太低，最小给 5mm
        If layH < 5 Then layH = 5: layBottom = layTop - 5
        
        ' --- 创建层矩形 ---
        Dim rect As Shape
        Set rect = ActiveLayer.CreateRectangle(COL_LEFT, layBottom, colRight, layTop)
        rect.Fill.UniformColor.CMYKAssign lay.C, lay.M, lay.Y_, lay.K
        rect.Outline.Color.CMYKAssign 0, 0, 0, 40
        rect.Outline.Width = 0.2
        rect.Outline.ScaleWithShape = False
        
        ' --- 应用花纹 ---
        ApplyPattern rect, lay.Pattern, COL_LEFT, layBottom, COL_WIDTH, layH
        
        ' --- 层分隔线（顶线加粗）---
        Dim sepLine As Shape
        Set sepLine = ActiveLayer.CreateLineSegment(COL_LEFT, layTop, colRight, layTop)
        sepLine.Outline.Color.CMYKAssign 0, 0, 0, 60
        sepLine.Outline.Width = 0.35
        sepLine.Outline.ScaleWithShape = False
        
        ' --- 地层标签 ---
        Dim lbl As Shape
        Dim lblY As Double
        lblY = layBottom + layH / 2  ' 标签垂直居中于层
        Set lbl = ActiveLayer.CreateArtisticText(LABEL_LEFT, lblY, lay.Name)
        With lbl.Text.Story
            .Font = "黑体"
            .Size = 8
        End With
        If lay.TextWhite Then
            lbl.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
        Else
            lbl.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        End If
        
        ' 层深度标注
        Dim depthLabel As Shape
        Set depthLabel = ActiveLayer.CreateArtisticText(LABEL_LEFT, lblY - 3, _
            Format(lay.TopDepth, "0.0") & "-" & Format(lay.BottomDepth, "0.0") & "m")
        With depthLabel.Text.Story
            .Font = "Arial"
            .Size = 6.5
        End With
        depthLabel.Fill.UniformColor.CMYKAssign 0, 0, 0, 60
        
    Next li
    
    ' ========== 4. 图例 ==========
    ' 图例框
    Dim legendBox As Shape
    Set legendBox = ActiveLayer.CreateRectangle(LEGEND_LEFT, LEGEND_TOP - 8 - numLayers * LEGEND_ITEM_H, _
        LEGEND_LEFT + LEGEND_WIDTH, LEGEND_TOP)
    legendBox.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    legendBox.Outline.Color.CMYKAssign 0, 0, 0, 60
    legendBox.Outline.Width = 0.3
    legendBox.Outline.ScaleWithShape = False
    
    ' 图例标题
    Dim legendTitle As Shape
    Set legendTitle = ActiveLayer.CreateArtisticText(LEGEND_LEFT + 3, LEGEND_TOP - 1, "图 例")
    With legendTitle.Text.Story
        .Font = "黑体"
        .Size = 8
        .Bold = True
    End With
    legendTitle.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    
    ' 各图例项
    ' 收集已出现的花纹类型
    Dim legendItems() As String
    Dim legendColors() As Long  ' CMYK packed
    Dim legendCount As Long
    legendCount = 0
    
    For li = LBound(layers) To UBound(layers)
        Dim alreadyIn As Boolean
        alreadyIn = False
        Dim j As Long
        For j = 0 To legendCount - 1
            If legendItems(j) = layers(li).Pattern Then
                alreadyIn = True
                Exit For
            End If
        Next j
        If Not alreadyIn Then
            ReDim Preserve legendItems(0 To legendCount)
            ReDim Preserve legendColors(0 To legendCount)
            legendItems(legendCount) = layers(li).Pattern
            legendColors(legendCount) = layers(li).C * 1000000 + layers(li).M * 10000 + layers(li).Y_ * 100 + layers(li).K
            legendCount = legendCount + 1
        End If
    Next j
    
    ' 画出图例项
    Dim li2 As Long
    For li2 = 0 To legendCount - 1
        Dim itemY As Double
        itemY = LEGEND_TOP - 8 - li2 * LEGEND_ITEM_H
        
        ' 颜色方块
        Dim itemRect As Shape
        Set itemRect = ActiveLayer.CreateRectangle(LEGEND_LEFT + 4, itemY - LEGEND_ITEM_H + 1.5, _
            LEGEND_LEFT + 11, itemY + 1)
        
        Dim cc As Long, cm As Long, cy_ As Long, ck As Long
        cc = legendColors(li2) \ 1000000
        cm = (legendColors(li2) Mod 1000000) \ 10000
        cy_ = (legendColors(li2) Mod 10000) \ 100
        ck = legendColors(li2) Mod 100
        
        itemRect.Fill.UniformColor.CMYKAssign cc, cm, cy_, ck
        itemRect.Outline.Color.CMYKAssign 0, 0, 0, 40
        itemRect.Outline.Width = 0.15
        itemRect.Outline.ScaleWithShape = False
        
        ' 花纹名
        Dim patName As String
        Select Case legendItems(li2)
            Case "pure":   patName = "纯色"
            Case "dots":   patName = "粗点"
            Case "hdots":  patName = "细砂点"
            Case "hlines": patName = "水平层理"
            Case "diag":   patName = "斜层理"
            Case "cross":  patName = "交错层理"
            Case "brick":  patName = "砖格 / 石灰岩"
            Case "chaos":  patName = "杂乱 / 填土"
            Case Else:     patName = legendItems(li2)
        End Select
        
        Dim patLabel As Shape
        Set patLabel = ActiveLayer.CreateArtisticText(LEGEND_LEFT + 13, itemY - LEGEND_ITEM_H / 2 + 0.5, patName)
        With patLabel.Text.Story
            .Font = "黑体"
            .Size = 6
        End With
        patLabel.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        
        ' 花纹缩略图
        ApplyPattern itemRect, legendItems(li2), LEGEND_LEFT + 4, itemY - LEGEND_ITEM_H + 1.5, 7, LEGEND_ITEM_H - 1.5
    Next li2
    
    ' ========== 5. 底部图签 ==========
    Dim footer As Shape
    Set footer = ActiveLayer.CreateArtisticText(COL_LEFT, COL_BOTTOM - 7, _
        "（数据驱动模板 · 修改 GetLayers() 函数中的数组即可更新地层）")
    With footer.Text.Story
        .Font = "黑体"
        .Size = 6
        .Italic = True
    End With
    footer.Fill.UniformColor.CMYKAssign 0, 0, 0, 40
    
    ' ========== 完成 ==========
    ActiveDocument.EndCommandGroup
    MsgBox "钻孔柱状图绘制完成！" & vbCrLf & vbCrLf & _
           "共 " & numLayers & " 层 · 深度 " & totalDepth & "m", vbInformation, "钻孔柱状图"
    Exit Sub
    
ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "绘图出错: " & Err.Description & " (错误号 " & Err.Number & ")", vbCritical, "钻孔柱状图 — 错误"
End Sub
