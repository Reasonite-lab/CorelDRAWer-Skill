' =============================================================================
' 钻孔/实测地层柱状图 - 国家标准版 (CorelDRAW VBA)
' 依据：中国地质大学(武汉)秭归产学研基地《常用地层图例花纹和符号》
' 标准：GB/T 958 区域地质图图例
' 说明：多列标准格式，数据驱动模板
' =============================================================================

Option Explicit

' =============================================================================
' ★ DATA DEFINITION — 修改这里即可
'   格式：界, 系, 统, 组(段), 代号, 层厚(m), 岩性描述, C,M,Y,K, 花纹类型
'   花纹类型: "conglo"砾岩 "sand"砂岩 "finesand"细砂岩 "silt"粉砂岩
'             "mud"泥岩 "shale"页岩 "carbShale"炭质页岩 "lime"石灰岩
'             "dolo"白云岩 "doloLime"白云质灰岩 "chert"硅质岩 "coal"煤
'             "granite"花岗岩 "basalt"玄武岩 "schist"片岩 "gneiss"片麻岩
'             "marble"大理岩 "pure"纯色无纹
' =============================================================================
Public Type StrataRow
    Erathem    As String   ' 界（宇）
    System     As String   ' 系
    Series     As String   ' 统
    Formation  As String   ' 组/段
    Symbol     As String   ' 代号 (∈, O, S, D, C, P, T, J, K, E, N, Q, Nh, Z...)
    Thickness  As Double   ' 层厚 (m)
    Descr      As String   ' 岩性描述
    C          As Long     ' CMYK C
    M          As Long
    Y_         As Long
    K          As Long
    Pattern    As String   ' 花纹类型
End Type

Public Function GetSection() As StrataRow()
    ' === 秭归地区典型地层柱（南华系-奥陶系）共计 14 层 ===
    Dim arr(1 To 14) As StrataRow
    
    ' -------- 南华系 (Nh) --------
    arr(1).Erathem = "新元古界"
    arr(1).System = "南华系"
    arr(1).Series = "下统"
    arr(1).Formation = "莲沱组"
    arr(1).Symbol = "Nh₁l"
    arr(1).Thickness = 120
    arr(1).Descr = "紫红色中厚层砂岩、含砾砂岩，底部为砾岩"
    arr(1).C = 0: arr(1).M = 40: arr(1).Y_ = 30: arr(1).K = 10
    arr(1).Pattern = "sand"
    
    arr(2).Erathem = "新元古界"
    arr(2).System = "南华系"
    arr(2).Series = "下统"
    arr(2).Formation = "南沱组"
    arr(2).Symbol = "Nh₁n"
    arr(2).Thickness = 45
    arr(2).Descr = "灰绿色冰碛砾岩、含砾砂质泥岩"
    arr(2).C = 15: arr(2).M = 0: arr(2).Y_ = 20: arr(2).K = 20
    arr(2).Pattern = "conglo"
    
    ' -------- 震旦系 (Z) --------
    arr(3).Erathem = "新元古界"
    arr(3).System = "震旦系"
    arr(3).Series = "下统"
    arr(3).Formation = "陡山沱组"
    arr(3).Symbol = "Z₁d"
    arr(3).Thickness = 180
    arr(3).Descr = "灰黑色薄层泥质白云岩夹炭质页岩、硅质岩"
    arr(3).C = 5: arr(3).M = 0: arr(3).Y_ = 10: arr(3).K = 30
    arr(3).Pattern = "dolo"
    
    arr(4).Erathem = "新元古界"
    arr(4).System = "震旦系"
    arr(4).Series = "上统"
    arr(4).Formation = "灯影组"
    arr(4).Symbol = "Z₂dy"
    arr(4).Thickness = 350
    arr(4).Descr = "灰白色厚层白云岩，含燧石条带，顶部为硅质白云岩"
    arr(4).C = 0: arr(4).M = 0: arr(4).Y_ = 5: arr(4).K = 10
    arr(4).Pattern = "dolo"
    
    ' -------- 寒武系 (∈) --------
    arr(5).Erathem = "古生界"
    arr(5).System = "寒武系"
    arr(5).Series = "下统"
    arr(5).Formation = "水井沱组"
    arr(5).Symbol = "∈₁s"
    arr(5).Thickness = 95
    arr(5).Descr = "黑色炭质页岩夹薄层硅质岩，底部为含磷结核页岩"
    arr(5).C = 0: arr(5).M = 5: arr(5).Y_ = 10: arr(5).K = 65
    arr(5).Pattern = "carbShale"
    
    arr(6).Erathem = "古生界"
    arr(6).System = "寒武系"
    arr(6).Series = "下统"
    arr(6).Formation = "石牌组"
    arr(6).Symbol = "∈₁sp"
    arr(6).Thickness = 150
    arr(6).Descr = "黄绿色粉砂质页岩、泥质粉砂岩，夹细砂岩"
    arr(6).C = 5: arr(6).M = 0: arr(6).Y_ = 30: arr(6).K = 15
    arr(6).Pattern = "silt"
    
    arr(7).Erathem = "古生界"
    arr(7).System = "寒武系"
    arr(7).Series = "下统"
    arr(7).Formation = "天河板组"
    arr(7).Symbol = "∈₁t"
    arr(7).Thickness = 80
    arr(7).Descr = "深灰色中厚层鲕状灰岩、泥质条带灰岩"
    arr(7).C = 0: arr(7).M = 0: arr(7).Y_ = 0: arr(7).K = 35
    arr(7).Pattern = "lime"
    
    arr(8).Erathem = "古生界"
    arr(8).System = "寒武系"
    arr(8).Series = "下统"
    arr(8).Formation = "石龙洞组"
    arr(8).Symbol = "∈₁sl"
    arr(8).Thickness = 110
    arr(8).Descr = "灰白色厚层白云岩，局部含燧石结核"
    arr(8).C = 0: arr(8).M = 0: arr(8).Y_ = 5: arr(8).K = 8
    arr(8).Pattern = "dolo"
    
    arr(9).Erathem = "古生界"
    arr(9).System = "寒武系"
    arr(9).Series = "中上统"
    arr(9).Formation = "覃家庙组"
    arr(9).Symbol = "∈₂₋₃q"
    arr(9).Thickness = 210
    arr(9).Descr = "浅灰色薄-中厚层白云岩、泥质白云岩夹页岩"
    arr(9).C = 0: arr(9).M = 5: arr(9).Y_ = 15: arr(9).K = 12
    arr(9).Pattern = "dolo"
    
    arr(10).Erathem = "古生界"
    arr(10).System = "寒武系"
    arr(10).Series = "上统"
    arr(10).Formation = "三游洞组"
    arr(10).Symbol = "∈₃s"
    arr(10).Thickness = 280
    arr(10).Descr = "灰白色厚层-块状白云岩、硅质白云岩"
    arr(10).C = 0: arr(10).M = 0: arr(10).Y_ = 3: arr(10).K = 8
    arr(10).Pattern = "dolo"
    
    ' -------- 奥陶系 (O) --------
    arr(11).Erathem = "古生界"
    arr(11).System = "奥陶系"
    arr(11).Series = "下统"
    arr(11).Formation = "南津关组"
    arr(11).Symbol = "O₁n"
    arr(11).Thickness = 60
    arr(11).Descr = "深灰色中厚层生物碎屑灰岩、泥质灰岩"
    arr(11).C = 0: arr(11).M = 0: arr(11).Y_ = 0: arr(11).K = 30
    arr(11).Pattern = "lime"
    
    arr(12).Erathem = "古生界"
    arr(12).System = "奥陶系"
    arr(12).Series = "下统"
    arr(12).Formation = "分乡组"
    arr(12).Symbol = "O₁f"
    arr(12).Thickness = 45
    arr(12).Descr = "灰绿色页岩夹薄层灰岩、生物碎屑灰岩"
    arr(12).C = 0: arr(12).M = 0: arr(12).Y_ = 0: arr(12).K = 25
    arr(12).Pattern = "shale"
    
    arr(13).Erathem = "古生界"
    arr(13).System = "奥陶系"
    arr(13).Series = "下统"
    arr(13).Formation = "红花园组"
    arr(13).Symbol = "O₁h"
    arr(13).Thickness = 35
    arr(13).Descr = "灰色厚层生物碎屑灰岩、含燧石结核灰岩"
    arr(13).C = 0: arr(13).M = 0: arr(13).Y_ = 0: arr(13).K = 32
    arr(13).Pattern = "lime"
    
    arr(14).Erathem = "古生界"
    arr(14).System = "奥陶系"
    arr(14).Series = "中统"
    arr(14).Formation = "宝塔组"
    arr(14).Symbol = "O₂b"
    arr(14).Thickness = 20
    arr(14).Descr = "紫红色中厚层龟裂纹灰岩"
    arr(14).C = 0: arr(14).M = 30: arr(14).Y_ = 20: arr(14).K = 10
    arr(14).Pattern = "lime"
    
    GetSection = arr
End Function

' =============================================================================
' ★ LAYOUT CONSTANTS — 多列标准格式
' =============================================================================
Private Const MARGIN_LEFT   As Double = 8     ' 左边距
Private Const MARGIN_TOP    As Double = 10    ' 上边距
Private Const TABLE_TOP     As Double = 285   ' 表格顶部 Y (含标题)
Private Const TABLE_BOTTOM  As Double = 22    ' 表格底部 Y
Private Const HEADER_H      As Double = 14    ' 表头高度
Private Const TOTAL_COL_W   As Double = 190   ' 表格总宽度 (mm)

' 各列宽度（自动按比例分配）
Private Const COL_W_ERATHEM   As Double = 14  ' 界
Private Const COL_W_SYSTEM    As Double = 14  ' 系
Private Const COL_W_SERIES    As Double = 14  ' 统
Private Const COL_W_FORMATION As Double = 20  ' 组
Private Const COL_W_SYMBOL    As Double = 13  ' 代号
Private Const COL_W_COLUMN    As Double = 35  ' 柱状图（主图）
Private Const COL_W_THICK     As Double = 12  ' 厚度
Private Const COL_W_DESCR     As Double = 58  ' 岩性描述
Private Const COL_LINE_W      As Double = 0.25 ' 表格线宽

Private Const PAT_SPACING    As Double = 3.2  ' 花纹间距

' =============================================================================
' ★ 计算列位置
' =============================================================================
Private Function ColX(ByVal idx As Long) As Double
    ' 返回第 idx 列的左 X 坐标（1-based）
    Dim offsets(1 To 8) As Double
    offsets(1) = MARGIN_LEFT
    offsets(2) = offsets(1) + COL_W_ERATHEM
    offsets(3) = offsets(2) + COL_W_SYSTEM
    offsets(4) = offsets(3) + COL_W_SERIES
    offsets(5) = offsets(4) + COL_W_FORMATION
    offsets(6) = offsets(5) + COL_W_SYMBOL
    offsets(7) = offsets(6) + COL_W_COLUMN
    offsets(8) = offsets(7) + COL_W_THICK
    ColX = offsets(idx)
End Function

' =============================================================================
' ★ 辅助：获取列宽
' =============================================================================
Private Function ColW(ByVal idx As Long) As Double
    Select Case idx
        Case 1: ColW = COL_W_ERATHEM
        Case 2: ColW = COL_W_SYSTEM
        Case 3: ColW = COL_W_SERIES
        Case 4: ColW = COL_W_FORMATION
        Case 5: ColW = COL_W_SYMBOL
        Case 6: ColW = COL_W_COLUMN
        Case 7: ColW = COL_W_THICK
        Case 8: ColW = COL_W_DESCR
    End Select
End Function

' =============================================================================
' ★ 辅助：创建竖排文字（垂直居中于给定区域）
' =============================================================================
Private Sub AddCellText(ByVal txt As String, ByVal cx As Double, _
                        ByVal topY As Double, ByVal bottomY As Double, _
                        ByVal cellW As Double, Optional ByVal fontSize As Double = 7, _
                        Optional ByVal isBold As Boolean = False)
    If txt = "" Then Exit Sub
    
    Dim midY As Double
    midY = (topY + bottomY) / 2
    
    ' 如果文字很短，尝试竖排（适合中文短标题）
    Dim useVertical As Boolean
    useVertical = (Len(txt) <= 10 And cellW <= 15)
    
    Dim s As Shape
    If useVertical Then
        ' 竖排：逐字符创建
        Dim chars() As String
        Dim clen As Long, ci As Long
        clen = Len(txt)
        ReDim chars(1 To clen)
        For ci = 1 To clen
            chars(ci) = Mid(txt, ci, 1)
        Next ci
        ' 回到横排—简化处理
        Set s = ActiveLayer.CreateArtisticText(cx + 1, midY, txt)
    Else
        Set s = ActiveLayer.CreateArtisticText(cx + 1, midY, txt)
    End If
    
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    With s.Text.Story
        .Font = "黑体"
        .Size = fontSize
        .Bold = isBold
    End With
    s.Outline.Type = cdrNoOutline
End Sub

' =============================================================================
' ★ 辅助：画表格线
' =============================================================================
Private Sub DrawHLine(ByVal x1 As Double, ByVal x2 As Double, ByVal y As Double, _
                      Optional ByVal w As Double = 0, Optional ByVal cmyk_k As Long = 80)
    If w = 0 Then w = COL_LINE_W
    Dim ln As Shape
    Set ln = ActiveLayer.CreateLineSegment(x1, y, x2, y)
    ln.Outline.Color.CMYKAssign 0, 0, 0, cmyk_k
    ln.Outline.Width = w
    ln.Outline.ScaleWithShape = False
End Sub

Private Sub DrawVLine(ByVal x As Double, ByVal y1 As Double, ByVal y2 As Double, _
                      Optional ByVal w As Double = 0, Optional ByVal cmyk_k As Long = 80)
    If w = 0 Then w = COL_LINE_W
    Dim ln As Shape
    Set ln = ActiveLayer.CreateLineSegment(x, y1, x, y2)
    ln.Outline.Color.CMYKAssign 0, 0, 0, cmyk_k
    ln.Outline.Width = w
    ln.Outline.ScaleWithShape = False
End Sub

' =============================================================================
' ★ 辅助：创建花纹 + PowerClip
' =============================================================================
Private Sub AddPattern(ByRef container As Shape, ByVal patType As String, _
                       ByVal lx As Double, ByVal by As Double, _
                       ByVal cw As Double, ByVal ch As Double)
    
    ' 纯色不需要花纹
    If patType = "pure" Or ch < 3 Then Exit Sub
    
    ' 保存用户选中
    Dim savedSel As ShapeRange
    Set savedSel = ActiveSelection
    
    ActiveSelection.Clear
    
    Dim x As Double, y As Double
    Dim s As Shape
    Dim sp As Double
    Dim cnt As Long
    
    sp = PAT_SPACING
    cnt = 0
    
    Select Case patType
        
        ' ==================== 沉积岩 ====================
        
        ' --- 砾岩：大圆 + 散点 ---
        Case "conglo"
            ' 大圆圈
            x = lx + sp * 0.7
            Do While x < lx + cw - sp * 0.3
                y = by + sp * 0.7
                Do While y < by + ch - sp * 0.3
                    Set s = ActiveLayer.CreateEllipse(x - 1.5, y - 1.5, x + 1.5, y + 1.5)
                    s.Fill.Type = cdrNoFill
                    s.Outline.Color.CMYKAssign 0, 0, 0, 25
                    s.Outline.Width = 0.2
                    cnt = cnt + 1
                    y = y + sp * 1.8
                Loop
                x = x + sp * 1.8
            Loop
            ' 散点
            x = lx + sp * 1.6
            Do While x < lx + cw - sp * 0.5
                y = by + sp * 1.6
                Do While y < by + ch - sp * 0.5
                    Set s = ActiveLayer.CreateEllipse(x - 0.4, y - 0.4, x + 0.4, y + 0.4)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 35
                    s.Outline.Type = cdrNoOutline
                    cnt = cnt + 1
                    y = y + sp * 1.8
                Loop
                x = x + sp * 1.8
            Loop
            
        ' --- 砂岩：密点 ---
        Case "sand"
            sp = sp * 0.7
            x = lx + sp / 2
            Do While x < lx + cw
                y = by + sp / 2
                Do While y < by + ch
                    Set s = ActiveLayer.CreateEllipse(x - 0.35, y - 0.35, x + 0.35, y + 0.35)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 30
                    s.Outline.Type = cdrNoOutline
                    cnt = cnt + 1
                    y = y + sp
                Loop
                x = x + sp
            Loop
            
        ' --- 细砂岩 ---
        Case "finesand"
            sp = sp * 0.55
            x = lx + sp / 2
            Do While x < lx + cw
                y = by + sp / 2
                Do While y < by + ch
                    Set s = ActiveLayer.CreateEllipse(x - 0.25, y - 0.25, x + 0.25, y + 0.25)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 25
                    s.Outline.Type = cdrNoOutline
                    cnt = cnt + 1
                    y = y + sp
                Loop
                x = x + sp
            Loop
            
        ' --- 粉砂岩：稀疏微点 + 横线 ---
        Case "silt"
            ' 横线
            y = by + sp * 0.5
            Do While y < by + ch
                Set s = ActiveLayer.CreateLineSegment(lx + 1, y, lx + cw - 1, y)
                s.Outline.Color.CMYKAssign 0, 0, 0, 20
                s.Outline.Width = 0.2
                cnt = cnt + 1
                y = y + sp
            Loop
            ' 稀疏点
            x = lx + sp
            Do While x < lx + cw
                y = by + sp
                Do While y < by + ch
                    Set s = ActiveLayer.CreateEllipse(x - 0.2, y - 0.2, x + 0.2, y + 0.2)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 18
                    s.Outline.Type = cdrNoOutline
                    cnt = cnt + 1
                    y = y + sp * 2
                Loop
                x = x + sp * 2
            Loop
            
        ' --- 泥岩：密横线 ---
        Case "mud"
            sp = sp * 0.6
            y = by + sp * 0.3
            Do While y < by + ch
                Set s = ActiveLayer.CreateLineSegment(lx + 0.5, y, lx + cw - 0.5, y)
                s.Outline.Color.CMYKAssign 0, 0, 0, 30
                s.Outline.Width = 0.22
                cnt = cnt + 1
                y = y + sp
            Loop
            
        ' --- 页岩：横线 + 短竖线 ---
        Case "shale"
            sp = sp * 0.7
            y = by + sp * 0.3
            Do While y < by + ch
                ' 主横线
                Set s = ActiveLayer.CreateLineSegment(lx + 0.3, y, lx + cw - 0.3, y)
                s.Outline.Color.CMYKAssign 0, 0, 0, 30
                s.Outline.Width = 0.22
                cnt = cnt + 1
                ' 短竖线点缀（隔一根加一根）
                If cnt Mod 3 <> 0 Then
                    Dim svx As Double
                    svx = lx + sp * 0.8
                    Do While svx < lx + cw
                        Dim svLen As Double
                        svLen = sp * 0.3
                        Set s = ActiveLayer.CreateLineSegment(svx, y - svLen, svx, y + svLen)
                        s.Outline.Color.CMYKAssign 0, 0, 0, 20
                        s.Outline.Width = 0.15
                        cnt = cnt + 1
                        svx = svx + sp * 1.5
                    Loop
                End If
                y = y + sp
            Loop
            
        ' --- 炭质页岩：底色黑 + 横虚线 ---
        Case "carbShale"
            ' 底色直接用 CMYK 填充，花纹加白虚线
            sp = sp * 0.7
            y = by + sp * 0.3
            Do While y < by + ch
                Set s = ActiveLayer.CreateLineSegment(lx + 0.3, y, lx + cw - 0.3, y)
                s.Outline.Color.CMYKAssign 0, 0, 0, 0   ' 白色线
                s.Outline.Width = 0.18
                cnt = cnt + 1
                y = y + sp
            Loop
            
        ' --- 石灰岩：砖格 ---
        Case "lime"
            Dim bw As Double, bh As Double, lrow As Long
            bw = sp * 1.5: bh = sp * 0.6
            lrow = 0
            y = by + 0.3
            Do While y < by + ch
                x = lx + 0.3
                If lrow Mod 2 = 1 Then x = x + bw / 2
                Do While x < lx + cw
                    Set s = ActiveLayer.CreateRectangle(x, y, x + bw - 0.4, y + bh)
                    s.Fill.Type = cdrNoFill
                    s.Outline.Color.CMYKAssign 0, 0, 0, 22
                    s.Outline.Width = 0.16
                    cnt = cnt + 1
                    x = x + bw
                Loop
                y = y + bh + 0.25
                lrow = lrow + 1
            Loop
            
        ' --- 白云岩：菱格/十字 ---
        Case "dolo"
            ' 45° 交叉细线（菱格效果）
            Dim dx1 As Double
            dx1 = lx - (cw + ch)
            Do While dx1 < lx + cw + ch
                Set s = ActiveLayer.CreateLineSegment(dx1, by, dx1 + ch, by + ch)
                s.Outline.Color.CMYKAssign 0, 0, 0, 18
                s.Outline.Width = 0.15
                cnt = cnt + 1
                dx1 = dx1 + sp * 1.2
            Loop
            dx1 = lx - (cw + ch)
            Do While dx1 < lx + cw + ch
                Set s = ActiveLayer.CreateLineSegment(dx1, by + ch, dx1 + ch, by)
                s.Outline.Color.CMYKAssign 0, 0, 0, 18
                s.Outline.Width = 0.15
                cnt = cnt + 1
                dx1 = dx1 + sp * 1.2
            Loop
            
        ' --- 白云质灰岩：砖格+菱格混合 ---
        Case "doloLime"
            ' 先画石灰岩砖格（稀）
            bw = sp * 2: bh = sp * 0.7
            lrow = 0
            y = by + 0.3
            Do While y < by + ch
                x = lx + 0.3
                If lrow Mod 2 = 1 Then x = x + bw / 2
                Do While x < lx + cw
                    Set s = ActiveLayer.CreateRectangle(x, y, x + bw - 0.4, y + bh)
                    s.Fill.Type = cdrNoFill
                    s.Outline.Color.CMYKAssign 0, 0, 0, 15
                    s.Outline.Width = 0.12
                    cnt = cnt + 1
                    x = x + bw
                Loop
                y = y + bh + 0.25
                lrow = lrow + 1
            Loop
            ' 再加稀疏菱格
            dx1 = lx - (cw + ch)
            Do While dx1 < lx + cw + ch
                Set s = ActiveLayer.CreateLineSegment(dx1, by, dx1 + ch, by + ch)
                s.Outline.Color.CMYKAssign 0, 0, 0, 12
                s.Outline.Width = 0.1
                cnt = cnt + 1
                dx1 = dx1 + sp * 2
            Loop
            
        ' --- 硅质岩：交叉粗线 ---
        Case "chert"
            dx1 = lx - (cw + ch)
            Do While dx1 < lx + cw + ch
                Set s = ActiveLayer.CreateLineSegment(dx1, by, dx1 + ch, by + ch)
                s.Outline.Color.CMYKAssign 0, 0, 0, 28
                s.Outline.Width = 0.25
                cnt = cnt + 1
                dx1 = dx1 + sp
            Loop
            dx1 = lx - (cw + ch)
            Do While dx1 < lx + cw + ch
                Set s = ActiveLayer.CreateLineSegment(dx1, by + ch, dx1 + ch, by)
                s.Outline.Color.CMYKAssign 0, 0, 0, 28
                s.Outline.Width = 0.25
                cnt = cnt + 1
                dx1 = dx1 + sp
            Loop
            
        ' --- 煤：全黑 ---
        Case "coal"
            ' 全黑填充即可，加两条白线表示层理
            y = by + ch / 3
            Set s = ActiveLayer.CreateLineSegment(lx + 1, y, lx + cw - 1, y)
            s.Outline.Color.CMYKAssign 0, 0, 0, 0
            s.Outline.Width = 0.3
            cnt = cnt + 1
            y = by + ch * 2 / 3
            Set s = ActiveLayer.CreateLineSegment(lx + 1, y, lx + cw - 1, y)
            s.Outline.Color.CMYKAssign 0, 0, 0, 0
            s.Outline.Width = 0.3
            cnt = cnt + 1
            
        ' ==================== 岩浆岩 ====================
        
        ' --- 花岗岩：叉 + 点 ---
        Case "granite"
            ' 随机叉号
            Randomize 123
            Dim gi As Long, gn As Long
            gn = Int(cw * ch / (sp * sp) * 0.8)
            For gi = 1 To gn
                x = lx + Rnd() * cw
                y = by + Rnd() * ch
                Dim gs As Double: gs = sp * 0.35
                Set s = ActiveLayer.CreateLineSegment(x - gs, y - gs, x + gs, y + gs)
                s.Outline.Color.CMYKAssign 0, 0, 0, 35
                s.Outline.Width = 0.18
                cnt = cnt + 1
                Set s = ActiveLayer.CreateLineSegment(x - gs, y + gs, x + gs, y - gs)
                s.Outline.Color.CMYKAssign 0, 0, 0, 35
                s.Outline.Width = 0.18
                cnt = cnt + 1
            Next gi
            ' 点
            x = lx + sp * 0.5
            Do While x < lx + cw
                y = by + sp * 0.5
                Do While y < by + ch
                    Set s = ActiveLayer.CreateEllipse(x - 0.3, y - 0.3, x + 0.3, y + 0.3)
                    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 30
                    s.Outline.Type = cdrNoOutline
                    cnt = cnt + 1
                    y = y + sp
                Loop
                x = x + sp
            Loop
            
        ' --- 玄武岩：斜向 V 字形 ---
        Case "basalt"
            sp = sp * 0.9
            y = by + sp * 0.5
            Do While y < by + ch
                x = lx + sp * 0.3
                Do While x < lx + cw
                    Dim vh As Double: vh = sp * 0.4
                    Set s = ActiveLayer.CreateLineSegment(x, y - vh, x + vh, y)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 30
                    s.Outline.Width = 0.2
                    cnt = cnt + 1
                    Set s = ActiveLayer.CreateLineSegment(x + vh, y, x + vh * 2, y - vh)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 30
                    s.Outline.Width = 0.2
                    cnt = cnt + 1
                    x = x + sp * 1.5
                Loop
                y = y + sp
            Loop
            
        ' ==================== 变质岩 ====================
        
        ' --- 片岩：波浪斜线 ---
        Case "schist"
            dx1 = lx - cw
            Do While dx1 < lx + cw * 2
                ' 用三点折线模拟波浪
                Dim mx As Double, my As Double
                mx = dx1 + ch * 0.5
                my = by + ch * 0.5
                Set s = ActiveLayer.CreateLineSegment(dx1, by, mx, my)
                s.Outline.Color.CMYKAssign 0, 0, 0, 28
                s.Outline.Width = 0.22
                cnt = cnt + 1
                Set s = ActiveLayer.CreateLineSegment(mx, my, dx1 + ch, by + ch)
                s.Outline.Color.CMYKAssign 0, 0, 0, 28
                s.Outline.Width = 0.22
                cnt = cnt + 1
                dx1 = dx1 + sp * 0.9
            Loop
            
        ' --- 片麻岩：条带状 ---
        Case "gneiss"
            sp = sp * 0.8
            y = by + sp * 0.2
            Dim bandOn As Boolean: bandOn = True
            Do While y < by + ch
                If bandOn Then
                    ' 粗条带
                    Set s = ActiveLayer.CreateLineSegment(lx + 0.3, y, lx + cw - 0.3, y)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 40
                    s.Outline.Width = 0.35
                Else
                    ' 细条带
                    Set s = ActiveLayer.CreateLineSegment(lx + 0.3, y, lx + cw - 0.3, y)
                    s.Outline.Color.CMYKAssign 0, 0, 0, 15
                    s.Outline.Width = 0.12
                End If
                cnt = cnt + 1
                bandOn = Not bandOn
                y = y + sp * IIf(bandOn, 0.6, 0.3)
            Loop
            
        ' --- 大理岩：网格+稀疏菱格 ---
        Case "marble"
            ' 细网格
            x = lx + sp * 0.5
            Do While x < lx + cw
                DrawVLine x, by, by + ch, 0.08, 12
                cnt = cnt + 1
                x = x + sp
            Loop
            y = by + sp * 0.5
            Do While y < by + ch
                DrawHLine lx, lx + cw, y, 0.08, 12
                cnt = cnt + 1
                y = y + sp
            Loop
            
    End Select
    
    ' --- PowerClip ---
    If cnt > 0 Then
        Dim grp As Shape
        Set grp = ActiveSelection.Group
        container.PowerClip.PlaceInside grp
    End If
    
    ' 恢复原始选中
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
Public Sub DrawColumn()
    
    On Error GoTo ErrHandler
    
    If ActiveDocument Is Nothing Then
        MsgBox "请先打开一个 CorelDRAW 文档", vbExclamation, "地层柱状图"
        Exit Sub
    End If
    
    ActiveDocument.BeginCommandGroup "绘制地层柱状图"
    
    ' ========== 初始化 ==========
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    Dim section() As StrataRow
    section = GetSection()
    Dim nLayers As Long
    nLayers = UBound(section) - LBound(section) + 1
    
    ' 总厚度
    Dim totalThick As Double
    totalThick = 0
    Dim i As Long
    For i = LBound(section) To UBound(section)
        totalThick = totalThick + section(i).Thickness
    Next i
    
    ' 可用柱体高度
    Dim drawH As Double
    drawH = TABLE_TOP - TABLE_BOTTOM - HEADER_H
    
    ' 比例系数
    Dim scaleF As Double
    scaleF = drawH / totalThick  ' mm per m
    
    ' 表格右边界
    Dim tableRight As Double
    tableRight = ColX(8) + ColW(8)
    
    ' ========== 标题 ==========
    Dim title As Shape
    Set title = ActiveLayer.CreateArtisticText(ColX(4), TABLE_TOP + 3, "秭归地区综合地层柱状图")
    title.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    With title.Text.Story
        .Font = "黑体"
        .Size = 11
        .Bold = True
    End With
    title.SetPosition (MARGIN_LEFT + tableRight) / 2, TABLE_TOP + 3
    
    ' ========== 表头 ==========
    Dim headerT As Double, headerB As Double
    headerT = TABLE_TOP
    headerB = TABLE_TOP - HEADER_H
    
    ' 表头底框
    DrawHLine MARGIN_LEFT, tableRight, headerB, 0.4, 100
    DrawHLine MARGIN_LEFT, tableRight, headerT, 0.3, 80
    
    ' 表头文字
    AddCellText "界", ColX(1), headerT, headerB, ColW(1), 7, True
    AddCellText "系", ColX(2), headerT, headerB, ColW(2), 7, True
    AddCellText "统", ColX(3), headerT, headerB, ColW(3), 7, True
    AddCellText "组", ColX(4), headerT, headerB, ColW(4), 7, True
    AddCellText "代号", ColX(5), headerT, headerB, ColW(5), 7, True
    AddCellText "柱 状 图", ColX(6), headerT, headerB, ColW(6), 7, True
    AddCellText "厚度(m)", ColX(7), headerT, headerB, ColW(7), 6.5, True
    AddCellText "岩 性 描 述", ColX(8), headerT, headerB, ColW(8), 7, True
    
    ' 表头竖向分隔线
    Dim ci As Long
    For ci = 2 To 8
        DrawVLine ColX(ci), TABLE_BOTTOM, TABLE_TOP, COL_LINE_W, 80
    Next ci
    
    ' ========== 深度标尺（在柱状图列右侧）==========
    Dim tickX As Double, tickLeftX As Double
    tickX = ColX(6)
    tickLeftX = ColX(6) + ColW(6) - 2
    
    ' 每 100m 一标
    Dim tickInterval As Double
    tickInterval = 100
    If totalThick <= 50 Then tickInterval = 10
    If totalThick > 50 And totalThick <= 200 Then tickInterval = 50
    If totalThick > 200 Then tickInterval = 100
    
    Dim tickM As Double
    tickM = 0
    Do While tickM <= totalThick
        Dim tickY As Double
        tickY = TABLE_TOP - HEADER_H - tickM * scaleF
        If tickY >= TABLE_BOTTOM Then
            DrawHLine tickX, tickLeftX, tickY, 0.2, 60
            ' 标注
            If tickM Mod (tickInterval * 2) = 0 Then
                Dim tkl As Shape
                Set tkl = ActiveLayer.CreateArtisticText(tickLeftX - 1.5, tickY - 1.5, Format(tickM, "0"))
                With tkl.Text.Story
                    .Font = "Arial"
                    .Size = 5.5
                End With
                tkl.Fill.UniformColor.CMYKAssign 0, 0, 0, 60
                tkl.Outline.Type = cdrNoOutline
            End If
        End If
        tickM = tickM + tickInterval
    Loop
    
    ' 深度标尺垂直线
    DrawVLine tickX, TABLE_BOTTOM, headerB, 0.1, 40
    
    ' ========== 逐层绘制 ==========
    Dim currentY As Double
    currentY = headerB   ' 从表头下方开始往下画
    
    For i = LBound(section) To UBound(section)
        Dim lay As StrataRow
        lay = section(i)
        
        Dim layH As Double
        layH = lay.Thickness * scaleF
        If layH < 3.5 Then layH = 3.5   ' 最小显示高度
        
        Dim layBottom As Double
        layBottom = currentY - layH
        
        ' --- 1) 柱状图矩形（col 6）---
        Dim colLeft As Double:  colLeft = ColX(6) + 0.3
        Dim colRight As Double: colRight = ColX(6) + ColW(6) - 2.5
        Dim colW_actual As Double: colW_actual = colRight - colLeft
        
        Dim rect As Shape
        Set rect = ActiveLayer.CreateRectangle(colLeft, layBottom, colRight, currentY)
        rect.Fill.UniformColor.CMYKAssign lay.C, lay.M, lay.Y_, lay.K
        rect.Outline.Color.CMYKAssign 0, 0, 0, 35
        rect.Outline.Width = 0.15
        rect.Outline.ScaleWithShape = False
        
        ' 花纹
        AddPattern rect, lay.Pattern, colLeft, layBottom, colW_actual, layH
        
        ' --- 2) 文本列 ---
        ' 界 (col 1)
        AddCellText lay.Erathem, ColX(1), currentY, layBottom, ColW(1), 5.5
        ' 系 (col 2)
        AddCellText lay.System, ColX(2), currentY, layBottom, ColW(2), 5.5
        ' 统 (col 3)
        AddCellText lay.Series, ColX(3), currentY, layBottom, ColW(3), 5.5
        ' 组 (col 4)
        AddCellText lay.Formation, ColX(4), currentY, layBottom, ColW(4), 5.5
        ' 代号 (col 5)
        AddCellText lay.Symbol, ColX(5), currentY, layBottom, ColW(5), 5, True
        ' 厚度 (col 7)
        AddCellText Format(lay.Thickness, "0.0"), ColX(7), currentY, layBottom, ColW(7), 5.5
        ' 岩性描述 (col 8)
        AddCellText lay.Descr, ColX(8), currentY, layBottom, ColW(8), 5
        
        ' --- 3) 层间分隔线 ---
        DrawHLine MARGIN_LEFT, tableRight, currentY, COL_LINE_W, 60
        
        currentY = layBottom
    Next i
    
    ' ========== 底部封闭线 ==========
    DrawHLine MARGIN_LEFT, tableRight, TABLE_BOTTOM, 0.4, 100
    
    ' ========== 表格外边框加粗 ==========
    DrawHLine MARGIN_LEFT, tableRight, TABLE_TOP, 0.5, 100
    DrawVLine MARGIN_LEFT, TABLE_BOTTOM, TABLE_TOP, 0.4, 100
    DrawVLine tableRight, TABLE_BOTTOM, TABLE_TOP, 0.4, 100
    
    ' ========== 底部图签 ==========
    Dim footer As Shape
    Set footer = ActiveLayer.CreateArtisticText(MARGIN_LEFT, TABLE_BOTTOM - 6, _
        "秭归地区综合地层柱状图  |  总厚 " & Format(totalThick, "#,##0") & "m  |  " & _
        nLayers & " 层  |  比例尺 1:" & Format(totalThick * 1000 / drawH, "#,##0") & _
        "  |  依据：秭归实习基地标准图例")
    With footer.Text.Story
        .Font = "黑体"
        .Size = 5.5
        .Italic = True
    End With
    footer.Fill.UniformColor.CMYKAssign 0, 0, 0, 40
    footer.Outline.Type = cdrNoOutline
    
    ' ========== 图例（右下） ==========
    ' 提取去重花纹
    Dim pTypes() As String, pNames() As String
    Dim pCMYs() As Long
    Dim pCount As Long
    pCount = 0
    
    For i = LBound(section) To UBound(section)
        Dim found As Boolean
        found = False
        Dim j As Long
        For j = 0 To pCount - 1
            If pTypes(j) = section(i).Pattern Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            ReDim Preserve pTypes(0 To pCount)
            ReDim Preserve pNames(0 To pCount)
            ReDim Preserve pCMYs(0 To pCount)
            pTypes(pCount) = section(i).Pattern
            pNames(pCount) = section(i).Formation & " — " & section(i).Descr
            pCMYs(pCount) = section(i).C * 1000000 + section(i).M * 10000 + _
                            section(i).Y_ * 100 + section(i).K
            pCount = pCount + 1
        End If
    Next i
    
    ' 只在超过 1 种花纹时画图例
    If pCount > 1 Then
        Dim lgX As Double, lgY As Double
        lgX = tableRight + 4
        lgY = TABLE_TOP - HEADER_H
        
        ' 图例框
        Dim lgBox As Shape
        Set lgBox = ActiveLayer.CreateRectangle(lgX, lgY - pCount * 8 - 6, lgX + 40, lgY + 2)
        lgBox.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
        lgBox.Outline.Color.CMYKAssign 0, 0, 0, 50
        lgBox.Outline.Width = 0.25
        
        ' 标题
        Dim lgTitle As Shape
        Set lgTitle = ActiveLayer.CreateArtisticText(lgX + 3, lgY + 1, "图例")
        With lgTitle.Text.Story
            .Font = "黑体"
            .Size = 6.5
            .Bold = True
        End With
        lgTitle.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        
        For j = 0 To pCount - 1
            Dim iy As Double
            iy = lgY - 6 - j * 8
            
            ' 色块
            Dim ib As Shape
            Set ib = ActiveLayer.CreateRectangle(lgX + 3, iy - 4, lgX + 10, iy + 3)
            Dim ic As Long, im As Long, iy_ As Long, ik As Long
            ic = pCMYs(j) \ 1000000
            im = (pCMYs(j) Mod 1000000) \ 10000
            iy_ = (pCMYs(j) Mod 10000) \ 100
            ik = pCMYs(j) Mod 100
            ib.Fill.UniformColor.CMYKAssign ic, im, iy_, ik
            ib.Outline.Color.CMYKAssign 0, 0, 0, 40
            ib.Outline.Width = 0.12
            
            ' 花纹缩略
            If pTypes(j) <> "pure" Then
                AddPattern ib, pTypes(j), lgX + 3, iy - 4, 7, 7
            End If
            
            ' 文字
            Dim il As Shape
            Dim patLabel As String
            patLabel = Left(pNames(j), 18)
            Set il = ActiveLayer.CreateArtisticText(lgX + 12, iy - 0.5, patLabel)
            With il.Text.Story
                .Font = "黑体"
                .Size = 5
            End With
            il.Fill.UniformColor.CMYKAssign 0, 0, 0, 80
            il.Outline.Type = cdrNoOutline
        Next j
    End If
    
    ' ========== 完成 ==========
    ActiveDocument.EndCommandGroup
    MsgBox "地层柱状图绘制完成！" & vbCrLf & vbCrLf & _
           nLayers & " 个地层单元 · 总厚度 " & Format(totalThick, "#,##0.0") & "m" & vbCrLf & _
           "比例尺约 1:" & Format(totalThick * 1000 / drawH, "#,##0"), _
           vbInformation, "地层柱状图（国家标准版）"
    Exit Sub
    
ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "绘图出错: " & Err.Description & " (错误号 " & Err.Number & ")", vbCritical, "地层柱状图 — 错误"
End Sub
