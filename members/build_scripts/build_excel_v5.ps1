$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()

# Colors
$darkBg = 0x1a1a2e
$gold = 0xd4a574
$white = 0xffffff
$lightGray = 0xf5f5f5
$headerBlue = 0x2c3e50
$accentGreen = 0x27ae60
$accentRed = 0xe74c3c
$accentOrange = 0xf39c12
$lightBlue = 0xd6eaf8
$softPurple = 0x8e44ad
$softBlue = 0x2980b9
$coverBg = 0x1a1a2e

function Set-Cell($ws, $row, $col, $value, $bold, $size, $color, $bg, $wrap) {
    $cell = $ws.Cells.Item($row, $col)
    $cell.Value = "$value"
    $cell.Font.Bold = $bold
    $cell.Font.Size = $size
    if ($color) { $cell.Font.Color = $color }
    if ($bg) { $cell.Interior.Color = $bg }
    $cell.WrapText = $wrap
}

function Merge-Cells($ws, $row, $colStart, $colEnd) {
    $ws.Range($ws.Cells.Item($row, $colStart), $ws.Cells.Item($row, $colEnd)).Merge()
}

function Set-Borders($ws, $row, $colStart, $colEnd) {
    $range = $ws.Range($ws.Cells.Item($row, $colStart), $ws.Cells.Item($row, $colEnd))
    $range.Borders.LineStyle = 1
    $range.Borders.Weight = 2
}

function Write-TrainingSection($ws, $refRow, $title, $subtitle, $estTime, $actions) {
    Set-Cell $ws $refRow 1 $title $true 13 $white $darkBg $true
    Merge-Cells $ws $refRow 1 9
    $ws.Range("A$($refRow):I$($refRow)").RowHeight = 28
    $ws.Range("A$($refRow):I$($refRow)").Font.Color = $white
    $ws.Range("A$($refRow):I$($refRow)").Interior.Color = $darkBg

    $refRow++
    Set-Cell $ws $refRow 1 $subtitle $false 10 $headerBlue $lightGray $true
    Merge-Cells $ws $refRow 1 9
    $refRow++

    Set-Cell $ws $refRow 1 "序" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 2 "动作" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 3 "组数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 4 "次数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 5 "重量/负载" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 6 "休息" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 7 "RPE" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 8 "降级" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 9 "动作要点（呼吸+纠偏+发力感）" $true 10 $white $headerBlue $true
    $ws.Range("A$($refRow):I$($refRow)").RowHeight = 28
    $refRow++

    $i = 1
    foreach ($a in $actions) {
        Set-Cell $ws $refRow 1 $i $false 9 0 $lightGray $false
        Set-Cell $ws $refRow 2 $a[0] $true 9 $headerBlue 0 $false
        Set-Cell $ws $refRow 3 $a[1] $false 9 0 0 $false
        Set-Cell $ws $refRow 4 $a[2] $false 9 0 0 $false
        Set-Cell $ws $refRow 5 $a[3] $false 8 0 0 $false
        Set-Cell $ws $refRow 6 $a[4] $false 9 0 0 $false
        Set-Cell $ws $refRow 7 $a[5] $false 9 0 0 $false
        Set-Cell $ws $refRow 8 $a[6] $false 8 0 0 $false
        Set-Cell $ws $refRow 9 $a[7] $false 8 0 0 $true
        $ws.Range("A$($refRow):I$($refRow)").RowHeight = 62
        Set-Borders $ws $refRow 1 9
        $refRow++
        $i++
    }
    return $refRow + 2
}

Write-Host "Setup complete"

# ============================================
# SHEET 0: 封面-今天练什么
# ============================================
$ws0 = $wb.Worksheets.Item(1)
$ws0.Name = " 今天练什么"

Set-Cell $ws0 1 1 "Joe 何  训练方案" $true 22 $gold $darkBg $true
Merge-Cells $ws0 1 1 5
$ws0.Range("A1:E1").RowHeight = 50
$ws0.Range("A1:E1").Font.Color = $gold
$ws0.Range("A1:E1").Interior.Color = $darkBg

Set-Cell $ws0 2 1 "崔知行教练  2026-06-18  |  V5" $false 11 $white $darkBg $true
Merge-Cells $ws0 2 1 5
$ws0.Range("A2:E2").Font.Color = $white
$ws0.Range("A2:E2").Interior.Color = $darkBg

$row = 4
Set-Cell $ws0 $row 1 "  快速开始——找到今天" $true 16 $white $accentGreen $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 34
$ws0.Range("A$($row):E$($row)").Font.Color = $white
$ws0.Range("A$($row):E$($row)").Interior.Color = $accentGreen
$row++

# Week 1-3 blocks
for ($w = 1; $w -le 3; $w++) {
    Set-Cell $ws0 $row 1 "第${w}周" $true 14 $white $headerBlue $true
    Merge-Cells $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight = 28
    $ws0.Range("A$($row):E$($row)").Font.Color = $white
    $ws0.Range("A$($row):E$($row)").Interior.Color = $headerBlue
    $row++

    Set-Cell $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
    Set-Cell $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
    Set-Cell $ws0 $row 3 "强度/感觉" $true 10 $white $headerBlue $false
    Set-Cell $ws0 $row 4 "去哪个Sheet" $true 10 $white $headerBlue $false
    Set-Cell $ws0 $row 5 "一句话提醒" $true 10 $white $headerBlue $false
    $row++

    $dayData = @(
        @("周一","强度日","RPE 8-9 / 重","Sheet 2  第1-3周","做完剩1-2次力竭——不留余力也别力竭"),
        @("周三","容量日","RPE 6-7 / 累","Sheet 2  第1-3周","组间短，泵感强——享受充血"),
        @("周五","爆发日","RPE 5-6 / 快","Sheet 2  第1-3周","轻快有力——如果觉得重就是太重了")
    )

    foreach ($d in $dayData) {
        Set-Cell $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
        Set-Cell $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
        Set-Cell $ws0 $row 3 $d[2] $false 10 0 0 $false
        Set-Cell $ws0 $row 4 $d[3] $false 9 0 0 $false
        Set-Cell $ws0 $row 5 $d[4] $false 9 0 0 $true
        $ws0.Range("A$($row):E$($row)").RowHeight = 26
        Set-Borders $ws0 $row 1 5
        $row++
    }
    $row++
}

# Week 4 Deload
Set-Cell $ws0 $row 1 "第4周  减载周" $true 14 $white $accentGreen $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 28
$ws0.Range("A$($row):E$($row)").Font.Color = $white
$ws0.Range("A$($row):E$($row)").Interior.Color = $accentGreen
$row++

Set-Cell $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "感觉" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 4 "去哪个Sheet" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 5 "一句话提醒" $true 10 $white $headerBlue $false
$row++

$deloadCover = @(
    @("周一/三/五","减载周","RPE 4-5 / 轻松","Sheet 3  第4周","换动作——做完应该感觉'没练够'"),
    @("3天都练","'动但不同'","不碰主力动作","","弹力带+自重为主——关节在休息")
)

foreach ($d in $deloadCover) {
    Set-Cell $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    Set-Cell $ws0 $row 3 $d[2] $false 10 0 0 $false
    Set-Cell $ws0 $row 4 $d[3] $false 9 0 0 $false
    Set-Cell $ws0 $row 5 $d[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight = 26
    Set-Borders $ws0 $row 1 5
    $row++
}

# RPE quick reference
$row += 2
Set-Cell $ws0 $row 1 "  RPE 快速理解" $true 14 $white $softPurple $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 28
$ws0.Range("A$($row):E$($row)").Font.Color = $white
$ws0.Range("A$($row):E$($row)").Interior.Color = $softPurple
$row++

Set-Cell $ws0 $row 1 "RPE就是你做完一组后凭感觉打分。1分=躺着不动，10分=拼了老命。不用纠结数字精确到0.5，感觉对了就行。" $false 10 $lightGray 0 $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 22
$row++

Set-Cell $ws0 $row 1 "RPE" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "感觉" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "还剩几次力竭" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 4 "你的训练日" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 5 "一句话" $true 10 $white $headerBlue $false
$row++

$rpeCover = @(
    @("8-9","很重，接近极限","1-2次","周一  强度日","做完感觉'再来一个就起不来了'——但你没起不来"),
    @("6-7","中等偏强","3-4次","周三  容量日","肌肉充血发胀，就像充了气——但还能再挤几个"),
    @("5-6","轻松→中等","6-8次","周五  爆发日","轻快有力，像弹簧——如果觉得重，就是太重了"),
    @("4-5","非常轻松","8-10次+","减载周","做完怀疑'我今天练了什么？'——这就对了")
)

foreach ($rpe in $rpeCover) {
    Set-Cell $ws0 $row 1 $rpe[0] $true 12 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $rpe[1] $false 10 0 0 $false
    Set-Cell $ws0 $row 3 $rpe[2] $false 10 0 0 $false
    Set-Cell $ws0 $row 4 $rpe[3] $false 10 0 0 $false
    Set-Cell $ws0 $row 5 $rpe[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight = 30
    Set-Borders $ws0 $row 1 5
    $row++
}

# Emergency rule
$row += 2
Set-Cell $ws0 $row 1 "  今天状态极差？降级不丢人" $true 14 $white $accentOrange $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 28
$ws0.Range("A$($row):E$($row)").Font.Color = $white
$ws0.Range("A$($row):E$($row)").Interior.Color = $accentOrange
$row++

Set-Cell $ws0 $row 1 "  起床就觉得累、没睡好、压力大？不要硬撑，按下面降级：" $false 10 0 0 $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 22
$row++

Set-Cell $ws0 $row 1 "今天的主题" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "降级为" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "RPE降为" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 4 "底线" $true 10 $white $headerBlue $false
$row++

$emergency = @(
    @("强度日(周一)","容量日","8-9→6-7","至少做热身后3个动作"),
    @("容量日(周三)","爆发日","6-7→5-6","至少热身+有氧10min"),
    @("爆发日(周五)","休息或只热身+有氧","5-6→休息","做了热身就是胜利"),
    @("减载周","更轻——只做弹力带动作","4-5→3-4","多滚泡沫轴多拉伸")
)

foreach ($em in $emergency) {
    Set-Cell $ws0 $row 1 $em[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $em[1] $true 10 $white $darkBg $false
    Set-Cell $ws0 $row 3 $em[2] $false 10 0 0 $false
    Set-Cell $ws0 $row 4 $em[3] $false 10 0 0 $true
    Merge-Cells $ws0 $row 4 5
    $ws0.Range("A$($row):E$($row)").RowHeight = 28
    Set-Borders $ws0 $row 1 5
    $row++
}

$row++
Set-Cell $ws0 $row 1 "  一次降级不会影响长期进步。一次硬撑却可能受伤停练2周。算不过来账吗？" $false 10 $accentOrange 0 $true
Merge-Cells $ws0 $row 1 5

# Other sheet index
$row += 2
Set-Cell $ws0 $row 1 "  其他内容速查" $true 14 $white $darkBg $true
Merge-Cells $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight = 28
$ws0.Range("A$($row):E$($row)").Font.Color = $white
$ws0.Range("A$($row):E$($row)").Interior.Color = $darkBg
$row++

Set-Cell $ws0 $row 1 "想查什么" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "去这个Sheet" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "内容" $true 10 $white $headerBlue $false
Merge-Cells $ws0 $row 3 5
$row++

$quickIndex = @(
    @("今天练什么（周一/三/五）","Sheet 2","每天完整动作表+重量+休息+RPE"),
    @("减载周怎么练（第4周）","Sheet 3","10对动作替换表+规则"),
    @("这周怎么调整（渐进）","Sheet 4","4周渐进节奏+器械上限策略"),
    @("热身/有氧/练前练后吃什么","Sheet 5","8步热身+有氧参数+训练前后餐"),
    @("膝盖保护/营养/进阶规则","Sheet 6","RPE指南+膝盖分级+全休周规则"),
    @("应酬周只能练2次/复盘","Sheet 7","精简备用方案+4周后复盘清单"),
    @("整体方案逻辑","Sheet 方案概览","DUP周期解释+全年框架")
)

foreach ($qi in $quickIndex) {
    Set-Cell $ws0 $row 1 $qi[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $qi[1] $false 10 0 0 $false
    Set-Cell $ws0 $row 3 $qi[2] $false 10 0 0 $true
    Merge-Cells $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight = 24
    Set-Borders $ws0 $row 1 5
    $row++
}

$ws0.Range("A:A").ColumnWidth = 22
$ws0.Range("B:B").ColumnWidth = 20
$ws0.Range("C:E").ColumnWidth = 22

Write-Host "Cover sheet done"

# ============================================
# SHEET 1: 方案概览
# ============================================
$ws = $wb.Worksheets.Add()
$ws.Name = "方案概览"

Set-Cell $ws 1 1 "知行AI教练  定制训练方案 V5" $true 18 $headerBlue 0 $true
Merge-Cells $ws 1 1 9
$ws.Range("A1:I1").RowHeight = 38

Set-Cell $ws 2 1 "Joe  He  |  37岁  |  168cm / 80kg  |  减脂增肌  身体重组  |  DUP波动周期" $false 12 $gold $darkBg $true
Merge-Cells $ws 2 1 9
$ws.Range("A2:I2").RowHeight = 28
$ws.Range("A2:I2").Font.Color = $white
$ws.Range("A2:I2").Interior.Color = $darkBg

$row = 4
Set-Cell $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $row 1 9
$ws.Range("A$($row):I$($row)").RowHeight = 26

$overview = @(
    @("训练频率","每周3练 + 2练精简备用"),
    @("每次时长","55-65分钟：热身8-10 + 力量35-45 + 有氧10-14 + 拉伸5"),
    @("训练方式","全身训练x3/周，DUP波动周期——每天不同刺激主题"),
    @("周一=强度日(80-85% RPE8-9)","周三=容量日(65-70% RPE6-7)","周五=爆发日(50-60% RPE5-6)"),
    @("周期结构","3周DUP训练 + 1周减载(换动作) = 1小周期","每3小周期 + 1周全休 = 1中周期(季度)"),
    @("加重原则","上肢+1-2.5kg","下肢+2.5-5kg"),
    @("方案特点","DUP波动 RPE智能调节 神经激活 减载换动作 训练窗口营养 状态降级兜底")
)

$r = 5
foreach ($item in $overview) {
    Set-Cell $ws $r 1 $item[0] $true 10 $headerBlue $lightGray $false
    if ($item.Count -eq 2) {
        Set-Cell $ws $r 3 $item[1] $false 10 0 0 $true
        Merge-Cells $ws $r 3 9
    } elseif ($item.Count -eq 3) {
        Set-Cell $ws $r 3 $item[1] $false 10 0 0 $true
        Merge-Cells $ws $r 3 7
        Set-Cell $ws $r 8 $item[2] $false 10 0 0 $true
        Merge-Cells $ws $r 8 9
    }
    Set-Borders $ws $r 1 9
    $r++
}

# DUP logic
$r++
Set-Cell $ws $r 1 "DUP波动周期——为什么每天不同主题？" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 9
$r++

$dupWhy = @(
    "你有1.5年训练基础——已过新手期。线性周期（全周同一主题）容易适应性停滞。",
    "DUP每周轮换3种刺激——神经系统始终面临'新'挑战，适应更慢，进步更持续。",
    "周一重(强度)、周三累(容量)、周五快(爆发)——每次训练期待感不同，坚持更容易。",
    "减载周结束后，身体对3种刺激都保持敏感——恢复后能同时突破多个维度。"
)
foreach ($dw in $dupWhy) {
    Set-Cell $ws $r 1 $dw $false 10 0 0 $true
    Merge-Cells $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight = 22
    $r++
}

# Year cycle
$r++
Set-Cell $ws $r 1 "全年周期框架" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 9
$r++

Set-Cell $ws $r 1 "大周期(1年)" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "身体重组80kg→70-72kg + 力量稳步提升" $false 10 0 0 $true
Merge-Cells $ws $r 2 9
Set-Borders $ws $r 1 9
$r++
Set-Cell $ws $r 1 "中周期(1季度)" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "3个小周期(每月) + 1周全休——每3个月彻底重置系统。全休周只做散步/游泳/拉伸/泡沫轴，不碰器械。" $false 10 0 0 $true
Merge-Cells $ws $r 2 9
Set-Borders $ws $r 1 9
$r++
Set-Cell $ws $r 1 "小周期(1个月)" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "3周DUP训练 + 1周减载(换动作)——你手里这份方案的结构" $false 10 0 0 $true
Merge-Cells $ws $r 2 9
Set-Borders $ws $r 1 9
$r++
Set-Cell $ws $r 1 "全休周逻辑" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "即使每月有减载周，3个月连续训练仍会累积神经系统+内分泌+关节的深层疲劳。全休1周不是偷懒——是给身体'系统重置'的机会。休息完回来，力量不会掉，反而会涨——因为超量恢复真正完成了。" $false 10 0 0 $true
Merge-Cells $ws $r 2 9
$ws.Range("A$($r):I$($r)").RowHeight = 34
Set-Borders $ws $r 1 9
$r++

$r++
Set-Cell $ws $r 1 "三大核心原则" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 9
$r++

$principles = @(
    @("1","无痛优先","膝关节不适期间，所有下肢动作以无痛为第一原则。宁可退阶不冒险。"),
    @("2","RPE优先于数字","80% 1RM是理论值——当天睡眠/压力/饮食会改变实际表现。RPE让你和身体对话，而非死磕数字。"),
    @("3","动但不同","减载周不是'做轻一点'——是'做不同的动作'。换动作=关节不同应力路径，让主力动作的关节真正休息。")
)

foreach ($p in $principles) {
    Set-Cell $ws $r 1 $p[0] $false 10 0 0 $false
    Set-Cell $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws $r 3 $p[2] $false 10 0 0 $true
    Merge-Cells $ws $r 3 9
    $ws.Range("A$($r):I$($r)").RowHeight = 30
    Set-Borders $ws $r 1 9
    $r++
}

$ws.Range("A:A").ColumnWidth = 28
$ws.Range("B:I").ColumnWidth = 16

Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 第1-3周 DUP训练计划（带降级列）
# ============================================
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "第1-3周训练"

Set-Cell $ws2 1 1 "第1-3周  DUP波动周期训练（周一/三/五）" $true 16 $headerBlue 0 $true
Merge-Cells $ws2 1 1 9
$ws2.Range("A1:I1").RowHeight = 34
Set-Cell $ws2 2 1 "每周3天不同主题。找到今天的格子，对号入座。'降级'列=今天状态差时做什么。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws2 2 1 9

$r = 4

# === 周一：强度日 ===
$actStrength = @(
    @("哑铃平板卧推","4组","4-6次","单边18-20kg(80-85%)","120秒","8-9","降为16kg做容量","【呼吸】下放吸气屏住(瓦式)，推起呼气。【纠偏】手肘与躯干45-60，不外展90。【发力感】推起时想象把哑铃'推弯'。"),
    @("器械坐姿推肩","3组","4-6次","插片选4-6RM","90秒","8-9","降1-2片做8-10次","【呼吸】推起呼气、下放吸气。【纠偏】下背贴靠背不反弓。【发力感】推至顶端不锁死。上肢精细+1-2.5kg。"),
    @("哑铃分腿蹲","4组","4-6次/侧","单边16-18kg(强度日)","120秒","8-9","降10kg做12次/ 膝不适→跳过","【呼吸】下蹲吸气、起身呼气。【纠偏】躯干竖直，膝盖无痛范围。【发力感】前腿发力。 无痛才加重，下肢可+2.5-5kg。"),
    @("哑铃臀推","3组","6-8次","单颗18-20kg(冲上限)","90秒","8-9","降14kg做15次臀桥","【呼吸】推起呼气、下放吸气。【纠偏】下巴微收不超伸腰椎。【发力感】顶峰夹紧2秒。达20kg→单腿臀推。"),
    @("绳索飞鸟+下压超级组","各3组","各10-12次","比容量日重1档","60秒","6-7","降轻重量做15次","强度日的辅助按容量做。飞鸟：胸肌驱动。下压：大臂固定不外摆。")
)
$r = Write-TrainingSection $ws2 $r "周一  强度日（Strength）" "80-85%极限  |  主项4-6次  |  RPE 8-9  |  做完剩1-2次力气" "42-48" $actStrength

# === 周三：容量日 ===
$actVolume = @(
    @("哑铃平板卧推","4组","10-12次","单边14-16kg(65-70%)","60秒","6-7","降12kg做15次爆发推","【呼吸】下放吸气、推起呼气。【纠偏】轻了动作质量更要高。【发力感】追求胸肌泵感，收缩1秒。容量日组间短代谢压力大。"),
    @("高位下拉","4组","10-12次","插片选12-15RM","60秒","6-7","降轻2片做15次慢离心","【呼吸】下拉呼气、还原吸气。【纠偏】先沉肩胛再拉——不借力。【发力感】背阔肌全程控制，顶峰收缩1秒。"),
    @("哑铃罗马尼亚硬拉","4组","10-12次","单边14-18kg","60秒","6-7","降12kg做15次/膝不适→跳过","【呼吸】下放吸气、拉起呼气。【纠偏】膝盖微屈，哑铃贴腿下滑。容量日高次数追求腘绳肌泵感。"),
    @("龙门架面拉","4组","15-20次","轻重量","45秒","6-7","降更轻做25次","【呼吸】拉向面部呼气、还原吸气。【纠偏】拉向额头。高次数面拉=肩后束代谢压力积累。"),
    @("哑铃弯举+仰卧举腿超级组","各3组","弯举15次/举腿力竭","弯举单边6-10kg","45秒","7-8","降轻弯举做20次","弯举：杜绝摆动，离心3秒。举腿：下腹带动骨盆上卷。")
)
$r = Write-TrainingSection $ws2 $r "周三  容量日（Volume）" "65-70%极限  |  组间短(45-60s)  |  高次数泵感  |  RPE 6-7  |  充血像气球" "40-45" $actVolume

# === 周五：爆发日 ===
$actPower = @(
    @("爆发哑铃卧推","6组","3次","单边10-12kg(50-60%)","45秒","5-6","降8kg或换俯卧撑爆发","【呼吸】下放控制吸气、爆发推起呼气。【纠偏】轻重量——重了就没速度。【发力感】想象把哑铃'弹'出去。质量>数量。"),
    @("爆发高位下拉+慢放","5组","5次","插片选50-60%容量日","60秒","5-6","降轻2片做8次","【纠偏】爆发下拉+慢离心3秒——拉快放慢。【发力感】下拉追求启动速度，还原感受背阔肌拉伸。"),
    @("分腿蹲爆发起","4组","5次/侧","自重或持轻哑铃4-6kg","60秒","5-6","只做自重控制节奏","【纠偏】下蹲后爆发起身——不是跳，是快速伸展髋膝。【发力感】前腿爆发力输出。 膝关节无冲击。"),
    @("爆发臀推+慢放","4组","5次","单颗10-14kg(60%强度日)","60秒","5-6","降自重+弹力带臀桥","【纠偏】爆发推起+慢离心3秒。【发力感】臀肌瞬间爆发收紧。"),
    @("弹力带面拉+弹力带推胸","各3组","各8-10次","弹力带中等阻力","45秒","5-6","降更轻弹力带做12次","爆发日弹力带收尾。变阻特性适合速度日——顶端阻力更大。")
)
$r = Write-TrainingSection $ws2 $r "周五  爆发/速度日（Power）" "50-60%极限  |  爆发向心+控制离心  |  RPE 5-6  |  轻、快、弹" "35-40" $actPower

$ws2.Range("A:A").ColumnWidth = 4
$ws2.Range("B:B").ColumnWidth = 18
$ws2.Range("C:C").ColumnWidth = 7
$ws2.Range("D:D").ColumnWidth = 12
$ws2.Range("E:E").ColumnWidth = 18
$ws2.Range("F:F").ColumnWidth = 9
$ws2.Range("G:G").ColumnWidth = 7
$ws2.Range("H:H").ColumnWidth = 16
$ws2.Range("I:I").ColumnWidth = 56

Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 减载周（带无工具替代）
# ============================================
$ws3 = $wb.Worksheets.Add()
$ws3.Name = "第4周-减载"

Set-Cell $ws3 1 1 "第4周  减载周  '动但不同'" $true 16 $headerBlue 0 $true
Merge-Cells $ws3 1 1 9
$ws3.Range("A1:I1").RowHeight = 34
Set-Cell $ws3 2 1 "核心原则：不是做轻一点——是做不同的动作。换动作=关节不同应力路径。所有动作RPE 4-5，做完应感觉'没练够'。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws3 2 1 9

$r = 4
Set-Cell $ws3 $r 1 "动作替换表（含无工具备用）" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 9
$ws3.Range("A$($r):I$($r)").RowHeight = 26
$ws3.Range("A$($r):I$($r)").Font.Color = $white
$ws3.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "原动作" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "首选替换" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 4 "参数" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 5 "如果没有↑" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 7 "备选参数" $true 10 $white $headerBlue $false
$r++

$subsV5 = @(
    @("哑铃平板卧推","俯卧撑(3秒下/爆发推)","3x8-10","地上俯卧撑(标准)","3x力竭的60%"),
    @("器械坐姿推肩","弹力带推肩(站姿)","3x10-12","哑铃阿诺德推举(轻量)","3x10-12 8-10kg"),
    @("高位下拉","弹力带高位下拉(单臂)","3x10/侧","哑铃单臂划船(轻量慢离心)","3x10/侧 10-12kg"),
    @("坐姿水平划船","哑铃单臂划船(轻量)","3x10-12/侧","桌沿划船(自重，脚跟着地)","3x10-12"),
    @("哑铃分腿蹲","高脚杯深蹲(轻哑铃)","3x10次 10-14kg","自重深蹲(慢速控制)","3x15次"),
    @("哑铃臀推","自重臀桥+弹力带","3x15-20","自重单腿臀桥","3x10-12/侧"),
    @("哑铃罗马尼亚硬拉","早安式体前屈(轻哑铃)","3x10-12","站姿体前屈(自重)","3x15次"),
    @("绳索飞鸟","弹力带飞鸟(绑后方)","2x12-15","哑铃飞鸟(最轻重量)","2x15次 4-6kg"),
    @("绳索下压","弹力带下压(踩脚下单臂)","2x12/侧","窄距俯卧撑(跪姿)","2x力竭的60%"),
    @("哑铃弯举","弹力带弯举(踩脚下)","2x12次","哑铃弯举(最轻重量慢离心)","2x15次 4-6kg"),
    @("面拉","弹力带面拉","2x15次","俯身哑铃反向飞鸟(最轻)","2x15次 2-4kg")
)

foreach ($s in $subsV5) {
    Set-Cell $ws3 $r 1 $s[0] $true 9 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $s[1] $true 9 $accentGreen 0 $false
    Merge-Cells $ws3 $r 2 3
    Set-Cell $ws3 $r 4 $s[2] $false 9 0 0 $false
    Set-Cell $ws3 $r 5 $s[3] $false 9 0 0 $false
    Merge-Cells $ws3 $r 5 6
    Set-Cell $ws3 $r 7 $s[4] $false 9 0 0 $false
    Merge-Cells $ws3 $r 7 8
    $ws3.Range("A$($r):I$($r)").RowHeight = 30
    Set-Borders $ws3 $r 1 9
    $r++
}

$r += 2
Set-Cell $ws3 $r 1 "减载周规则" $true 14 $white $accentGreen $true
Merge-Cells $ws3 $r 1 9
$ws3.Range("A$($r):I$($r)").RowHeight = 26
$ws3.Range("A$($r):I$($r)").Font.Color = $white
$ws3.Range("A$($r):I$($r)").Interior.Color = $accentGreen
$r++

$drules = @(
    "  不分主项辅助——所有动作统一对待，每2-3组。",
    "  不冲重量、不力竭、不尝试新动作极限。",
    "  有氧同步减载：跑步机坡度8% / 4km/h / 10min。",
    "  多喝水、多睡觉——减载周恢复质量决定下一周期能冲多高。",
    "  结束理想感受：精力充沛、关节轻松、'想练但没练够'。",
    "  如果只有弹力带+自重也够——减载周不需要器械。"
)
foreach ($dr in $drules) {
    Set-Cell $ws3 $r 1 $dr $true 10 $accentGreen 0 $true
    Merge-Cells $ws3 $r 1 9
    $ws3.Range("A$($r):I$($r)").RowHeight = 22
    Set-Borders $ws3 $r 1 9
    $r++
}

$ws3.Range("A:A").ColumnWidth = 22
$ws3.Range("B:B").ColumnWidth = 24
$ws3.Range("C:C").ColumnWidth = 12
$ws3.Range("D:D").ColumnWidth = 12
$ws3.Range("E:E").ColumnWidth = 24
$ws3.Range("F:I").ColumnWidth = 14

Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 渐进与突破
# ============================================
$ws4 = $wb.Worksheets.Add()
$ws4.Name = "渐进与突破"

Set-Cell $ws4 1 1 "4周渐进 + 器械上限突破" $true 16 $headerBlue 0 $true
Merge-Cells $ws4 1 1 9
$ws4.Range("A1:I1").RowHeight = 34

$r = 3
Set-Cell $ws4 $r 1 "4周渐进（DUP框架内 每周3天主题不变 逐周递增）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 9
$ws4.Range("A$($r):I$($r)").RowHeight = 26
$ws4.Range("A$($r):I$($r)").Font.Color = $white
$ws4.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "强度日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "容量日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "爆发日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 8 "有氧" $true 10 $white $headerBlue $false
$r++

$prog = @(
    @("1","基准 80%/RPE8","65%/RPE6-7","50%/RPE5","12min/10%"),
    @("2","上肢+1-2.5kg 下肢+2.5-5kg RPE8-9","缩短休息5-10s 次数上限+2 RPE7","+1-2kg 保持爆发 RPE5-6","13min/10%"),
    @("3","再+1-2.5kg(上肢) +2.5-5kg(下肢) RPE8-9","维持第2周 轻松再加组 RPE7-8","维持第2周 追求快启动 RPE6","14min/12%"),
    @("4","重量50-60% 换动作 RPE4-5","换动作 组数减半 RPE4-5","换动作+弹力带 2组 RPE4","10min/8%")
)

foreach ($pw in $prog) {
    Set-Cell $ws4 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $pw[1] $false 9 0 0 $true
    Set-Cell $ws4 $r 4 $pw[2] $false 9 0 0 $true
    Set-Cell $ws4 $r 6 $pw[3] $false 9 0 0 $true
    Set-Cell $ws4 $r 8 $pw[4] $false 9 0 0 $false
    $ws4.Range("A$($r):I$($r)").RowHeight = 46
    Set-Borders $ws4 $r 1 9
    $r++
}

# Breakthrough strategies (unchanged from V4)
$r += 2
Set-Cell $ws4 $r 1 "器械上限突破（哑铃20kg/器械插片到顶时）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 9
$ws4.Range("A$($r):I$($r)").RowHeight = 26
$ws4.Range("A$($r):I$($r)").Font.Color = $white
$ws4.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "操作" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "最佳日用" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "效果" $true 10 $white $headerBlue $false
$r++

$bt = @(
    @("慢离心","3-4s下放+爆发推起","全3天","增加TUT"),
    @("1.5次法","下放→半起→全下→全起=1次","强度日/容量日","双倍刺激"),
    @("缩短休息","90s→45s","容量日","增加密度"),
    @("单侧变式","单腿RDL/单臂划船/单腿臀推","强度日","轻重量高刺激"),
    @("弹力带叠加","哑铃+弹力带双阻力","爆发日/强度日","变相加重"),
    @("预疲劳","孤立力竭→复合动作","容量日","轻量化力竭"),
    @("爆发向心","加速推/拉阶段","爆发日","提高RFD")
)

foreach ($b in $bt) {
    Set-Cell $ws4 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $b[1] $false 10 0 0 $true
    Set-Cell $ws4 $r 4 $b[2] $false 10 0 0 $false
    Set-Cell $ws4 $r 6 $b[3] $false 10 0 0 $true
    Merge-Cells $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight = 28
    Set-Borders $ws4 $r 1 9
    $r++
}

$ws4.Range("A:A").ColumnWidth = 18
$ws4.Range("B:I").ColumnWidth = 16

Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 有氧+热身+冷身+训练窗口营养
# ============================================
$ws5 = $wb.Worksheets.Add()
$ws5.Name = "有氧热身冷身营养"

Set-Cell $ws5 1 1 "有氧  热身(含神经激活)  冷身  训练前后营养" $true 16 $headerBlue 0 $true
Merge-Cells $ws5 1 1 9
$ws5.Range("A1:I1").RowHeight = 34

# Cardio
$r = 3
Set-Cell $ws5 $r 1 "有氧（力量后，跑步机爬坡；备选酒店楼梯/快走）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 9
$ws5.Range("A$($r):I$($r)").RowHeight = 26
$ws5.Range("A$($r):I$($r)").Font.Color = $white
$ws5.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "第1周" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 2 "12min / 坡度10% / 4.5-5km/h" $false 10 0 0 $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "第2周" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 6 "13min / 坡度10%" $false 10 0 0 $false
Merge-Cells $ws5 $r 6 9
Set-Borders $ws5 $r 1 9
$r++

Set-Cell $ws5 $r 1 "第3周" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 2 "14min / 坡度12% / 5km/h" $false 10 0 0 $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "第4周(减载)" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 6 "10min / 坡度8% / 4km/h" $false 10 0 0 $false
Merge-Cells $ws5 $r 6 9
Set-Borders $ws5 $r 1 9
$r++

# Warmup
$r += 2
Set-Cell $ws5 $r 1 "训练前热身（8-10分钟  必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 9
$ws5.Range("A$($r):I$($r)").RowHeight = 26
$ws5.Range("A$($r):I$($r)").Font.Color = $white
$ws5.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "步" $true 9 $white $headerBlue $false
Set-Cell $ws5 $r 2 "动作" $true 9 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "时间/次数" $true 9 $white $headerBlue $false
Set-Cell $ws5 $r 6 "目的" $true 9 $white $headerBlue $false
Merge-Cells $ws5 $r 6 9
$r++

$warmups = @(
    @("1","原地踏步+手臂画圈","1min","全身升温"),
    @("2","靠墙静蹲","30sx2组","激活股四头(零膝压)"),
    @("3","猫牛式","8次","脊柱活动"),
    @("4","肩关节环绕","前后各8","肩关节润滑"),
    @("5","摆腿(前后+左右)","各10/侧","髋关节活动度"),
    @("6","死虫式","每侧8次","核心激活"),
    @("7","弹力带肩胛激活","15次","背部激活"),
    @("8*","神经激活(选1个 30s)","30s x1-2组","唤醒神经系统——强度日/爆发日必做!")
)
foreach ($w in $warmups) {
    Set-Cell $ws5 $r 1 $w[0] $false 9 0 0 $false
    Set-Cell $ws5 $r 2 $w[1] $true 9 $headerBlue 0 $false
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $w[2] $false 9 0 0 $false
    Set-Cell $ws5 $r 6 $w[3] $false 9 0 0 $true
    Merge-Cells $ws5 $r 6 9
    $ws5.Range("A$($r):I$($r)").RowHeight = 24
    Set-Borders $ws5 $r 1 9
    $r++
}

Set-Cell $ws5 $r 1 "  *步8神经激活选项" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 2 "有弹力带→侧向走30s | 有低箱→跳箱5次(30-40cm) | 都没→原地高抬腿快速交替20s | 强度日/爆发日优先用跳箱或高抬腿，容量日用弹力带侧走" $false 10 0 0 $true
Merge-Cells $ws5 $r 2 9
$ws5.Range("A$($r):I$($r)").RowHeight = 34
$r++

# Cooldown simple
$r++
Set-Cell $ws5 $r 1 "训练后冷身（5-8分钟）：胸大肌拉伸→背阔肌拉伸→髋屈肌拉伸→腘绳肌拉伸→婴儿式→腹式深呼吸5-8次" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws5 $r 1 9
$ws5.Range("A$($r):I$($r)").RowHeight = 24
$r++

# Peri-workout nutrition
$r += 2
Set-Cell $ws5 $r 1 "训练前后营养窗口  V5强化" $true 14 $white $accentGreen $true
Merge-Cells $ws5 $r 1 9
$ws5.Range("A$($r):I$($r)").RowHeight = 26
$ws5.Range("A$($r):I$($r)").Font.Color = $white
$ws5.Range("A$($r):I$($r)").Interior.Color = $accentGreen
$r++

Set-Cell $ws5 $r 1 "什么时候" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "吃什么" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "为什么" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 5 9
$r++

$peri = @(
    @("训练前2h","香蕉+蛋白粉/全麦吐司+2蛋（碳水30-40g+蛋白15-20g）","提供训练能量，防止分解肌肉。不吃→训练打七折。"),
    @("训练中","水500-750ml小口慢喝（强度日可加电解质）","维持表现，防抽筋。"),
    @("训练后1h内","蛋白粉+香蕉+牛奶/米饭+鸡胸肉（碳水40-50g+蛋白25-30g）","补充糖原，启动修复。蛋白质吸收效率最高时段。"),
    @("强度日vs容量日","强度日：训练后碳水偏高(50g)","神经系统消耗大→需要更快补糖原。"),
    @("减载周","训练前后碳水各减10-15g","训练量小→不需要那么多能量。")
)
foreach ($pw in $peri) {
    Set-Cell $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $pw[1] $false 10 0 0 $true
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $pw[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 5 9
    $ws5.Range("A$($r):I$($r)").RowHeight = 36
    Set-Borders $ws5 $r 1 9
    $r++
}

# Purchases
$r++
Set-Cell $ws5 $r 1 "建议买：泡沫轴30-50元(训前滚大腿/训后滚背) + 弹力带15-30元(热身激活+减载周主力) = 总约80元。放行李箱不占地，减载周没弹力带就是没工具。" $false 10 $accentGreen 0 $true
Merge-Cells $ws5 $r 1 9
$ws5.Range("A$($r):I$($r)").RowHeight = 28

$ws5.Range("A:A").ColumnWidth = 20
$ws5.Range("B:I").ColumnWidth = 18

Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 膝关节+安全+RPE+营养+规则+全休
# ============================================
$ws6 = $wb.Worksheets.Add()
$ws6.Name = "保护规则RPE"

Set-Cell $ws6 1 1 "膝关节保护  RPE指南  营养总览  进阶规则  季度全休" $true 16 $headerBlue 0 $true
Merge-Cells $ws6 1 1 9
$ws6.Range("A1:I1").RowHeight = 34

# Knee
$r = 3
Set-Cell $ws6 $r 1 "膝关节保护" $true 14 $white $accentRed $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $accentRed
$r++

$knee = @(
    @("禁止","坐姿腿屈伸机/腿弯举机——4周不碰"),
    @("谨慎","分腿蹲——无痛才加重，不适换臀推+RDL"),
    @("安全","臀推/臀桥/RDL——放心练"),
    @("康复","每次热身前靠墙静蹲30sx2"),
    @("停止","锐痛/针刺感/肿胀→当天不练下肢")
)

foreach ($k in $knee) {
    Set-Cell $ws6 $r 1 $k[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $k[1] $false 10 0 0 $true
    Merge-Cells $ws6 $r 2 9
    $ws6.Range("A$($r):I$($r)").RowHeight = 22
    Set-Borders $ws6 $r 1 9
    $r++
}

# RPE Guide
$r++
Set-Cell $ws6 $r 1 "RPE自感用力指南（本方案核心工具）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "RPE" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "感觉" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "还剩几次" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "用在哪个日" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 8 "一句话" $true 10 $white $headerBlue $false
$r++

$rpeAll = @(
    @("4-5","非常轻松","8-10次+","减载周","'我今天练了什么？'——这就对了"),
    @("5-6","轻松→中等","6-8次","爆发日","轻快有力像弹簧"),
    @("6-7","中等偏强","3-4次","容量日","充血发胀像气球"),
    @("8-9","很重接近极限","1-2次","强度日","'再来一个就起不来了'——但你别起不来"),
    @("10","完全力竭","0次","本方案不用！","留1-2次余量是长期进步的秘诀")
)

foreach ($rp in $rpeAll) {
    Set-Cell $ws6 $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $rp[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $rp[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $rp[3] $false 10 0 0 $false
    Set-Cell $ws6 $r 8 $rp[4] $false 10 0 0 $true
    $ws6.Range("A$($r):I$($r)").RowHeight = 28
    Set-Borders $ws6 $r 1 9
    $r++
}

# Safety
$r++
Set-Cell $ws6 $r 1 "安全提示" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "优势：无伤病史(膝盖近期可逆) | 1.5年私教基础 | 37岁恢复能力尚好 | DUP降低适应性停滞" $false 10 $accentGreen 0 $true
Merge-Cells $ws6 $r 1 9
$r++
Set-Cell $ws6 $r 1 "注意：睡眠7-8h必须保证 | BMI28.3不加入跳跃类 | 应酬饮酒抑制蛋白合成 | 减载周必须真减载 | 季度全休不可跳" $false 10 $accentOrange 0 $true
Merge-Cells $ws6 $r 1 9
$r++
Set-Cell $ws6 $r 1 "避免：颈后推举/下拉 | 腿屈伸/腿弯举 | 空腹力量 | 减载周冲重量 | 全休周偷偷练" $false 10 $accentRed 0 $true
Merge-Cells $ws6 $r 1 9
$r++

# Nutrition summary
$r++
Set-Cell $ws6 $r 1 "营养总览" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "热量" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "蛋白质" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "碳水" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 5 "脂肪" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "训练后碳水" $true 10 $white $headerBlue $false
$r++

$nutri = @(
    @("训练周1-3","2100-2300","144-160g","230-260g","50-55g","40-50g(强度日偏高)"),
    @("减载周4","1800-2000","144-160g","160-180g","45-50g","30g"),
    @("全休周(季度)","1700-1800","130-140g","150-160g","45-50g","无特殊要求")
)
foreach ($n in $nutri) {
    Set-Cell $ws6 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $n[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 3 $n[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $n[3] $false 10 0 0 $false
    Set-Cell $ws6 $r 5 $n[4] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $n[5] $false 10 0 0 $false
    Set-Borders $ws6 $r 1 9
    $r++
}

# Rules (with emergency fallback)
$r++
Set-Cell $ws6 $r 1 "进阶判断规则（含状态极差兜底）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r++

$rulesV5 = @(
    @("强度日该加重","RPE目标8-9实际只有7","下次下肢+2.5-5kg 上肢+1-2.5kg"),
    @("强度日太重","RPE到10(力竭)或动作变形","退回上周重量——RPE优先于数字"),
    @("容量日该加量","RPE只有5 '不怎么累'","缩短休息10s或次数+2"),
    @("爆发日速度下降","推/拉速度明显变慢","减重1-2kg——速度优先于重量"),
    @("今天状态极差","起床就累/没睡好/压力大","强度→容量(RPE降为6-7)；容量→爆发(RPE降为5-6)；爆发→休息或只热身+有氧。一次降级不影响进步，一次硬撑可能停练2周"),
    @("膝盖分腿蹲不适","锐痛/针刺感","当天换臀推+RDL——完全无膝压组合"),
    @("减载周太轻松","—","  这就是对的"),
    @("恢复不过来","48h+明显酸痛持续2周","该减载了——37岁别硬撑"),
    @("体重没变力量涨","—","  最好信号！身体重组进行中"),
    @("3个月到","—","全休1周——散步/拉伸/泡沫轴，不碰器械")
)

Set-Cell $ws6 $r 1 "信号" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "条件" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 5 "做法" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 2 4
Merge-Cells $ws6 $r 5 9
$r++

foreach ($rv in $rulesV5) {
    Set-Cell $ws6 $r 1 $rv[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $rv[1] $false 10 0 0 $true
    Merge-Cells $ws6 $r 2 4
    Set-Cell $ws6 $r 5 $rv[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 5 9
    $ws6.Range("A$($r):I$($r)").RowHeight = 32
    Set-Borders $ws6 $r 1 9
    $r++
}

# Quarterly full rest
$r++
Set-Cell $ws6 $r 1 "季度全休周：每3个月1次" $true 14 $white $softPurple $true
Merge-Cells $ws6 $r 1 9
$ws6.Range("A$($r):I$($r)").RowHeight = 26
$ws6.Range("A$($r):I$($r)").Font.Color = $white
$ws6.Range("A$($r):I$($r)").Interior.Color = $softPurple
$r++

$qr = @(
    "即使有减载周，连续3个月仍累积神经系统+内分泌+关节深层疲劳——不是一周减载能消化的。",
    "全休周做：散步30-40min/天 + 泡沫轴 + 大量拉伸。不碰器械、不碰哑铃、不做力量训练。",
    "心态：这不是退步——这是为下季度蓄力。休息完力量不会掉，反而涨——因为超量恢复真正完成了。",
    "长期坚持训练的人，往往是那些懂得何时休息的人。37岁的身体尤其需要这个节奏。"
)
foreach ($q in $qr) {
    Set-Cell $ws6 $r 1 $q $false 10 0 0 $true
    Merge-Cells $ws6 $r 1 9
    $ws6.Range("A$($r):I$($r)").RowHeight = 24
    $r++
}

$ws6.Range("A:A").ColumnWidth = 24
$ws6.Range("B:I").ColumnWidth = 16

Write-Host "Sheet 6 done"

# ============================================
# SHEET 7: 备用+自主+复盘
# ============================================
$ws7 = $wb.Worksheets.Add()
$ws7.Name = "备用自主复盘"

Set-Cell $ws7 1 1 "精简备用  自主周期框架  4周复盘" $true 16 $headerBlue 0 $true
Merge-Cells $ws7 1 1 9
$ws7.Range("A1:I1").RowHeight = 34

# Simplified
$r = 3
Set-Cell $ws7 $r 1 "2练精简备用（应酬/出差只能练2次时）" $true 14 $white $accentOrange $true
Merge-Cells $ws7 $r 1 9
$ws7.Range("A$($r):I$($r)").RowHeight = 26
$ws7.Range("A$($r):I$($r)").Font.Color = $white
$ws7.Range("A$($r):I$($r)").Interior.Color = $accentOrange
$r++

Set-Cell $ws7 $r 1 "序号" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "第1练：强度+下肢" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 2 5
Set-Cell $ws7 $r 6 "第2练：容量+爆发混合" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 6 9
$r++

$simple = @(
    @("  ","哑铃卧推 4x4-6(强度)","哑铃卧推 3x10-12(容量)"),
    @("  ","高位下拉 4x4-6(强度)","高位下拉 3x10-12(容量)"),
    @("  ","哑铃分腿蹲 3x4-6","哑铃RDL 3x10-12"),
    @("  ","臀推+面拉超级组","爆发臀推+弹力带面拉"),
    @("  ","飞鸟+下压超级组","弯举+举腿"),
    @("有氧","12min坡度10%","12min坡度10%"),
    @("时长","约40min力量+12min有氧","约40min力量+12min有氧")
)
foreach ($si in $simple) {
    Set-Cell $ws7 $r 1 $si[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $si[1] $false 10 0 0 $false
    Merge-Cells $ws7 $r 2 5
    Set-Cell $ws7 $r 6 $si[2] $false 10 0 0 $false
    Merge-Cells $ws7 $r 6 9
    $ws7.Range("A$($r):I$($r)").RowHeight = 24
    Set-Borders $ws7 $r 1 9
    $r++
}

# Self-training framework
$r += 2
Set-Cell $ws7 $r 1 "学会自己做周期计划  5步框架" $true 14 $white $accentGreen $true
Merge-Cells $ws7 $r 1 9
$ws7.Range("A$($r):I$($r)").RowHeight = 26
$ws7.Range("A$($r):I$($r)").Font.Color = $white
$ws7.Range("A$($r):I$($r)").Interior.Color = $accentGreen
$r += 2

$fwV5 = @(
    @("定训练频率","每周能练几天？","2练→全身x2 | 3练→DUP | 4练→上下肢"),
    @("定DUP主题","每天不同刺激","强度日(RPE8-9)→容量日(RPE6-7)→爆发日(RPE5-6)"),
    @("选动作+标RPE","主项4组→辅助3组→孤立3组","每个动作标目标RPE，按身体反馈调整"),
    @("定周期节奏","3周训练+1周减载(换动作)+季度全休","这是最小周期单元——够简单也够有效"),
    @("每4周复盘","主要看RPE数据和力量变化","RPE比重量更诚实——RPE偏高→该减载，偏低→该推进")
)

Set-Cell $ws7 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "做什么" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 4 "关键信息" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 4 9
$r++

foreach ($fw in $fwV5) {
    Set-Cell $ws7 $r 1 $fw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $fw[1] $true 11 $headerBlue 0 $false
    Set-Cell $ws7 $r 4 $fw[2] $false 10 0 0 $true
    Merge-Cells $ws7 $r 4 9
    $ws7.Range("A$($r):I$($r)").RowHeight = 26
    Set-Borders $ws7 $r 1 9
    $r++
}

# Review
$r += 2
Set-Cell $ws7 $r 1 "4周后复盘清单" $true 14 $white $darkBg $true
Merge-Cells $ws7 $r 1 9
$ws7.Range("A$($r):I$($r)").RowHeight = 26
$ws7.Range("A$($r):I$($r)").Font.Color = $white
$ws7.Range("A$($r):I$($r)").Interior.Color = $darkBg
$r += 2

$rvAll = @(
    @("  ","RPE数据","强度日RPE稳定在8-9？容量日在6-7？RPE比重量更诚实"),
    @("  ","力量变化","三个主题日各涨了多少？哪个主题进步最明显？"),
    @("  ","体重+腰围","晨起空腹称重趋势，腰带松了几孔？"),
    @("  ","膝盖恢复","能否逐步恢复动作？分腿蹲深度加深了吗？"),
    @("  ","减载感受","结束是否精力充沛、'想练但没练够'？"),
    @("  ","适合哪个主题","强度日/容量日/爆发日——哪个做起来最舒服？下周期可微调"),
    @("  ","自主能力","看懂DUP逻辑了吗？有信心自己调整下周期？")
)
foreach ($rv in $rvAll) {
    Set-Cell $ws7 $r 1 $rv[0] $false 14 0 0 $false
    Set-Cell $ws7 $r 2 $rv[1] $true 12 $headerBlue 0 $false
    Set-Cell $ws7 $r 3 $rv[2] $false 11 0 0 $true
    Merge-Cells $ws7 $r 3 9
    $ws7.Range("A$($r):I$($r)").RowHeight = 28
    $r++
}

$ws7.Range("A:A").ColumnWidth = 10
$ws7.Range("B:I").ColumnWidth = 20

Write-Host "Sheet 7 done"

# ============================================
# Save
# ============================================
$savePath = "D:\Codex\Joe何_训练方案_V5.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()

[GC]::Collect()

Write-Host "Saved to: $savePath"
Write-Host "ALL DONE - V5 Complete"
