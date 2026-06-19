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
$lightPurple = 0xe8daef
$softPurple = 0x8e44ad
$softBlue = 0x2980b9

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

function Write-DayPlan($ws, $refRow, $title, $subtitle, $estTime, $actions) {
    Set-Cell $ws $refRow 1 $title $true 13 $white $darkBg $true
    Merge-Cells $ws $refRow 1 8
    $ws.Range("A$($refRow):H$($refRow)").RowHeight = 28
    $ws.Range("A$($refRow):H$($refRow)").Font.Color = $white
    $ws.Range("A$($refRow):H$($refRow)").Interior.Color = $darkBg

    $refRow++
    Set-Cell $ws $refRow 1 $subtitle $false 10 $headerBlue $lightGray $true
    Merge-Cells $ws $refRow 1 8
    $refRow++

    Set-Cell $ws $refRow 1 "序号" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 2 "动作" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 3 "组数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 4 "次数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 5 "重量/负载" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 6 "组间休息" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 7 "RPE" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 8 "动作要点（呼吸+纠偏+发力感）" $true 10 $white $headerBlue $true
    $ws.Range("A$($refRow):H$($refRow)").RowHeight = 28
    $refRow++

    $i = 1
    foreach ($a in $actions) {
        Set-Cell $ws $refRow 1 $i $false 10 0 $lightGray $false
        Set-Cell $ws $refRow 2 $a[0] $true 10 $headerBlue 0 $false
        Set-Cell $ws $refRow 3 $a[1] $false 10 0 0 $false
        Set-Cell $ws $refRow 4 $a[2] $false 10 0 0 $false
        Set-Cell $ws $refRow 5 $a[3] $false 9 0 0 $false
        Set-Cell $ws $refRow 6 $a[4] $false 10 0 0 $false
        Set-Cell $ws $refRow 7 $a[5] $false 10 0 0 $false
        Set-Cell $ws $refRow 8 $a[6] $false 9 0 0 $true
        $ws.Range("A$($refRow):H$($refRow)").RowHeight = 66
        Set-Borders $ws $refRow 1 8
        $refRow++
        $i++
    }
    return $refRow + 2
}

Write-Host "Setup complete"

# ============================================
# SHEET 1: 方案概览
# ============================================
$ws = $wb.Worksheets.Item(1)
$ws.Name = "方案概览"

Set-Cell $ws 1 1 "知行AI教练  定制训练方案 V4" $true 18 $headerBlue 0 $true
Merge-Cells $ws 1 1 8
$ws.Range("A1:H1").RowHeight = 38

Set-Cell $ws 2 1 "Joe  He  |  37岁  |  168cm / 80kg  |  减脂增肌  身体重组  |  波动周期(DUP)" $false 12 $gold $darkBg $true
Merge-Cells $ws 2 1 8
$ws.Range("A2:H2").RowHeight = 28
$ws.Range("A2:H2").Font.Color = $white
$ws.Range("A2:H2").Interior.Color = $darkBg

$row = 4
Set-Cell $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $row 1 8
$ws.Range("A$($row):H$($row)").RowHeight = 26

$overview = @(
    @("训练频率","每周3练（附2练精简备用方案）"),
    @("每次时长","总计55-65分钟：热身8-10min + 力量35-45min + 有氧10-15min + 拉伸5min"),
    @("训练方式","全身训练 x 3次/周，采用波动周期(DUP)——每天不同训练主题"),
    @("周期结构","3周训练 + 1周减载 = 1个小周期；3个小周期 + 1周全休 = 1个中周期"),
    @("3天主题","周一=强度日(80-85%)  |  周三=容量日(65-70%)  |  周五=爆发/速度日(50-60%)"),
    @("训练地点","酒店/小区健身房（哑铃最重20kg + 3台固定器械 + 小龙门架）"),
    @("有氧方式","跑步机爬坡走（力量训练后进行；备选：酒店楼梯/户外快走）"),
    @("RPE参考","容量日 RPE 6-7  |  强度日 RPE 8-9  |  爆发日 RPE 5-6  |  减载周 RPE 4-5"),
    @("方案特点","DUP波动周期  RPE智能调节  神经激活  减载换动作  训练前后营养窗口")
)

$r = 5
foreach ($item in $overview) {
    Set-Cell $ws $r 1 $item[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 3 $item[1] $false 10 0 0 $true
    Merge-Cells $ws $r 3 8
    Set-Borders $ws $r 1 8
    $r++
}

# DUP Structure
$r++
Set-Cell $ws $r 1 "DUP 波动周期结构（核心创新）" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 8
$r++

Set-Cell $ws $r 1 "训练日" $true 10 $white $headerBlue $false
Set-Cell $ws $r 2 "主题" $true 10 $white $headerBlue $false
Set-Cell $ws $r 3 "强度区间" $true 10 $white $headerBlue $false
Set-Cell $ws $r 4 "主项次数" $true 10 $white $headerBlue $false
Set-Cell $ws $r 5 "目标RPE" $true 10 $white $headerBlue $false
Set-Cell $ws $r 6 "辅助主题" $true 10 $white $headerBlue $false
Set-Cell $ws $r 7 "训练感觉" $true 10 $white $headerBlue $false
$r++

$dup = @(
    @("周一","强度日","80-85%","4-6次","8-9","容量辅助(3x10-12)","重而不力竭——每次做完剩1-2次力气"),
    @("周三","容量日","65-70%","10-12次","6-7","代谢辅助(高次数)","泵感强烈——肌肉充血，代谢压力大"),
    @("周五","爆发/速度日","50-60%","3-5次(爆发)","5-6","强度辅助(5x5)","轻快有力——质量>数量，感受爆发节奏")
)

foreach ($d in $dup) {
    Set-Cell $ws $r 1 $d[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 2 $d[1] $true 10 $white $darkBg $false
    Set-Cell $ws $r 3 $d[2] $false 10 0 0 $false
    Set-Cell $ws $r 4 $d[3] $false 10 0 0 $false
    Set-Cell $ws $r 5 $d[4] $false 10 0 0 $false
    Set-Cell $ws $r 6 $d[5] $false 10 0 0 $true
    Set-Cell $ws $r 7 $d[6] $false 10 0 0 $true
    $ws.Range("A$($r):H$($r)").RowHeight = 28
    Set-Borders $ws $r 1 8
    $r++
}

# Why DUP
$r++
Set-Cell $ws $r 1 "为什么选DUP而非线性周期？" $true 12 $headerBlue 0 $true
Merge-Cells $ws $r 1 8
$r++

$dupReasons = @(
    "线性周期（容量周 强度周 强度周 减载周）对初学者有效——每次只练一种刺激，身体有3-4周时间适应。",
    "Joe何有1.5年训练基础，已过新手期。线性周期容易产生适应性停滞——身体会'习惯'同一主题的连续刺激。",
    "DUP每周内轮换3种刺激（强度/容量/爆发），神经系统始终面临'新'挑战——适应性更慢，进步更持续。",
    "每次训练主题不同 = 每次都有新鲜感——心理疲劳更低，长期依从性更好。",
    "减载周结束后，身体对强度、容量、爆发3种刺激都保持敏感——恢复后能同时突破多个维度。"
)

foreach ($dr in $dupReasons) {
    Set-Cell $ws $r 1 $dr $false 10 0 0 $true
    Merge-Cells $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight = 22
    $r++
}

# Yearly framework
$r++
Set-Cell $ws $r 1 "全年周期框架" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 8
$r++

Set-Cell $ws $r 1 "层级" $true 10 $white $headerBlue $false
Set-Cell $ws $r 2 "时长" $true 10 $white $headerBlue $false
Set-Cell $ws $r 3 "结构" $true 10 $white $headerBlue $false
Merge-Cells $ws $r 3 5
Set-Cell $ws $r 6 "关键动作" $true 10 $white $headerBlue $false
Merge-Cells $ws $r 6 8
$r++

$yearly = @(
    @("大周期","1年","围绕长期目标：身体重组 80kg  70-72kg + 力量稳步提升","每季度末评估进展，调整方向"),
    @("中周期","1季度(3个月)","3个小周期(每月) + 1周全休","第3个月末全休1周——彻底重置系统"),
    @("小周期","1个月(4周)","3周DUP训练 + 1周减载(换动作)","每周3天=强度/容量/爆发轮换"),
    @("全休周","每3个月1次","只做低强度有氧+拉伸+泡沫轴，不训练","让神经系统+内分泌+关节彻底修复"),
    @("上下肢差异","持续所有周期","上肢+1-2.5kg  |  下肢+2.5-5kg","上肢精细保护，下肢潜力释放")
)

foreach ($y in $yearly) {
    Set-Cell $ws $r 1 $y[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 2 $y[1] $false 10 0 0 $false
    Set-Cell $ws $r 3 $y[2] $false 10 0 0 $true
    Merge-Cells $ws $r 3 5
    Set-Cell $ws $r 6 $y[3] $false 10 0 0 $true
    Merge-Cells $ws $r 6 8
    $ws.Range("A$($r):H$($r)").RowHeight = 26
    Set-Borders $ws $r 1 8
    $r++
}

# Core principles
$r++
Set-Cell $ws $r 1 "三大核心原则" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 8
$r++

$principles = @(
    @("1","无痛优先","膝关节不适期间，所有下肢动作以无痛为第一原则。宁可退阶不冒险。"),
    @("2","RPE优先于数字","80% 1RM是理论值，当天的睡眠/压力/饮食会改变实际表现。RPE让训练更智能——根据身体反馈实时调整，而非死磕数字。"),
    @("3","动但不同","减载周不是'做轻一点'——是'做不同的动作'。换动作=给关节不同的应力路径，保持血液循环和神经激活，同时让主力动作模式的关节彻底休息。")
)

foreach ($p in $principles) {
    Set-Cell $ws $r 1 $p[0] $false 10 0 0 $false
    Set-Cell $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws $r 2 3
    Set-Cell $ws $r 4 $p[2] $false 10 0 0 $true
    Merge-Cells $ws $r 4 8
    $ws.Range("A$($r):H$($r)").RowHeight = 32
    Set-Borders $ws $r 1 8
    $r++
}

$ws.Range("A:A").ColumnWidth = 28
$ws.Range("B:H").ColumnWidth = 16

Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 第1-3周 DUP训练计划
# ============================================
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "第1-3周-DUP训练"

Set-Cell $ws2 1 1 "第1-3周  DUP波动周期训练计划" $true 16 $headerBlue 0 $true
Merge-Cells $ws2 1 1 8
$ws2.Range("A1:H1").RowHeight = 34

Set-Cell $ws2 2 1 "每周3天，每天不同刺激主题。3周共9次训练，每周递增（见Sheet 4）。第4周减载（见Sheet 3）。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws2 2 1 8
$ws2.Range("A2:H2").RowHeight = 26

$r = 4

# ===== 周一：强度日 =====
$actStrength = @(
    @("哑铃平板卧推","4组","4-6次","单边18-20kg（80-85%极限）","120秒","8-9","【呼吸】下放吸气屏住(瓦式)，推起呼气。【纠偏】手肘与躯干45-60 不外展90 。【发力感】推起时想象把杠铃'推弯'。强度日：每组只留1-2次力竭余地。"),
    @("器械坐姿推肩","3组","4-6次","插片选4-6RM","90秒","8-9","【呼吸】推起呼气、下放吸气。【纠偏】下背贴紧靠背，不反弓借力。【发力感】推至顶端不锁死肘。强度日：上肢精细加重+1-2.5kg。"),
    @("哑铃分腿蹲","4组","4-6次/侧","单边16-18kg（强度日加重）","120秒","8-9","【呼吸】下蹲吸气、起身呼气。【纠偏】躯干竖直、膝盖无痛范围内。【发力感】前腿发力起身。 无痛才加重。强度日：下肢可+2.5-5kg。"),
    @("哑铃臀推","3组","6-8次","单颗18-20kg（强度日冲上限）","90秒","8-9","【呼吸】推起呼气、下放吸气。【纠偏】下巴微收，不超伸腰椎。【发力感】顶峰臀大肌夹紧2秒。已达20kg上限 换单腿臀推。"),
    @("绳索飞鸟+绳索下压 超级组","各3组","各10-12次","比容量日重1档","做完一对休息60秒","6-7","强度日的辅助以容量方式做——让主项冲强度，辅助积累代谢。飞鸟：手臂微屈，胸肌驱动。下压：大臂固定不外摆。"),
    @("平板支撑","3组","30-40秒","自重","30秒","—","强度日核心稳定：骨盆微后倾，腰不塌。")
)
$r = Write-DayPlan $ws2 $r "周一  强度日（Strength Day）" "主题：80-85%极限重量  主项4-6次力竭前1-2次停  |  辅助以容量方式做  |  当天关键词：重、稳、控" "42-48" $actStrength

# ===== 周三：容量日 =====
$actVolume = @(
    @("哑铃平板卧推","4组","10-12次","单边14-16kg（65-70%极限）","60秒","6-7","【呼吸】下放吸气、推起呼气。【纠偏】同周一——重量轻了动作质量更要高。【发力感】追求胸肌泵感，每次收缩1秒。容量日：重量轻了但组间更短，代谢压力更大。"),
    @("高位下拉","4组","10-12次","插片选12-15RM","60秒","6-7","【呼吸】下拉呼气、还原吸气。【纠偏】先沉肩胛再拉——容量日不要借力。【发力感】背阔肌全程控制，顶峰收缩1秒。"),
    @("哑铃罗马尼亚硬拉","4组","10-12次","单边14-18kg","60秒","6-7","【呼吸】下放吸气、拉起呼气。【纠偏】膝盖微屈，哑铃贴腿滑动。容量日：高次数RDL追求腘绳肌泵感和离心控制。"),
    @("龙门架面拉","4组","15-20次","轻重量","45秒","6-7","【呼吸】拉向面部呼气、还原吸气。【纠偏】拉向额头。容量日：高次数面拉——肩后束的代谢压力积累。"),
    @("哑铃弯举+仰卧举腿 超级组","各3组","弯举15次/举腿力竭","弯举单边6-10kg","做完一对休息45秒","7-8","容量日手臂+腹——高次数充血。弯举：杜绝摆动，离心3秒。举腿：下腹带动骨盆上卷。"),
    @("鸟狗式","3组","每侧10次","自重","30秒","—","容量日核心——对侧手脚延伸，骨盆稳定不旋转。")
)
$r = Write-DayPlan $ws2 $r "周三  容量日（Volume Day）" "主题：65-70%极限重量  组间短(45-60s)  高次数泵感  |  高强度代谢压力  |  当天关键词：泵、累、充血" "40-45" $actVolume

# ===== 周五：爆发/速度日 =====
$actPower = @(
    @("爆发哑铃卧推","6组","3次","单边10-12kg（50-60%极限）","45秒","5-6","【呼吸】下放控制吸气、爆发推起呼气。【纠偏】不要用大重量——轻重量追求推起速度。【发力感】想象把哑铃'弹'出去，爆发推起同时保持控制。爆发日核心：质量  数量。"),
    @("爆发高位下拉+慢放","5组","5次","插片选50-60%（比容量日轻）","60秒","5-6","【纠偏】爆发下拉+慢离心3秒还原——'拉快放慢'。【发力感】下拉时追求启动速度，还原时感受背阔肌拉伸。"),
    @("分腿蹲爆发起","4组","5次/侧","自重或持轻哑铃(4-6kg)","60秒","5-6","【纠偏】下蹲后爆发起身——不是跳起来，是快速伸展髋膝。【发力感】感受前腿的爆发力输出。 膝关节无冲击——不追求跳跃高度。"),
    @("爆发臀推+慢放","4组","5次","单颗10-14kg（60%强度日重量）","60秒","5-6","【纠偏】爆发推起+慢离心3秒还原。推起时想象把杠铃'弹'开。【发力感】臀大肌瞬间爆发收紧。"),
    @("弹力带面拉+弹力带推胸","各3组","各8-10次","弹力带中等阻力","45秒","5-6","爆发日弹力带收尾——面拉刺激肩后束，推胸刺激胸肌(不同于哑铃角度)。弹力带变阻特性适合速度日。"),
    @("死虫式","3组","每侧8次(慢速控制)","自重","30秒","—","爆发日核心控制——慢速高质量死虫式，感受核心抗旋转。")
)
$r = Write-DayPlan $ws2 $r "周五  爆发/速度日（Power/Speed Day）" "主题：50-60%极限重量  追求动作速率而非绝对重量  |  爆发向心+控制离心  |  当天关键词：快、轻、弹" "35-40" $actPower

$ws2.Range("A:A").ColumnWidth = 4
$ws2.Range("B:B").ColumnWidth = 20
$ws2.Range("C:C").ColumnWidth = 8
$ws2.Range("D:D").ColumnWidth = 12
$ws2.Range("E:E").ColumnWidth = 20
$ws2.Range("F:F").ColumnWidth = 12
$ws2.Range("G:G").ColumnWidth = 8
$ws2.Range("H:H").ColumnWidth = 56

Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 减载周（换动作）
# ============================================
$ws3 = $wb.Worksheets.Add()
$ws3.Name = "第4周-减载周"

Set-Cell $ws3 1 1 "第4周  减载周  动但不同" $true 16 $headerBlue 0 $true
Merge-Cells $ws3 1 1 8
$ws3.Range("A1:H1").RowHeight = 34

Set-Cell $ws3 2 1 "减载周核心原则：不是'做轻一点'——是'做不同的动作'。换动作=给关节不同的应力路径，保持血液循环和神经激活，同时让主力动作模式的关节彻底休息。RPE 4-5，结束应感觉'没练够'。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws3 2 1 8
$ws3.Range("A2:H2").RowHeight = 30

$r = 4

# Deload substitutions
Set-Cell $ws3 $r 1 "动作替换规则" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 26
$ws3.Range("A$($r):H$($r)").Font.Color = $white
$ws3.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "原动作" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "减载周替换" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 4 "替换理由" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 6 "参数" $true 10 $white $headerBlue $false
$r++

$subs = @(
    @("哑铃平板卧推","俯卧撑（控制节奏：3秒下/爆发推）","肩关节不同受力角度，腕关节压力分散","3组x8-10次 RPE 4"),
    @("器械坐姿推肩","弹力带推肩（站姿，弹力带踩脚下）","自由轨迹激活肩胛稳定肌群","3组x10-12次 RPE 4"),
    @("高位下拉","弹力带高位下拉（单臂，站姿）","单臂=纠正左右不平衡，弹力带=不同阻力曲线","3组x10次/侧 RPE 4"),
    @("坐姿水平划船","TRX/弹力带反向划船（如有）/哑铃单臂划船轻量","身体角度改变=中背部不同区域受力","3组x10-12次/侧 RPE 4"),
    @("哑铃分腿蹲","高脚杯深蹲（轻哑铃10-14kg）","双脚对称=膝关节应力分布不同","3组x10次 RPE 4"),
    @("哑铃臀推","自重臀桥+弹力带臀桥","去掉负重，关注臀肌激活和顶峰收缩质量","3组x15-20次 RPE 4"),
    @("哑铃罗马尼亚硬拉","早安式体前屈（轻哑铃抱胸前）","髋铰链模式不变，负荷向量改变=下背压力不同","3组x10-12次 RPE 4"),
    @("绳索飞鸟","弹力带飞鸟（站姿，弹力带绑后方）","不同角度，弹力带变阻=顶端更难","2组x12-15次 RPE 4"),
    @("绳索下压","弹力带下压（踩脚下，单臂）","单臂纠正左右三头肌力量平衡","2组x12次/侧 RPE 4"),
    @("哑铃弯举","弹力带弯举（踩脚下）","弹力带变阻曲线=顶端收缩更充分","2组x12次 RPE 4")
)

foreach ($s in $subs) {
    Set-Cell $ws3 $r 1 $s[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $s[1] $true 10 $accentGreen 0 $false
    Set-Cell $ws3 $r 4 $s[2] $false 10 0 0 $true
    Set-Cell $ws3 $r 6 $s[3] $false 10 0 0 $false
    $ws3.Range("A$($r):H$($r)").RowHeight = 34
    Set-Borders $ws3 $r 1 8
    $r++
}

# Deload week schedule
$r += 2
Set-Cell $ws3 $r 1 "减载周3天安排（仍保持3练，维持习惯节奏）" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 26
$ws3.Range("A$($r):H$($r)").Font.Color = $white
$ws3.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "减载周一（原强度日位置）" $true 11 $softBlue $lightBlue $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 24
$r++

Set-Cell $ws3 $r 1 "做6个动作，每个2-3组，RPE 4-5。从替换动作中选——上半身偏推力+下肢。重点是感受身体，不是完成数字。" $false 10 0 0 $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 22
$r++

Set-Cell $ws3 $r 1 "减载周三（原容量日位置）" $true 11 $softBlue $lightBlue $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 24
$r++

Set-Cell $ws3 $r 1 "做6个动作，每个2-3组，RPE 4-5。从替换动作中选——上半身偏拉力+下肢。注意关节的轻松感。" $false 10 0 0 $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 22
$r++

Set-Cell $ws3 $r 1 "减载周五（原爆发日位置）" $true 11 $softBlue $lightBlue $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 24
$r++

Set-Cell $ws3 $r 1 "做4-5个动作，每个2组，RPE 4。以弹力带动作为主+拉伸。练完应该感觉'我今天练了什么？'——这就是减载周的理想感受。" $false 10 0 0 $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 22
$r++

# Deload rules
$r += 2
Set-Cell $ws3 $r 1 "减载周规则" $true 14 $white $accentGreen $true
Merge-Cells $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight = 26
$ws3.Range("A$($r):H$($r)").Font.Color = $white
$ws3.Range("A$($r):H$($r)").Interior.Color = $accentGreen
$r++

$deloadRules = @(
    @("  不分主项辅助——所有动作统一对待，每个动作2-3组。"),
    @("  不冲重量、不力竭、不尝试新动作的极限——减载周不是'试试新纪录'的时候。"),
    @("  有氧同步减载：跑步机坡度8%，速度4km/h，10分钟——轻松走。"),
    @("  多喝水、多睡觉——减载周的恢复质量决定了下一周期你能冲多高。"),
    @("  结束时的理想感受：精力充沛、关节轻松、'想练但没练够'。")
)

foreach ($dr in $deloadRules) {
    Set-Cell $ws3 $r 1 $dr $true 10 $accentGreen 0 $true
    Merge-Cells $ws3 $r 1 8
    $ws3.Range("A$($r):H$($r)").RowHeight = 24
    Set-Borders $ws3 $r 1 8
    $r++
}

$r++
Set-Cell $ws3 $r 1 "  减载周结束后，你应该带着更强的身体储备进入下一周期。'退'是为了更好的'进'——这是超量恢复的真正秘密。" $false 10 $accentGreen 0 $true
Merge-Cells $ws3 $r 1 8

$ws3.Range("A:A").ColumnWidth = 24
$ws3.Range("B:B").ColumnWidth = 28
$ws3.Range("C:C").ColumnWidth = 18
$ws3.Range("D:H").ColumnWidth = 16

Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 4周渐进 + 器械突破
# ============================================
$ws4 = $wb.Worksheets.Add()
$ws4.Name = "渐进与突破"

Set-Cell $ws4 1 1 "4周渐进负荷  +  器械上限突破策略" $true 16 $headerBlue 0 $true
Merge-Cells $ws4 1 1 8
$ws4.Range("A1:H1").RowHeight = 34

$r = 3
Set-Cell $ws4 $r 1 "4周渐进节奏（DUP框架内，每周3天主题不变，逐周递增）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight = 26
$ws4.Range("A$($r):H$($r)").Font.Color = $white
$ws4.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "主题" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 3 "强度日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "容量日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 5 "爆发日变化" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "有氧" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 7 "RPE范围" $true 10 $white $headerBlue $false
$r++

$progWeeks = @(
    @("第1周","基准周","80%极限  RPE 8`n建立强度基准","65%极限  RPE 6-7`n建立容量基准","50%极限  RPE 5`n建立速度基准","12min`n坡度10%","5-8"),
    @("第2周","递增周","+1-2.5kg(上肢)`n+2.5-5kg(下肢)`nRPE 8-9 争取更多","缩短休息5-10秒`n次数上限+2次`nRPE 7 更累","+1-2kg 保持爆发`nRPE 5-6","13min`n坡度10%","6-9"),
    @("第3周","冲击周","再+1-2.5kg(上肢)`n+2.5-5kg(下肢)`nRPE 8-9 接近极限","维持第2周容量`n如果轻松再加组`nRPE 7-8","维持第2周`n追求更快启动`nRPE 6","14min`n坡度12%","7-9"),
    @("第4周","减载周","重量50-60%`n换动作(见Sheet3)`nRPE 4-5","换动作`n组数减半`nRPE 4-5","换动作+弹力带`n2组足够`nRPE 4","10min`n坡度8%","4-5")
)

foreach ($pw in $progWeeks) {
    Set-Cell $ws4 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $pw[1] $true 10 $white $darkBg $false
    Set-Cell $ws4 $r 3 $pw[2] $false 9 0 0 $true
    Set-Cell $ws4 $r 4 $pw[3] $false 9 0 0 $true
    Set-Cell $ws4 $r 5 $pw[4] $false 9 0 0 $true
    Set-Cell $ws4 $r 6 $pw[5] $false 9 0 0 $false
    Set-Cell $ws4 $r 7 $pw[6] $false 10 0 0 $false
    $ws4.Range("A$($r):H$($r)").RowHeight = 52
    Set-Borders $ws4 $r 1 8
    $r++
}

# Breakthrough strategies
$r += 2
Set-Cell $ws4 $r 1 "器械上限突破策略（哑铃20kg/器械插片到顶时）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight = 26
$ws4.Range("A$($r):H$($r)").Font.Color = $white
$ws4.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "操作" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "最佳使用日" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "效果" $true 10 $white $headerBlue $false
$r++

$btStrats = @(
    @("慢离心训练","3-4秒下放+爆发推起","所有3天","增加TUT，轻重量达到增肌效果"),
    @("1.5次法","下放 起一半 再下放 完全起身=1次","强度日、容量日","单次双倍刺激，代谢压力翻倍"),
    @("缩短休息","从90s/60s逐步缩到45s","容量日（首选）","增加训练密度，提升代谢压力"),
    @("单侧变式","单腿RDL、单臂划船、单腿臀推","强度日（首选）","轻重量高刺激，稳定肌同时工作"),
    @("弹力带叠加","哑铃+弹力带双阻力","爆发日、强度日","变相增加负重，顶峰阻力更大"),
    @("预疲劳法","孤立动作力竭  复合动作","容量日进阶","轻重量也能让复合动作接近力竭"),
    @("爆发式向心","加速推/拉阶段","爆发日（核心）","提高发力率(RFD)，轻重量高神经输出")
)

foreach ($b in $btStrats) {
    Set-Cell $ws4 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $b[1] $false 10 0 0 $true
    Set-Cell $ws4 $r 4 $b[2] $false 10 0 0 $false
    Set-Cell $ws4 $r 6 $b[3] $false 10 0 0 $true
    Merge-Cells $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight = 32
    Set-Borders $ws4 $r 1 8
    $r++
}

$ws4.Range("A:A").ColumnWidth = 22
$ws4.Range("B:H").ColumnWidth = 18

Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 有氧+热身(含神经激活)+冷身+购置+训练前后营养
# ============================================
$ws5 = $wb.Worksheets.Add()
$ws5.Name = "有氧热身冷身营养窗口"

Set-Cell $ws5 1 1 "有氧安排  热身(含神经激活)  冷身  训练前后营养窗口" $true 16 $headerBlue 0 $true
Merge-Cells $ws5 1 1 8
$ws5.Range("A1:H1").RowHeight = 34

# Cardio
$r = 3
Set-Cell $ws5 $r 1 "有氧安排（力量训练后，每周3次）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight = 26
$ws5.Range("A$($r):H$($r)").Font.Color = $white
$ws5.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "时长" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "坡度/速度" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 5 "强度感觉" $true 10 $white $headerBlue $false
$r++

$cardioDup = @(
    @("第1周","12分钟","坡度10% / 4.5-5km/h","能断续说话，不能流畅聊天"),
    @("第2周","13分钟","坡度10% / 4.5-5km/h","维持——有氧不做剧烈变化"),
    @("第3周","14分钟","坡度12% / 5km/h","微提坡度——增加消耗"),
    @("第4周（减载）","10分钟","坡度8% / 4km/h","轻松走——有氧也跟着减载")
)

foreach ($cd in $cardioDup) {
    Set-Cell $ws5 $r 1 $cd[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $cd[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 3 $cd[2] $false 10 0 0 $false
    Set-Cell $ws5 $r 5 $cd[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 5 8
    Set-Borders $ws5 $r 1 8
    $r++
}

# Warmup with neural activation
$r += 2
Set-Cell $ws5 $r 1 "训练前热身（8-10分钟  必须做  V4新增神经激活）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight = 26
$ws5.Range("A$($r):H$($r)").Font.Color = $white
$ws5.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "动作" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "时间/次数" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 6 "目的" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 6 8
$r++

$warmups_v4 = @(
    @("1","原地踏步+手臂画圈","1分钟","全身升温，关节润滑"),
    @("2","靠墙静蹲","30秒 x 2组","温和激活股四头肌，零膝关节压力——膝盖不适期间必备"),
    @("3","猫牛式","8次","脊柱逐节活动"),
    @("4","肩关节环绕","前后各8次","肩关节囊润滑"),
    @("5","站姿摆腿（前后+左右）","各10次/侧","髋关节活动度"),
    @("6","死虫式","每侧8次","核心激活"),
    @("7","弹力带肩胛激活","15次","背部激活"),
    @("8  V4新增","神经激活动作（选1个，30秒）`n  弹力带侧向走（臀中肌+下肢稳定）`n  跳箱（低高度30-40cm，5次）——强度日和爆发日优先`n  药球砸地（如有）——全身爆发力输出`n  如果都没有：原地高抬腿快速交替20秒","30秒 x 1-2组","唤醒神经系统——尤其强度日和爆发日必备。37岁的身体，神经需要'预热'才能输出最大力量。")
)

foreach ($w in $warmups_v4) {
    Set-Cell $ws5 $r 1 $w[0] $true 9 0 0 $false
    Set-Cell $ws5 $r 2 $w[1] $true 9 $headerBlue 0 $false
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $w[2] $false 9 0 0 $false
    Set-Cell $ws5 $r 6 $w[3] $false 9 0 0 $true
    Merge-Cells $ws5 $r 6 8
    $ws5.Range("A$($r):H$($r)").RowHeight = if ($r -eq 42) { 56 } else { 26 }
    Set-Borders $ws5 $r 1 8
    $r++
}

# Cooldown
$r += 2
Set-Cell $ws5 $r 1 "训练后冷身（5-8分钟  必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight = 26
$ws5.Range("A$($r):H$($r)").Font.Color = $white
$ws5.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

$cooldowns = @(
    @("1","胸大肌拉伸","每侧30秒","卧推/飞鸟后"),
    @("2","背阔肌拉伸","每侧30秒","下拉/划船后"),
    @("3","半跪姿髋屈肌拉伸","每侧30秒","分腿蹲后"),
    @("4","坐姿腘绳肌拉伸","30秒","RDL后"),
    @("5","婴儿式","30秒","脊柱减压"),
    @("6","腹式深呼吸","5-8次","副交感激活，促进恢复")
)
foreach ($c in $cooldowns) {
    Set-Cell $ws5 $r 1 $c[0] $false 10 0 0 $false
    Set-Cell $ws5 $r 2 $c[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $c[2] $false 10 0 0 $false
    Set-Cell $ws5 $r 6 $c[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 6 8
    $ws5.Range("A$($r):H$($r)").RowHeight = 24
    Set-Borders $ws5 $r 1 8
    $r++
}

# Peri-workout nutrition (NEW)
$r += 2
Set-Cell $ws5 $r 1 "训练前后营养窗口  V4新增" $true 14 $white $accentGreen $true
Merge-Cells $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight = 26
$ws5.Range("A$($r):H$($r)").Font.Color = $white
$ws5.Range("A$($r):H$($r)").Interior.Color = $accentGreen
$r++

Set-Cell $ws5 $r 1 "时机" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "内容" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 4 "逻辑" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 4 8
$r++

$periWorkout = @(
    @("训练前2小时","碳水30-40g + 蛋白质15-20g`n例：一根香蕉+一杯蛋白粉`n或：两片全麦吐司+2个蛋","提供训练能量、防止训练中分解肌肉`n训练前不吃  训练质量打七折"),
    @("训练中","水+电解质（尤其强度日）`n500-750ml水，小口慢喝","维持表现、防止抽筋`n强度日出汗多更要补"),
    @("训练后1小时","碳水40-50g + 蛋白质25-30g`n例：米饭一碗+鸡胸肉`n或：蛋白粉一勺+香蕉+牛奶","补充糖原、启动肌肉修复`n训练后窗口是蛋白质吸收效率最高的时段"),
    @("强度日 vs 容量日差异","强度日：训练后碳水可更高(50g)`n容量日：训练后碳水40g，前餐碳水30g","强度日神经系统消耗大→需要更快补充糖原`n容量日代谢压力大→总碳水够就好"),
    @("减载周","训练前后餐碳水均减少10-15g","训练量小了→不需要那么多能量补充")
)

foreach ($pw in $periWorkout) {
    Set-Cell $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $pw[1] $false 10 0 0 $true
    Set-Cell $ws5 $r 4 $pw[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight = if ($r -lt 50) { 44 } else { 36 }
    Set-Borders $ws5 $r 1 8
    $r++
}

# Purchases
$r += 2
Set-Cell $ws5 $r 1 "建议购置（总花费约50-80元）" $true 14 $white $accentGreen $true
Merge-Cells $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight = 26
$ws5.Range("A$($r):H$($r)").Font.Color = $white
$ws5.Range("A$($r):H$($r)").Interior.Color = $accentGreen
$r++

$purchases = @(
    @("泡沫轴（30-50元）","训练前滚股四头肌、髂胫束(大腿外侧)、腘绳肌各30秒——下肢日必做，大腿前/外侧过紧常是膝痛根源。训练后滚胸椎、上背各30秒。"),
    @("弹力带（15-30元）","热身肩胛激活15次、臀桥+膝外撑15次。减载周主力器械——几乎所有动作都可用弹力带做替换。"),
    @("总约50-80元","放行李箱不占地方。对恢复和关节健康的价值远超价格——几十块解决大问题。")
)

foreach ($p in $purchases) {
    Set-Cell $ws5 $r 1 $p[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $p[1] $false 10 0 0 $true
    Merge-Cells $ws5 $r 2 8
    $ws5.Range("A$($r):H$($r)").RowHeight = 40
    Set-Borders $ws5 $r 1 8
    $r++
}

$ws5.Range("A:A").ColumnWidth = 22
$ws5.Range("B:H").ColumnWidth = 18

Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 膝关节+安全+营养+规则+全休周
# ============================================
$ws6 = $wb.Worksheets.Add()
$ws6.Name = "保护营养规则"

Set-Cell $ws6 1 1 "膝关节保护  安全提示  营养建议  RPE指南  进阶规则  季度全休" $true 16 $headerBlue 0 $true
Merge-Cells $ws6 1 1 8
$ws6.Range("A1:H1").RowHeight = 34

# Knee
$r = 3
Set-Cell $ws6 $r 1 "膝关节保护专项" $true 14 $white $accentRed $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $accentRed
$r++

$knee = @(
    @("  禁止","坐姿腿屈伸机","膝关节前侧剪切力","4周内完全不碰"),
    @("  禁止","坐姿腿弯举机","上次练完出现膝盖不适","4周内完全不碰"),
    @("  谨慎","分腿蹲(强度日)","膝盖无痛才加重","不适时换臀推+RDL"),
    @("  安全","臀推/臀桥/RDL","膝关节几乎不参与","所有主题日放心练"),
    @("  康复","靠墙静蹲(热身前)","零膝压激活股四头肌","每次训练前30sx2"),
    @("  停止信号","锐痛/针刺感/肿胀","—","当天不做下肢动作")
)

Set-Cell $ws6 $r 1 "级别" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "动作" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "原因" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "处理" $true 10 $white $headerBlue $false
$r++

foreach ($k in $knee) {
    Set-Cell $ws6 $r 1 $k[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $k[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws6 $r 4 $k[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $k[3] $false 10 0 0 $false
    $ws6.Range("A$($r):H$($r)").RowHeight = 26
    Set-Borders $ws6 $r 1 8
    $r++
}

# Safety
$r += 2
Set-Cell $ws6 $r 1 "安全提示" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "  优势：无伤病史（膝盖近期可逆） | 1.5年私教基础 | 37岁恢复能力尚好 | DUP波动周期降低适应性停滞" $false 10 $accentGreen 0 $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 22
$r++

Set-Cell $ws6 $r 1 "  注意：睡眠7-8h必须保证 | BMI28.3不加入跳跃类 | 应酬饮酒抑制蛋白合成 | 减载周必须真减载 | 每3个月全休1周不可跳过" $false 10 $accentOrange 0 $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 28
$r++

Set-Cell $ws6 $r 1 "  避免：颈后推举/下拉 | 腿屈伸/腿弯举 | 空腹力量 | 减载周冲重量 | 季度全休周偷偷练" $false 10 $accentRed 0 $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 28
$r++

# RPE Guide
$r++
Set-Cell $ws6 $r 1 "RPE 自感用力程度指南  V4新增核心工具" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "RPE" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "感觉描述" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "还能做几次" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "用在哪个日子" $true 10 $white $headerBlue $false
$r++

$rpeGuide = @(
    @("4-5","非常轻松，毫无压力","还能做8-10次+","减载周专用——就是这种感觉"),
    @("5-6","轻松到中等强度","还能做6-8次","爆发日——轻而快，感受速度"),
    @("6-7","中等偏强，有挑战","还能做3-4次","容量日——泵感强烈但不过度力竭"),
    @("8-9","很重，接近极限","还能做1-2次","强度日——重但有力，不做到力竭"),
    @("10","绝对极限，完全力竭","0次——再做1次都做不到","本方案不使用！留1-2次余量是长期进步的关键")
)

foreach ($rpe in $rpeGuide) {
    Set-Cell $ws6 $r 1 $rpe[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $rpe[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $rpe[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $rpe[3] $false 10 0 0 $false
    $ws6.Range("A$($r):H$($r)").RowHeight = 26
    Set-Borders $ws6 $r 1 8
    $r++
}

$r++
Set-Cell $ws6 $r 1 "为什么用RPE？80% 1RM是理论值，你当天的睡眠、压力、饮食会改变实际表现。RPE让你根据身体反馈实时调整——感觉好可以微上，感觉差主动微下。死磕百分比容易'硬撑'或'偷懒'。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 28
$r++

# Nutrition summary
$r++
Set-Cell $ws6 $r 1 "营养建议总览（通用方向，非处方）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "总热量" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "蛋白质" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "碳水" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 5 "脂肪" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "训练后碳水" $true 10 $white $headerBlue $false
$r++

$nutriAll = @(
    @("训练周(1-3)","2100-2300","144-160g","230-260g","50-55g","40-50g（强度日偏高）"),
    @("减载周(4)","1800-2000","144-160g","160-180g","45-50g","30g"),
    @("全休周(每季度)","1700-1800","130-140g","150-160g","45-50g","无特殊要求")
)

foreach ($na in $nutriAll) {
    Set-Cell $ws6 $r 1 $na[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $na[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 3 $na[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $na[3] $false 10 0 0 $false
    Set-Cell $ws6 $r 5 $na[4] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $na[5] $false 10 0 0 $false
    Set-Borders $ws6 $r 1 8
    $r++
}

$r++
Set-Cell $ws6 $r 1 "三餐实操（生活化，不要求称克数）" $true 12 $headerBlue 0 $true
Merge-Cells $ws6 $r 1 8
$r++

Set-Cell $ws6 $r 1 "餐次" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "习惯" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "调整" $true 10 $white $headerBlue $true
Merge-Cells $ws6 $r 3 8
$r++

$meals = @(
    @("早餐(居家)","麦片+牛奶+1.5蛋","  保持！换2个整蛋。面包/馒头换全麦。"),
    @("午餐(食堂)","围桌炒菜","米饭一拳，先肉后菜，油多的沥一下，七分饱停。"),
    @("晚餐(居家)","跟家人吃","米饭减半/换杂粮，一份手掌大肉，蔬菜不限。"),
    @("训练前2h","—","  香蕉+蛋白粉/全麦吐司+蛋。训练前不吃  质量打折。"),
    @("训练后1h","蛋白粉一勺","  加香蕉+牛奶——碳水+蛋白双补。强度日碳水多加根香蕉。"),
    @("应酬(约2次/周)","—","当天午餐减法。喝酒前吃蛋白垫底。80%执行=成功。")
)

foreach ($m in $meals) {
    Set-Cell $ws6 $r 1 $m[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $m[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 3 $m[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 3 8
    $ws6.Range("A$($r):H$($r)").RowHeight = 34
    Set-Borders $ws6 $r 1 8
    $r++
}

# Rules
$r += 2
Set-Cell $ws6 $r 1 "进阶判断规则（按DUP主题 + RPE判断）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "信号" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "判断条件" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 2 4
Set-Cell $ws6 $r 5 "做法" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 5 8
$r++

$rulesV4 = @(
    @("  强度日该加重了","目标RPE 8-9，实际感觉只有7，轻松完成所有组","下肢+2.5-5kg，上肢+1-2.5kg"),
    @("  强度日太重了","RPE实际达到10(力竭)或动作变形","退回上周重量，RPE优先于数字"),
    @("  容量日该加量了","目标RPE 6-7，实际只有5，'不怎么累'","缩短休息10秒，或次数+2"),
    @("  爆发日速度下降","推/拉的速度明显变慢（不是轻快感）","减重1-2kg——速度优先于重量"),
    @("  减载周感觉太轻松","—","  这是对的。轻松是减载周的目标。"),
    @("  膝盖分腿蹲不适","锐痛/针刺感","当天换完全无膝压组合"),
    @("  恢复不过来","连续2周训练后48h+明显酸痛","该减载了——37岁不要硬撑"),
    @("  体重没变力量涨","—","  最好信号！身体重组进行中。"),
    @("  体重和力量双降","—","  热量缺口太大，加碳水100-150kcal"),
    @("  季度全休周到期","3个月周期结束","  1周完全不练，让系统彻底重置")
)

foreach ($rv in $rulesV4) {
    Set-Cell $ws6 $r 1 $rv[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $rv[1] $false 10 0 0 $true
    Merge-Cells $ws6 $r 2 4
    Set-Cell $ws6 $r 5 $rv[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 5 8
    $ws6.Range("A$($r):H$($r)").RowHeight = 36
    Set-Borders $ws6 $r 1 8
    $r++
}

# Quarterly full rest
$r += 2
Set-Cell $ws6 $r 1 "季度全休周（每3个月1次）  V4新增" $true 14 $white $softPurple $true
Merge-Cells $ws6 $r 1 8
$ws6.Range("A$($r):H$($r)").RowHeight = 26
$ws6.Range("A$($r):H$($r)").Font.Color = $white
$ws6.Range("A$($r):H$($r)").Interior.Color = $softPurple
$r++

$quarterRest = @(
    "即使有减载周，连续3个月的周期训练仍会累积系统性疲劳——神经系统、内分泌系统、关节的深层疲劳不是一周减载能完全消化的。",
    "每3个月安排1周完全休息（或只做散步、游泳、瑜伽、拉伸+泡沫轴）——这是'大减量'，不是'小减载'。",
    "全休周做什么：散步30-40min/天 + 泡沫轴放松 + 大量拉伸。不碰哑铃、不碰器械、不做力量训练。",
    "全休周心态：这不是'退步'——这是为下一季度'蓄力'。37岁的身体尤其需要这个节奏。长期坚持训练的人，往往是那些懂得何时休息的人。",
    "全休周结束后：身体和神经系统都'渴望'重新训练——这就是新季度周期的最佳起点。"
)

foreach ($qr in $quarterRest) {
    Set-Cell $ws6 $r 1 $qr $false 10 0 0 $true
    Merge-Cells $ws6 $r 1 8
    $ws6.Range("A$($r):H$($r)").RowHeight = 26
    $r++
}

$ws6.Range("A:A").ColumnWidth = 22
$ws6.Range("B:H").ColumnWidth = 18

Write-Host "Sheet 6 done"

# ============================================
# SHEET 7: 备用+自主计划+复盘
# ============================================
$ws7 = $wb.Worksheets.Add()
$ws7.Name = "备用自主复盘"

Set-Cell $ws7 1 1 "精简备用方案  自主周期框架  4周复盘清单" $true 16 $headerBlue 0 $true
Merge-Cells $ws7 1 1 8
$ws7.Range("A1:H1").RowHeight = 34

# Simplified
$r = 3
Set-Cell $ws7 $r 1 "2练精简备用（每周只能练2次时使用）" $true 14 $white $accentOrange $true
Merge-Cells $ws7 $r 1 8
$ws7.Range("A$($r):H$($r)").RowHeight = 26
$ws7.Range("A$($r):H$($r)").Font.Color = $white
$ws7.Range("A$($r):H$($r)").Interior.Color = $accentOrange
$r++

Set-Cell $ws7 $r 1 "精简原则：保留核心复合动作，压缩辅助动作，保留DUP逻辑（两天  强度+容量  或  强度+爆发）" $false 10 0 0 $true
Merge-Cells $ws7 $r 1 8
$r++

Set-Cell $ws7 $r 1 "序号" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "第1练：强度+下肢混合" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 2 4
Set-Cell $ws7 $r 5 "第2练：容量+爆发混合" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 5 8
$r++

$simpleItems = @(
    @("  ","哑铃卧推 4x4-6(强度)","哑铃卧推 3x10-12(容量)"),
    @("  ","高位下拉 4x4-6(强度)","高位下拉 3x10-12(容量)"),
    @("  ","哑铃分腿蹲 3x4-6","哑铃RDL 3x10-12"),
    @("  ","哑铃臀推+面拉超级组","爆发臀推+弹力带面拉超级组"),
    @("  ","绳索飞鸟+下压","弯举+举腿"),
    @("有氧","跑步机12min 坡度10%","跑步机12min 坡度10%"),
    @("时长","约40min力量+12min有氧","约40min力量+12min有氧")
)

foreach ($si in $simpleItems) {
    Set-Cell $ws7 $r 1 $si[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $si[1] $false 10 0 0 $false
    Merge-Cells $ws7 $r 2 4
    Set-Cell $ws7 $r 5 $si[2] $false 10 0 0 $false
    Merge-Cells $ws7 $r 5 8
    $ws7.Range("A$($r):H$($r)").RowHeight = 24
    Set-Borders $ws7 $r 1 8
    $r++
}

# Self-training framework
$r += 2
Set-Cell $ws7 $r 1 "学会自己做周期计划  自主训练框架" $true 14 $white $accentGreen $true
Merge-Cells $ws7 $r 1 8
$ws7.Range("A$($r):H$($r)").RowHeight = 26
$ws7.Range("A$($r):H$($r)").Font.Color = $white
$ws7.Range("A$($r):H$($r)").Interior.Color = $accentGreen
$r += 2

$frameworkV4 = @(
    @("第1步","定训练频率+分化","每周能练几天？","2练 全身x2  |  3练 全身x3(DUP)  |  4练 上下肢  |  5-6练 推拉腿"),
    @("第2步","定周期主题","DUP：每天不同刺激","强度日(80-85%) 容量日(65-70%) 爆发日(50-60%)——轮换"),
    @("第3步","选动作+定RPE","按优先级从高到低","主项4组 辅助3组 孤立3组。每个动作标上目标RPE"),
    @("第4步","定周期节奏","3+1+季度全休","3周训练+1周减载(换动作)+每3个月1周全休"),
    @("第5步","每4周复盘","用RPE数据评估","RPE比重量更诚实——RPE偏高 该减载，偏低 该推进")
)

Set-Cell $ws7 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "做什么" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 3 "核心问题" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 4 "关键信息" $true 10 $white $headerBlue $true
Merge-Cells $ws7 $r 4 8
$r++

foreach ($fw in $frameworkV4) {
    Set-Cell $ws7 $r 1 $fw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $fw[1] $true 11 $headerBlue 0 $false
    Set-Cell $ws7 $r 3 $fw[2] $false 10 0 0 $false
    Set-Cell $ws7 $r 4 $fw[3] $false 10 0 0 $true
    Merge-Cells $ws7 $r 4 8
    $ws7.Range("A$($r):H$($r)").RowHeight = 28
    Set-Borders $ws7 $r 1 8
    $r++
}

# Review
$r += 2
Set-Cell $ws7 $r 1 "4周后复盘清单" $true 14 $white $darkBg $true
Merge-Cells $ws7 $r 1 8
$ws7.Range("A$($r):H$($r)").RowHeight = 26
$ws7.Range("A$($r):H$($r)").Font.Color = $white
$ws7.Range("A$($r):H$($r)").Interior.Color = $darkBg
$r += 2

$reviews = @(
    @("  ","RPE数据对比","强度日RPE是否稳定在8-9？容量日RPE是否在6-7？RPE比重量更诚实。"),
    @("  ","力量变化","3个主题日的重量和次数各涨了多少？哪个主题进步最明显？"),
    @("  ","体重+腰围","晨起空腹称重趋势，腰带松了几孔？"),
    @("  ","膝盖恢复","能否逐步恢复动作？分腿蹲深度有没有加深？"),
    @("  ","减载周感受","结束后是否精力充沛、'想练但没练够'？——超量恢复的信号。"),
    @("  ","哪个主题日最适合自己","强度日/容量日/爆发日——哪个做完感觉最好？下周期可微调比例。"),
    @("  ","自主能力","能不能看懂DUP逻辑？有信心自己调整下周期？")
)

foreach ($rv in $reviews) {
    Set-Cell $ws7 $r 1 $rv[0] $false 14 0 0 $false
    Set-Cell $ws7 $r 2 $rv[1] $true 12 $headerBlue 0 $false
    Set-Cell $ws7 $r 3 $rv[2] $false 11 0 0 $true
    Merge-Cells $ws7 $r 3 8
    $ws7.Range("A$($r):H$($r)").RowHeight = 32
    $r++
}

$ws7.Range("A:A").ColumnWidth = 10
$ws7.Range("B:B").ColumnWidth = 22
$ws7.Range("C:H").ColumnWidth = 18

Write-Host "Sheet 7 done"

# ============================================
# Save
# ============================================
$savePath = "D:\Codex\Joe何_训练方案_V4.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()

for ($i = 1; $i -le 7; $i++) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb.Worksheets.Item($i)) | Out-Null
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()

Write-Host "Saved to: $savePath"
Write-Host "ALL DONE - V4 Complete"
