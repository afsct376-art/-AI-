$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()

# Colors
$darkBg = 0x1a1a2e
$gold = 0xd4a574
$white = 0xffffff
$lightGray = 0xf5f5f5
$midGray = 0xe8e8e8
$headerBlue = 0x2c3e50
$accentGreen = 0x27ae60
$accentRed = 0xe74c3c
$accentOrange = 0xf39c12
$lightBlue = 0xd6eaf8

function Set-Cell($ws, $row, $col, $value, $bold, $size, $color, $bg, $wrap) {
    $cell = $ws.Cells.Item($row, $col)
    $cell.Value = $value
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

# ============================================
# SHEET 1
# ============================================
$ws = $wb.Worksheets.Item(1)
$ws.Name = "方案概览"

Set-Cell $ws 1 1 "知行AI教练 · 定制训练方案" $true 18 $headerBlue 0 $true
Merge-Cells $ws 1 1 7
$ws.Range("A1:G1").RowHeight = 40

Set-Cell $ws 2 1 "Joe 何  |  37岁  |  168cm / 80kg  |  减脂增肌 · 身体重组" $false 12 $gold $darkBg $true
Merge-Cells $ws 2 1 7
$ws.Range("A2:G2").RowHeight = 28
$ws.Range("A2:G2").Font.Color = $white
$ws.Range("A2:G2").Interior.Color = $darkBg

$row = 4
Set-Cell $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $row 1 7
$ws.Range("A$($row):G$($row)").RowHeight = 26

$overview = @(
    @("训练频率","每周3练（附2练精简备用方案）"),
    @("每次时长","总计55-65分钟：热身8min + 力量35-45min + 有氧10-15min + 拉伸5min"),
    @("训练方式","全身训练 * 3次/周（每次不同主项侧重：A推力主导 / B拉力主导 / C综合补弱）"),
    @("本方案周期","4周（第4周结束后复盘，根据数据调整下一周期）"),
    @("训练地点","酒店/小区健身房（哑铃最重20kg + 3台固定器械 + 小龙门架）"),
    @("有氧方式","跑步机爬坡走（力量训练后进行；备选：酒店楼梯/户外快走）"),
    @("方案特点","无杠铃方案 - 哑铃上限突破策略 - 膝关节保护 - 应酬场景适配 - 附自主做计划框架")
)

$r = 5
foreach ($item in $overview) {
    Set-Cell $ws $r 1 $item[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 3 $item[1] $false 10 0 0 $true
    Merge-Cells $ws $r 3 7
    Set-Borders $ws $r 1 7
    $r++
}

$r++
Set-Cell $ws $r 1 "每周训练安排" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 7
$r++

Set-Cell $ws $r 1 "训练日" $true 10 $white $headerBlue $false
Set-Cell $ws $r 3 "类型" $true 10 $white $headerBlue $false
Merge-Cells $ws $r 3 4
Set-Cell $ws $r 5 "重点肌群" $true 10 $white $headerBlue $true
Merge-Cells $ws $r 5 7
$r++

$schedule = @(
    @("周一","全身A（推力主导）","胸大肌 / 三角肌前中束 / 肱三头肌 / 股四头肌 / 核心"),
    @("周三","全身B（拉力主导）","背阔肌 / 菱形肌 / 三角肌后束 / 肱二头肌 / 腘绳肌 / 臀"),
    @("周五","全身C（综合补弱）","上胸 / 背阔肌不同角度 / 下肢容量 / 手臂 / 核心稳定性"),
    @("周二/四/六/日","休息或轻量活动","散步、拉伸、泡沫轴放松——不做高强度训练")
)

foreach ($day in $schedule) {
    Set-Cell $ws $r 1 $day[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 3 $day[1] $false 10 0 0 $false
    Set-Cell $ws $r 5 $day[2] $false 10 0 0 $true
    Merge-Cells $ws $r 5 7
    Set-Borders $ws $r 1 7
    $r++
}

$r++
Set-Cell $ws $r 1 "本方案三大核心原则" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 7
$r++

Set-Cell $ws $r 1 "序号" $true 10 $white $headerBlue $false
Set-Cell $ws $r 2 "原则" $true 10 $white $headerBlue $false
Merge-Cells $ws $r 2 3
Set-Cell $ws $r 4 "说明" $true 10 $white $headerBlue $false
Merge-Cells $ws $r 4 7
$r++

$principles = @(
    @("1","无痛优先","膝盖不适期间，所有下肢动作以无痛为第一原则。宁可退阶不冒险。"),
    @("2","80%执行率就是成功","一周21顿饭，2顿应酬只占10%。不应酬的日子认真做，应酬时不过度自责。"),
    @("3","重质不重量","哑铃20kg上限——用慢离心、1.5次法、缩短休息突破。刺激质量 > 绝对重量。")
)

foreach ($p in $principles) {
    Set-Cell $ws $r 1 $p[0] $false 10 0 0 $false
    Set-Cell $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws $r 2 3
    Set-Cell $ws $r 4 $p[2] $false 10 0 0 $true
    Merge-Cells $ws $r 4 7
    Set-Borders $ws $r 1 7
    $r++
}

$ws.Range("A:A").ColumnWidth = 24
$ws.Range("B:B").ColumnWidth = 16
$ws.Range("C:G").ColumnWidth = 14

Write-Host "Sheet 1 done"

# ============================================
# SHEET 2
# ============================================
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "第1周训练计划"

Set-Cell $ws2 1 1 "第1周 · 训练计划（适应期）" $true 16 $headerBlue 0 $true
Merge-Cells $ws2 1 1 7
$ws2.Range("A1:G1").RowHeight = 34

Set-Cell $ws2 2 1 "第1周目标：重建动作模式，使用60-65%极限重量。每个动作最后一组可以做到接近力竭，感受当前真实水平。记录所有动作的重量和次数作为4周后的对比基准。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws2 2 1 7
$ws2.Range("A2:G2").RowHeight = 32

function Write-DayPlan($ws, $refRow, $title, $focus, $estTime, $actions) {
    Set-Cell $ws $refRow 1 $title $true 13 $white $darkBg $true
    Merge-Cells $ws $refRow 1 7
    $ws.Range("A$($refRow):G$($refRow)").RowHeight = 28
    $ws.Range("A$($refRow):G$($refRow)").Font.Color = $white
    $ws.Range("A$($refRow):G$($refRow)").Interior.Color = $darkBg

    $refRow++
    Set-Cell $ws $refRow 1 "重点肌群：$focus  |  力量部分预估：约$estTime 分钟" $false 10 $headerBlue $lightGray $true
    Merge-Cells $ws $refRow 1 7
    $refRow++

    Set-Cell $ws $refRow 1 "序号" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 2 "动作" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 3 "组数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 4 "次数" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 5 "重量建议" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 6 "组间休息" $true 10 $white $headerBlue $false
    Set-Cell $ws $refRow 7 "动作要点（呼吸 + 纠偏 + 发力感）" $true 10 $white $headerBlue $true
    $ws.Range("A$($refRow):G$($refRow)").RowHeight = 28
    $refRow++

    $i = 1
    foreach ($a in $actions) {
        Set-Cell $ws $refRow 1 $i $false 10 0 $lightGray $false
        Set-Cell $ws $refRow 2 $a[0] $true 10 $headerBlue 0 $false
        Set-Cell $ws $refRow 3 $a[1] $false 10 0 0 $false
        Set-Cell $ws $refRow 4 $a[2] $false 10 0 0 $false
        Set-Cell $ws $refRow 5 $a[3] $false 10 0 0 $false
        Set-Cell $ws $refRow 6 $a[4] $false 10 0 0 $false
        Set-Cell $ws $refRow 7 $a[5] $false 9 0 0 $true
        $ws.Range("A$($refRow):G$($refRow)").RowHeight = 68
        Set-Borders $ws $refRow 1 7
        $refRow++
        $i++
    }

    return $refRow + 2
}

$r = 4

# --- 全身A ---
$actA = @(
    @("哑铃平板卧推","4组","8-10次","单边16-18kg（最后一组可冲20kg）","90秒","【呼吸】下放吸气、推起呼气。【纠偏】手肘与躯干45-60度，不要外展90度——减少肩关节压力。【发力感】脚踩实地，肩胛收紧，推起时想象把哑铃'推弯'，胸肌向中间挤压。"),
    @("器械坐姿推肩","3组","10-12次","插片选8-10RM","60秒","【呼吸】推起呼气、下放吸气。【纠偏】下背贴紧靠背，不要反弓借力。【发力感】推到顶端三角肌收紧但不锁死肘关节。"),
    @("哑铃分腿蹲","3组","10-12次/侧","单边12-16kg（双手持哑铃）","90秒","【呼吸】下蹲吸气、起身呼气。【纠偏】躯干竖直、前脚全掌踩实、深度控制在膝盖无痛范围内。【发力感】前腿股四头主导发力，重心放脚跟。⚠️ 膝关节保护：不追求深蹲深度，无痛优先。"),
    @("哑铃臀推","3组","12-15次","单颗16-20kg放髋部","60秒","【呼吸】推起呼气、下放吸气。【纠偏】下巴微收，不要过度伸展腰椎。【发力感】顶峰臀大肌用力夹紧1-2秒。✅ 零膝关节压力。"),
    @("绳索飞鸟+绳索下压 超级组","各3组","各12-15次","龙门架轻重量","飞鸟→下压→休息60秒","【飞鸟】手臂微屈，想象抱一棵大树，胸肌驱动而非手臂拉动。【下压】大臂固定身体两侧，压到底微微外旋手腕。⚡ 超级组：飞鸟做完直接做下压→休息60秒→重复。")
)
$r = Write-DayPlan $ws2 $r "全身A · 推力主导" "胸大肌 / 三角肌前中束 / 肱三头肌 / 股四头肌 / 核心" "38-42" $actA

# --- 全身B ---
$actB = @(
    @("高位下拉","4组","8-10次","插片选8-10RM","75秒","【呼吸】下拉呼气、还原吸气。【纠偏】启动时肩胛骨先下沉再用手臂拉——不要用手臂硬拽。【发力感】想象把杆拉向锁骨，肘向下后方驱动，顶峰背阔肌收紧。"),
    @("坐姿水平划船","3组","10-12次","插片选8-10RM","60秒","【呼吸】拉向身体呼气、还原吸气。【纠偏】身体不要过多后仰借力，拉到腹部而非胸口。【发力感】拉到顶端肩胛骨向中间夹紧，中背部收缩。"),
    @("哑铃罗马尼亚硬拉","3组","10-12次","单边16-20kg","75秒","【呼吸】下放吸气、拉起呼气。【纠偏】膝盖微屈不弯曲过多，哑铃贴大腿前侧下滑，全程直背。【发力感】下放时感受腘绳肌拉伸，臀向后推。✅ 膝关节几乎不动。"),
    @("龙门架面拉","3组","12-15次","轻重量","45秒","【呼吸】拉向面部呼气、还原吸气。【纠偏】拉向额头而非脖子——保护肩关节。【发力感】肩后束+上背发力，手肘向外打开。"),
    @("哑铃弯举+仰卧举腿 超级组","各3组","弯举10-12次 / 举腿力竭","弯举单边8-12kg","弯举→举腿→休息60秒","【弯举】大臂固定不摆动，顶峰微微外旋手腕。【举腿】用下腹带动骨盆上卷，不靠惯性摆动。")
)
$r = Write-DayPlan $ws2 $r "全身B · 拉力主导" "背阔肌 / 菱形肌 / 三角肌后束 / 肱二头肌 / 腘绳肌 / 臀" "38-42" $actB

# --- 全身C ---
$actC = @(
    @("哑铃上斜卧推","3组","10-12次","单边14-18kg","75秒","【呼吸】同平板卧推。【纠偏】凳角30-45度，太高会变成肩推。【发力感】上胸主导，下放至胸锁骨水平位置。"),
    @("高位下拉（反握/窄握变式）","3组","10-12次","比宽握轻1-2片","60秒","【变式目的】反握更偏向下背阔肌，不同角度刺激。【呼吸+纠偏】同宽握，但侧重感受下背阔肌的收缩。"),
    @("哑铃分腿蹲","3组","12-15次/侧","单边12-16kg","75秒","【要点同全身A】比全身A多2-3次——增加代谢压力。⚠️ 深度仍在膝盖无痛范围内。"),
    @("单腿哑铃罗马尼亚硬拉","3组","8-10次/侧","单边10-14kg","75秒","【纠偏】支撑腿膝盖微屈，髋为轴后推。【发力感】腘绳肌+臀+踝关节稳定同时训练。✅ 膝关节友好，单腿更考验稳定性和核心控制。"),
    @("绳索下压+绳索弯举 超级组","各3组","各12-15次","轻重量","下压→弯举→休息60秒","三头+二头同时充血，手臂日式收尾。💡 如果当天时间特别紧（<55分钟），可削减此超级组。"),
    @("平板支撑","3组","30-45秒","自重","30秒","【纠偏】腰不要塌，骨盆微微后倾。【发力感】腹肌主动收紧，想象有人要打你肚子——全身绷紧。")
)
$r = Write-DayPlan $ws2 $r "全身C · 综合补弱" "上胸 / 背阔肌不同角度 / 下肢容量 / 手臂 / 核心稳定性" "40-45" $actC

$ws2.Range("A:A").ColumnWidth = 6
$ws2.Range("B:B").ColumnWidth = 22
$ws2.Range("C:C").ColumnWidth = 8
$ws2.Range("D:D").ColumnWidth = 14
$ws2.Range("E:E").ColumnWidth = 22
$ws2.Range("F:F").ColumnWidth = 14
$ws2.Range("G:G").ColumnWidth = 58

Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 第2-4周渐进负荷
# ============================================
$ws3 = $wb.Worksheets.Add()
$ws3.Name = "第2-4周渐进负荷"

Set-Cell $ws3 1 1 "渐进负荷计划 · 第2-4周" $true 16 $headerBlue 0 $true
Merge-Cells $ws3 1 1 6
$ws3.Range("A1:F1").RowHeight = 34
Set-Cell $ws3 2 1 "第1周是适应的基准线，第2-4周按以下节奏逐步加码。每个训练日都沿用第1周的动作框架，变化的是重量/次数/休息时间。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws3 2 1 6

$r = 4
Set-Cell $ws3 $r 1 "每周策略总览" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 6
$ws3.Range("A$($r):F$($r)").RowHeight = 26
$ws3.Range("A$($r):F$($r)").Font.Color = $white
$ws3.Range("A$($r):F$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 3 "具体操作" $true 10 $white $headerBlue $true
Merge-Cells $ws3 $r 3 6
$r++

$weeks = @(
    @("第1周","适应期","用60-65%极限重量，重建动作模式。每个动作最后一组可做到接近力竭（感受真实水平）。记录所有动作的重量和次数——这是4周后的对比基准。"),
    @("第2周","重量递增","主项动作（卧推、下拉、分腿蹲、RDL）在第1周基础上加重量：哑铃每边+1-2kg，器械+1片。辅助动作先不加重量，保证动作质量。如果膝盖在分腿蹲时仍有不适，下肢不加重量。"),
    @("第3周","容量递增","重量维持第2周水平，但每组目标次数取上限（如8-10→全部做到10次）。组间休息缩短10-15秒（如90秒→75秒）。如果轻松完成所有组目标次数上限，说明可以进入第4周加重。"),
    @("第4周","强度冲刺","主项动作再次加重量（哑铃+1-2kg，器械+1片）。次数回到下限（如8-10→先做到8次）。记录第4周数据——这是下一个4周周期的基准线。如果哑铃已经到20kg上限→用突破策略。")
)

foreach ($w in $weeks) {
    Set-Cell $ws3 $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $w[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws3 $r 3 $w[2] $false 10 0 0 $true
    Merge-Cells $ws3 $r 3 6
    $ws3.Range("A$($r):F$($r)").RowHeight = 52
    Set-Borders $ws3 $r 1 6
    $r++
}

$r += 2
Set-Cell $ws3 $r 1 "器械上限突破策略（当哑铃/器械重量达到上限时）" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 6
$ws3.Range("A$($r):F$($r)").RowHeight = 26
$ws3.Range("A$($r):F$($r)").Font.Color = $white
$ws3.Range("A$($r):F$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "操作" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 3 "适用动作" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 4 "效果" $true 10 $white $headerBlue $true
Merge-Cells $ws3 $r 4 6
$r++

$strategies = @(
    @("慢离心训练","3秒下放 + 爆发推起","卧推、下拉、分腿蹲","增加肌肉受力时间(TUT)，轻重量也能达到增肌效果"),
    @("1.5次法","下放→起一半→再下放→完全起身 = 1次","分腿蹲、臀推","单次动作双倍刺激，代谢压力翻倍"),
    @("缩短组间休息","从90秒逐步缩到60秒","所有动作","增加代谢压力和训练密度，提高做功效率"),
    @("增加次数","从8-10次进阶到12-15次","所有动作","从力量区间过渡到增肌区间，不同适应路径"),
    @("单侧变式","单腿RDL、单臂划船、单腿臀推","下肢、背部","增加稳定肌群参与，弥补绝对重量不足")
)

foreach ($s in $strategies) {
    Set-Cell $ws3 $r 1 $s[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $s[1] $false 10 0 0 $true
    Set-Cell $ws3 $r 3 $s[2] $false 10 0 0 $false
    Set-Cell $ws3 $r 4 $s[3] $false 10 0 0 $true
    Merge-Cells $ws3 $r 4 6
    $ws3.Range("A$($r):F$($r)").RowHeight = 38
    Set-Borders $ws3 $r 1 6
    $r++
}

$r += 2
Set-Cell $ws3 $r 1 "各动作逐周变化速查" $true 14 $white $darkBg $true
Merge-Cells $ws3 $r 1 6
$ws3.Range("A$($r):F$($r)").RowHeight = 26
$ws3.Range("A$($r):F$($r)").Font.Color = $white
$ws3.Range("A$($r):F$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws3 $r 1 "动作" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "第1周" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 3 "第2周" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 4 "第3周" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 5 "第4周" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 6 "上限策略" $true 10 $white $headerBlue $false
$r++

$exercises = @(
    @("哑铃卧推","16-18kg`n8-10次","+1-2kg/边`n8-10次","维持，做到10次","再+1-2kg`n8-10次","达20kg→慢离心3s"),
    @("器械推肩","8-10RM`n10-12次","+1片`n10-12次","维持，做到12次","再+1片`n10-12次","已到顶→缩至45s"),
    @("分腿蹲","12-16kg`n10-12次","⚠️ 无痛才加`n+1-2kg","维持/+2次`n12-14次","⚠️ 谨慎加重","达20kg→1.5次法"),
    @("哑铃臀推","16-20kg`n12-15次","+1-2kg`n12-15次","维持，做到15次","再+1-2kg","达20kg→单腿臀推"),
    @("高位下拉","8-10RM`n8-10次","+1片`n8-10次","维持，做到10次","再+1片`n8-10次","已到顶→慢离心3s"),
    @("坐姿划船","8-10RM`n10-12次","+1片`n10-12次","维持，做到12次","再+1片","已到顶→缩短休息"),
    @("哑铃RDL","16-20kg`n10-12次","+1-2kg/边`n10-12次","维持，做到12次","再+1-2kg","达20kg→单腿RDL"),
    @("辅助动作","轻重量`n12-15次","维持`n12-15次","缩短休息10s","做到15次","不追求大重量")
)

foreach ($ex in $exercises) {
    Set-Cell $ws3 $r 1 $ex[0] $true 9 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $ex[1] $false 9 0 0 $true
    Set-Cell $ws3 $r 3 $ex[2] $false 9 0 0 $true
    Set-Cell $ws3 $r 4 $ex[3] $false 9 0 0 $true
    Set-Cell $ws3 $r 5 $ex[4] $false 9 0 0 $true
    Set-Cell $ws3 $r 6 $ex[5] $false 9 0 0 $true
    $ws3.Range("A$($r):F$($r)").RowHeight = 45
    Set-Borders $ws3 $r 1 6
    $r++
}

$ws3.Range("A:A").ColumnWidth = 18
$ws3.Range("B:B").ColumnWidth = 18
$ws3.Range("C:C").ColumnWidth = 18
$ws3.Range("D:D").ColumnWidth = 18
$ws3.Range("E:E").ColumnWidth = 18
$ws3.Range("F:F").ColumnWidth = 22

Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 有氧+热身+冷身+建议购置
# ============================================
$ws4 = $wb.Worksheets.Add()
$ws4.Name = "有氧热身冷身"

Set-Cell $ws4 1 1 "有氧安排 · 热身 · 冷身 · 建议购置" $true 16 $headerBlue 0 $true
Merge-Cells $ws4 1 1 7
$ws4.Range("A1:G1").RowHeight = 34

$r = 3
Set-Cell $ws4 $r 1 "有氧安排（力量训练后进行）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "项目" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "内容" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 2 5
Set-Cell $ws4 $r 6 "备注" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 6 7
$r++

$cardio = @(
    @("方式","跑步机爬坡走（首选）","备选：酒店楼梯爬楼梯 / 户外绕酒店快走"),
    @("频率","每周3次（每次训练后）","如果某周只练2次→照样2次有氧"),
    @("时长","10min起步，每周+2min","第1周10min→第4周16min"),
    @("坡度/速度","坡度10-12%，速度4.5-5.5km/h","根据心率调整——不是越快越好"),
    @("强度","心率120-140bpm","能断续说话但不能流畅聊天"),
    @("选这个强度的原因","低强度稳态有氧(LISS)不产生过多中枢疲劳，不牺牲力量表现，同时持续贡献热量消耗——减脂增肌首选","")
)

foreach ($c in $cardio) {
    Set-Cell $ws4 $r 1 $c[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $c[1] $false 10 0 0 $true
    Merge-Cells $ws4 $r 2 5
    Set-Cell $ws4 $r 6 $c[2] $false 10 0 0 $true
    Merge-Cells $ws4 $r 6 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 32
    Set-Borders $ws4 $r 1 7
    $r++
}

# Warmup
$r += 2
Set-Cell $ws4 $r 1 "训练前热身（8分钟 · 必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "动作" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 2 4
Set-Cell $ws4 $r 5 "时间/次数" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "目的" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 6 7
$r++

$warmups = @(
    @("1","原地踏步+手臂画圈","1分钟","全身升温，关节润滑"),
    @("2","靠墙静蹲","30秒 x 2组","温和激活股四头肌，零膝关节压力——膝盖不适期间必备"),
    @("3","猫牛式","8次","脊柱逐节活动，释放腰背部紧张"),
    @("4","肩关节环绕","前后各8次","肩关节囊润滑，预防推类动作的肩部不适"),
    @("5","站姿摆腿（前后+左右）","各10次/侧","髋关节活动度——深蹲和硬拉的底层保障"),
    @("6","死虫式","每侧8次","核心激活——保护下背的关键"),
    @("7","弹力带肩胛激活（拉开+收回）","15次","激活中下斜方肌，下拉/划船的预热"),
    @("8","主项动作轻重量热身","1组 x 12次","神经适应+动作模式预热，用正式重量的40-50%")
)

foreach ($w in $warmups) {
    Set-Cell $ws4 $r 1 $w[0] $false 10 0 0 $false
    Set-Cell $ws4 $r 2 $w[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws4 $r 2 4
    Set-Cell $ws4 $r 5 $w[2] $false 10 0 0 $false
    Set-Cell $ws4 $r 6 $w[3] $false 10 0 0 $true
    Merge-Cells $ws4 $r 6 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 28
    Set-Borders $ws4 $r 1 7
    $r++
}

# Cooldown
$r += 2
Set-Cell $ws4 $r 1 "训练后冷身（5-8分钟 · 必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws4 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "动作" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 2 4
Set-Cell $ws4 $r 5 "时间" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 6 "针对" $true 10 $white $headerBlue $false
Merge-Cells $ws4 $r 6 7
$r++

$cooldowns = @(
    @("1","胸大肌拉伸（扶墙侧向伸展）","每侧30秒","卧推/飞鸟后胸肌放松"),
    @("2","背阔肌拉伸（扶墙侧屈）","每侧30秒","下拉/划船后背部放松"),
    @("3","半跪姿髋屈肌拉伸","每侧30秒","分腿蹲后的髋前侧放松"),
    @("4","坐姿腘绳肌拉伸（直腿前屈）","30秒","RDL后大腿后侧放松"),
    @("5","婴儿式（跪姿前趴）","30秒","脊柱减压，全身放松"),
    @("6","腹式深呼吸","5-8次","激活副交感神经，促进恢复，降低皮质醇")
)

foreach ($c in $cooldowns) {
    Set-Cell $ws4 $r 1 $c[0] $false 10 0 0 $false
    Set-Cell $ws4 $r 2 $c[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws4 $r 2 4
    Set-Cell $ws4 $r 5 $c[2] $false 10 0 0 $false
    Set-Cell $ws4 $r 6 $c[3] $false 10 0 0 $true
    Merge-Cells $ws4 $r 6 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 28
    Set-Borders $ws4 $r 1 7
    $r++
}

# Recommended purchases
$r += 2
Set-Cell $ws4 $r 1 "建议购置：泡沫轴 + 弹力带（总花费约50-80元）" $true 14 $white $accentGreen $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $accentGreen
$r++

Set-Cell $ws4 $r 1 "物品" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "价格" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 3 "用途与使用方法" $true 10 $white $headerBlue $true
Merge-Cells $ws4 $r 3 7
$r++

$purchases = @(
    @("泡沫轴（光滑面/中等硬度）","30-50元","【训练前】滚股四头肌、髂胫束、腘绳肌各30秒——尤其下肢日，大腿前侧/外侧过紧常是膝痛的根源。【训练后】滚胸椎、上背各30秒。酸痛点停住深呼吸10-15秒，让肌肉松开。"),
    @("弹力带（15-25磅）","15-30元","【热身时】肩胛激活15次（双手拉开弹力带）、臀桥+膝盖外撑15次（弹力带绑膝上）。【拉伸时】仰卧腘绳肌拉伸（弹力带套脚底辅助）、或辅助引体向上。"),
    @("放进行李箱不占地方","总约50-80元","酒店和家里都能用。对恢复和关节健康的价值远超这个价格——几十块解决大问题。")
)

foreach ($p in $purchases) {
    Set-Cell $ws4 $r 1 $p[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $p[1] $false 10 0 0 $false
    Set-Cell $ws4 $r 3 $p[2] $false 10 0 0 $true
    Merge-Cells $ws4 $r 3 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 48
    Set-Borders $ws4 $r 1 7
    $r++
}

$ws4.Range("A:A").ColumnWidth = 10
$ws4.Range("B:B").ColumnWidth = 24
$ws4.Range("C:C").ColumnWidth = 16
$ws4.Range("D:D").ColumnWidth = 16
$ws4.Range("E:E").ColumnWidth = 16
$ws4.Range("F:F").ColumnWidth = 16
$ws4.Range("G:G").ColumnWidth = 16

Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 膝关节保护+安全+营养+进阶规则
# ============================================
$ws5 = $wb.Worksheets.Add()
$ws5.Name = "保护饮食规则"

Set-Cell $ws5 1 1 "膝关节保护 · 安全提示 · 营养建议 · 进阶判断规则" $true 16 $headerBlue 0 $true
Merge-Cells $ws5 1 1 7
$ws5.Range("A1:G1").RowHeight = 34

# Knee
$r = 3
Set-Cell $ws5 $r 1 "膝关节保护专项 ⚠️ 重要！" $true 14 $white $accentRed $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $accentRed
$r++

Set-Cell $ws5 $r 1 "级别" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "动作" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "原因" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 3 6
Set-Cell $ws5 $r 7 "怎么办" $true 10 $white $headerBlue $false
$r++

$knee = @(
    @("🔴 禁止","坐姿腿屈伸机","膝关节前侧剪切力——你上次的不适就是相关训练后出现的","4周内完全不碰"),
    @("🔴 禁止","坐姿腿弯举机","上次练完这个后膝盖出现不适","4周内完全不碰"),
    @("🟡 谨慎","哑铃分腿蹲","从浅深度开始，每周试探性加深1-2cm，无痛才继续","哪个训练日膝盖感觉不好→当天换臀推+RDL组合"),
    @("🟢 安全","臀推/臀桥","膝关节几乎不参与","放心练，可以加重"),
    @("🟢 安全","罗马尼亚硬拉","膝关节几乎不动（只微屈），压力在髋关节","腘绳肌+臀部的核心动作，放心练"),
    @("🟢 康复","靠墙静蹲（每次热身前）","温和激活股四头肌，零膝关节压力","每次训练前30秒x2组"),
    @("⚠️ 停止信号","膝盖出现锐痛/针刺感/肿胀","—","当天不再做任何下肢动作。酸痛/紧绷感→正常，可继续但要降强度")
)

foreach ($k in $knee) {
    Set-Cell $ws5 $r 1 $k[0] $true 10 0 0 $false
    Set-Cell $ws5 $r 2 $k[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws5 $r 3 $k[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 3 6
    Set-Cell $ws5 $r 7 $k[3] $false 10 0 0 $true
    $ws5.Range("A$($r):G$($r)").RowHeight = 38
    Set-Borders $ws5 $r 1 7
    $r++
}

# Safety
$r += 2
Set-Cell $ws5 $r 1 "安全提示" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "✅ 你的优势" $true 10 $headerBlue $lightGray $false
Set-Cell $ws5 $r 2 "无伤病史（膝盖不适是近期可逆的） | 1.5年私教基础，动作模式已有基础 | 37岁尚有较好的恢复能力和睾酮水平" $false 10 0 0 $true
Merge-Cells $ws5 $r 2 7
Set-Borders $ws5 $r 1 7
$r++

Set-Cell $ws5 $r 1 "⚠️ 需要注意" $true 10 $accentOrange $lightGray $false
Set-Cell $ws5 $r 2 "37岁恢复速度不等于20岁——保证7-8小时睡眠 | 体重80kg/BMI 28.3——跳跃类动作不加入，保护膝关节 | 应酬饮酒抑制蛋白质合成——应酬日前后训练强度下调 | 酒店健身房无人保护——哑铃最后一组不做到力竭" $false 10 0 0 $true
Merge-Cells $ws5 $r 2 7
$ws5.Range("A$($r):G$($r)").RowHeight = 42
Set-Borders $ws5 $r 1 7
$r++

Set-Cell $ws5 $r 1 "❌ 建议避免" $true 10 $accentRed $lightGray $false
Set-Cell $ws5 $r 2 "颈后推举/颈后下拉（肩关节不必要风险） | 坐姿腿屈伸/腿弯举（膝盖恢复前） | 空腹力量训练 | 哑铃模拟大重量杠铃深蹲（没深蹲架导致下背风险大）" $false 10 0 0 $true
Merge-Cells $ws5 $r 2 7
$ws5.Range("A$($r):G$($r)").RowHeight = 42
Set-Borders $ws5 $r 1 7
$r++

# Nutrition
$r += 2
Set-Cell $ws5 $r 1 "营养建议（通用方向，非处方）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "当前状况：80kg，体脂25-30%，BMI 28.3。目标减脂增肌。过往从85kg减到72kg说明你能减，但回升到80kg说明维持策略需要调整。本方案不搞激进减脂，做可持续的身体重组。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 30
$r++

Set-Cell $ws5 $r 1 "营养素" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "训练日" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "非训练日" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 3 4
Set-Cell $ws5 $r 5 "说明" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 5 7
$r++

$nutrients = @(
    @("总热量","2100-2300 kcal","1800-2000 kcal","约300-500kcal缺口，不激进——优先保证训练质量"),
    @("蛋白质","144-160g","同训练日","1.8-2.0g/kg体重——蛋白质在任何日子都不能减"),
    @("碳水","230-260g","160-180g","训练日碳水集中在训练前后；非训练日自然做减法"),
    @("脂肪","50-55g","45-50g","控制但不极端，保证激素水平正常"),
    @("饮水","3-4L/天","同训练日","应酬日前后尤其多喝水，加速代谢")
)

foreach ($n in $nutrients) {
    Set-Cell $ws5 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $n[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 3 $n[2] $false 10 0 0 $false
    Merge-Cells $ws5 $r 3 4
    Set-Cell $ws5 $r 5 $n[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 5 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 26
    Set-Borders $ws5 $r 1 7
    $r++
}

$r++
Set-Cell $ws5 $r 1 "三餐实操（生活化调整，不要求每天称克数）" $true 12 $headerBlue 0 $true
Merge-Cells $ws5 $r 1 7
$r++

Set-Cell $ws5 $r 1 "餐次" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "你现在的习惯" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "可以这样调整" $true 10 $white $headerBlue $true
Merge-Cells $ws5 $r 3 7
$r++

$meals = @(
    @("早餐（居家）","麦片40g+牛奶250ml+1.5个蛋","✅ 保持！换成2个整蛋，蛋白质量提到约30g。面包/馒头时优先选全麦。"),
    @("午餐（食堂围桌）","食堂炒菜","米饭控在一拳大小。夹菜次序：肉/鱼/豆腐→蔬菜→主食。油多的菜先沥一下再夹。吃到七分饱停筷。"),
    @("晚餐（居家）","跟家人一起吃","米饭减半或换杂粮。保证一份手掌大瘦肉/鱼。蔬菜不限量。家人做菜油多→自己盛出来前控油。"),
    @("练后加餐","一勺蛋白粉","✅ 保持！加一根香蕉(便利店买)，补充肌糖原促进恢复。"),
    @("应酬（约每周2次）","—","当天午餐做减法（少油少碳水）。喝酒前吃蛋白质垫肚子。能不喝就不喝——白酒比啤酒好（啤酒热量炸弹）。应酬后多喝水。核心心态：80%执行率就是成功。一周21顿饭，2顿吃'不好'只占10%——不构成失败。")
)

foreach ($m in $meals) {
    Set-Cell $ws5 $r 1 $m[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $m[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 3 $m[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 3 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 48
    Set-Borders $ws5 $r 1 7
    $r++
}

# Advance judgment
$r += 2
Set-Cell $ws5 $r 1 "进阶判断规则（教你自己做决策）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "信号" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "判断条件" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "做法" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 5 7
$r++

$rules = @(
    @("✅ 可以加重","连续2次训练该动作所有正式组轻松完成目标次数上限","下次加重量：哑铃+1-2kg/边，器械+1片"),
    @("➡️ 维持不动","能完成目标次数但最后1-2组比较吃力","维持当前重量继续打磨——这恰恰是最佳训练区间"),
    @("⬇️ 需要减重","连续2次训练无法完成目标次数下限，或动作变形","减重5-10%，重建动作质量——退一步是为了进两步"),
    @("🔄 该换动作了","某个动作连续4周没有进步，或训练时关节持续不适","换同肌群的替代动作"),
    @("🦵 膝盖在分腿蹲时不适","训练中出现膝盖锐痛/针刺感","当天换成完全无膝压组合：臀推+RDL+站姿弹力带后踢腿"),
    @("😴 每周3练恢复不过来","连续2周训练后48小时+仍明显酸痛，精力下降","暂时降为每周2练（用精简备用方案），2周后加回来"),
    @("💪 体重没变但力量在涨","—","✅ 这是最好的信号！身体重组正在发生——脂肪换肌肉。继续！"),
    @("⚖️ 体重掉>0.5kg/周且力量也掉","—","⚠️ 热量缺口太大了。训练日增加100-150kcal碳水。")
)

foreach ($ru in $rules) {
    Set-Cell $ws5 $r 1 $ru[0] $true 10 0 0 $false
    Set-Cell $ws5 $r 2 $ru[1] $false 10 0 0 $true
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $ru[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 5 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 38
    Set-Borders $ws5 $r 1 7
    $r++
}

$ws5.Range("A:A").ColumnWidth = 18
$ws5.Range("B:B").ColumnWidth = 26
$ws5.Range("C:C").ColumnWidth = 14
$ws5.Range("D:D").ColumnWidth = 14
$ws5.Range("E:E").ColumnWidth = 24
$ws5.Range("F:F").ColumnWidth = 14
$ws5.Range("G:G").ColumnWidth = 14

Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 精简备用+自主计划+复盘
# ============================================
$ws6 = $wb.Worksheets.Add()
$ws6.Name = "备用方案与复盘"

Set-Cell $ws6 1 1 "精简备用方案 · 自主训练框架 · 4周复盘清单" $true 16 $headerBlue 0 $true
Merge-Cells $ws6 1 1 7
$ws6.Range("A1:G1").RowHeight = 34

# Simplified
$r = 3
Set-Cell $ws6 $r 1 "2练精简备用方案（应酬/出差挤占训练日时使用）" $true 14 $white $accentOrange $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $accentOrange
$r++

Set-Cell $ws6 $r 1 "当某周只能练2次时，用这个版本。精简原则：保留核心复合动作，压缩辅助动作，去掉单独核心训练。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws6 $r 1 7
$r++

Set-Cell $ws6 $r 1 "序号" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "精简第1练（推力+下肢混合）" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 2 4
Set-Cell $ws6 $r 5 "精简第2练（拉力+下肢混合）" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 5 7
$r++

$simple = @(
    @("①","哑铃卧推 4x8-10","高位下拉 4x8-10"),
    @("②","器械推肩 3x10-12","坐姿划船 3x10-12"),
    @("③","哑铃分腿蹲 3x10-12","哑铃RDL 3x10-12"),
    @("④","哑铃臀推 3x12-15","龙门架面拉 3x12-15"),
    @("⑤","绳索飞鸟+下压超级组","哑铃弯举+仰卧举腿超级组"),
    @("有氧","跑步机爬坡走 10-15min","跑步机爬坡走 10-15min"),
    @("总时长","约40-45min力量 + 10-15min有氧","约40-45min力量 + 10-15min有氧")
)

foreach ($si in $simple) {
    Set-Cell $ws6 $r 1 $si[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $si[1] $false 10 0 0 $false
    Merge-Cells $ws6 $r 2 4
    Set-Cell $ws6 $r 5 $si[2] $false 10 0 0 $false
    Merge-Cells $ws6 $r 5 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 28
    Set-Borders $ws6 $r 1 7
    $r++
}

# Self-training framework
$r += 2
Set-Cell $ws6 $r 1 "学会自己做计划 · 4步框架" $true 14 $white $accentGreen $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $accentGreen
$r += 2

$framework = @(
    @("第1步","定分化方式","根据每周能练几天","2练→全身x2  |  3练→全身x3  |  4练→上下肢分化  |  5-6练→推拉腿"),
    @("第2步","每次选动作","按优先级从高到低","①主项复合1-2个 → ②辅助变式1-2个 → ③孤立动作1-2个 → ④核心动作1个"),
    @("第3步","定组数次数","通用模板速查","主项：4组x6-10次  |  辅助：3组x10-12次  |  孤立：3组x12-15次"),
    @("第4步","每4周评估一次","用进阶判断规则","会看信号→知道什么情况加重量/减重量/换动作→下一个4周就有了方向")
)

Set-Cell $ws6 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "做什么" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "核心问题" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "关键信息" $true 10 $white $headerBlue $true
Merge-Cells $ws6 $r 4 7
$r++

foreach ($fw in $framework) {
    Set-Cell $ws6 $r 1 $fw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $fw[1] $true 11 $headerBlue 0 $false
    Set-Cell $ws6 $r 3 $fw[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $fw[3] $false 10 0 0 $true
    Merge-Cells $ws6 $r 4 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 30
    Set-Borders $ws6 $r 1 7
    $r++
}

# 4-week review
$r += 2
Set-Cell $ws6 $r 1 "4周后复盘清单（建议和教练一起过一遍）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r += 2

$reviews = @(
    @("📊","力量变化","卧推、下拉、分腿蹲、RDL的重量和次数，对比第1周基准数据——涨了多少？"),
    @("⚖️","体重+腰围","晨起空腹称重趋势（每周固定2-3次），腰带松了几个孔？"),
    @("🦵","膝盖恢复","是否已无不适？能否逐步恢复腿屈伸/腿弯举？分腿蹲深度有没有加深？"),
    @("🍺","应酬影响评估","哪些应对策略有效？哪些需要调整？应酬后第二天训练质量如何？"),
    @("🎯","下一周期方向","继续身体重组？还是偏向增肌？还是加大减脂力度？"),
    @("🧠","自主做计划能力","能不能看懂这套框架？有没有信心自己调整下个周期？学到了什么？")
)

foreach ($rv in $reviews) {
    Set-Cell $ws6 $r 1 $rv[0] $false 14 0 0 $false
    Set-Cell $ws6 $r 2 $rv[1] $true 12 $headerBlue 0 $false
    Set-Cell $ws6 $r 3 $rv[2] $false 11 0 0 $true
    Merge-Cells $ws6 $r 3 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 34
    $r++
}

$ws6.Range("A:A").ColumnWidth = 10
$ws6.Range("B:B").ColumnWidth = 22
$ws6.Range("C:C").ColumnWidth = 16
$ws6.Range("D:G").ColumnWidth = 14

Write-Host "Sheet 6 done"

# ============================================
# Save
# ============================================
$savePath = "D:\Codex\Joe何_训练方案_V2.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws2) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws3) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws4) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws5) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws6) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()

Write-Host "Saved to: $savePath"
Write-Host "DONE"
