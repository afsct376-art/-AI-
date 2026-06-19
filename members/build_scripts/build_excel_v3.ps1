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

Write-Host "Setup complete"

# ============================================
# SHEET 1: 方案概览
# ============================================
$ws = $wb.Worksheets.Item(1)
$ws.Name = "方案概览"

Set-Cell $ws 1 1 "知行AI教练 · 定制训练方案" $true 18 $headerBlue 0 $true
Merge-Cells $ws 1 1 7
$ws.Range("A1:G1").RowHeight = 40

Set-Cell $ws 2 1 "Joe  He  |  37岁  |  168cm / 80kg  |  减脂增肌 身体重组" $false 12 $gold $darkBg $true
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
    @("训练方式","全身训练 x 3次/周（每次不同主项侧重：A推力主导 / B拉力主导 / C综合补弱）"),
    @("本方案周期","4周为一个中周期（容量周→强度周→强度周→减载周）"),
    @("周期逻辑","容量先行建立动作储备→强度跟进突破力量→减载周让身体超量恢复"),
    @("训练地点","酒店/小区健身房（哑铃最重20kg + 3台固定器械 + 小龙门架）"),
    @("有氧方式","跑步机爬坡走（力量训练后进行；备选：酒店楼梯/户外快走）"),
    @("方案特点","周期化训练设计 · 容量-强度-减载 · 膝关节保护 · 应酬场景适配")
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
Set-Cell $ws $r 1 "4周中周期节奏" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 7
$r++

Set-Cell $ws $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws $r 2 "阶段" $true 10 $white $headerBlue $false
Set-Cell $ws $r 4 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws $r 6 "目标" $true 10 $white $headerBlue $false
$r++

$periodWeeks = @(
    @("第1周","容量周","增加组数和次数，重量适中（60-65%极限）","建立动作储备，累积训练容量，为强度周做准备"),
    @("第2周","强度周（上）","重量递增，次数减少，组数维持","突破力量平台，上肢+1-2kg，下肢+2.5-5kg"),
    @("第3周","强度周（下）","继续加重，接近极限重量（80-85%）","冲击本周期力量上限"),
    @("第4周","减载周","重量降至50-60%，组数和次数均减少","让身体超量恢复，神经-肌肉系统修复，为下一周期蓄力")
)

foreach ($pw in $periodWeeks) {
    Set-Cell $ws $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 2 $pw[1] $true 10 $white $darkBg $false
    Set-Cell $ws $r 4 $pw[2] $false 10 0 0 $true
    Merge-Cells $ws $r 4 5
    Set-Cell $ws $r 6 $pw[3] $false 10 0 0 $true
    $ws.Range("A$($r):G$($r)").RowHeight = 32
    Set-Borders $ws $r 1 7
    $r++
}

$r++
Set-Cell $ws $r 1 "全年周期框架（知道自己在哪条路上）" $true 14 $headerBlue $lightBlue $true
Merge-Cells $ws $r 1 7
$r++

Set-Cell $ws $r 1 "大周期 = 1年" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "围绕长期目标：减脂增肌身体重组 → 达到理想体重(约70-72kg) + 力量水平稳定提升" $false 10 0 0 $true
Merge-Cells $ws $r 2 7
$ws.Range("A$($r):G$($r)").RowHeight = 25
$r++

Set-Cell $ws $r 1 "中周期 = 1季度(3个月)" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "Month1=容量月(优先加组数次数)  Month2=强度月(优先加重量)  Month3=整合月(维持+突破)  每4周末尾设减载周" $false 10 0 0 $true
Merge-Cells $ws $r 2 7
$ws.Range("A$($r):G$($r)").RowHeight = 32
$r++

Set-Cell $ws $r 1 "小周期 = 1个月(4周)" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "容量周 强度周 强度周 减载周——这就是你现在手里这份方案的结构" $false 10 0 0 $true
Merge-Cells $ws $r 2 7
$ws.Range("A$($r):G$($r)").RowHeight = 25
$r++

Set-Cell $ws $r 1 "上下肢加重差异" $true 11 $headerBlue 0 $false
Set-Cell $ws $r 2 "上肢每次+1-2.5kg（精细控制，肩肘关节安全）  |  下肢每次+2.5-5kg（下肢力量潜力更大，能承受更大增幅）" $false 10 0 0 $true
Merge-Cells $ws $r 2 7
$ws.Range("A$($r):G$($r)").RowHeight = 28
$r++

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
    @("3","容量先行，强度跟进","先建立动作储备（容量周），再突破力量（强度周），最后让身体恢复（减载周）。进步不只是加重——加组、加次、缩短休息，都是进步。")
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
# SHEET 2: 第1周训练计划 (容量周)
# ============================================
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "第1周-容量周"

Set-Cell $ws2 1 1 "第1周  容量周  训练计划" $true 16 $headerBlue 0 $true
Merge-Cells $ws2 1 1 7
$ws2.Range("A1:G1").RowHeight = 34

Set-Cell $ws2 2 1 "容量周目标：增加组数和次数，重量使用60-65%极限。重点是累积训练容量——组数比平时多1组，次数取上限，为第2-3周强度周建立动作模式和体能储备。记录所有数据作为基准。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws2 2 1 7
$ws2.Range("A2:G2").RowHeight = 34

function Write-DayPlan($ws, $refRow, $title, $focus, $estTime, $actions) {
    Set-Cell $ws $refRow 1 $title $true 13 $white $darkBg $true
    Merge-Cells $ws $refRow 1 7
    $ws.Range("A$($refRow):G$($refRow)").RowHeight = 28
    $ws.Range("A$($refRow):G$($refRow)").Font.Color = $white
    $ws.Range("A$($refRow):G$($refRow)").Interior.Color = $darkBg

    $refRow++
    Set-Cell $ws $refRow 1 "重点肌群：$focus  |  力量部分预估：约$estTime 分钟  |  容量周特点：组数+1，次数取上限" $false 10 $headerBlue $lightGray $true
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

# --- 全身A (容量周: 组数+1, 次数取上限) ---
$actA = @(
    @("哑铃平板卧推","5组","10-12次","单边14-16kg（容量周降重加量）","75秒","【呼吸】下放吸气、推起呼气。【纠偏】手肘与躯干45-60度，不要外展90度——减少肩关节压力。【发力感】脚踩实地，肩胛收紧，推起时想象把哑铃'推弯'，胸肌向中间挤压。"),
    @("器械坐姿推肩","4组","12-15次","插片选12-15RM（比平时轻1-2片）","60秒","【呼吸】推起呼气、下放吸气。【纠偏】下背贴紧靠背，不要反弓借力。【发力感】推到顶端三角肌收紧但不锁死肘关节。容量周：高次数追求泵感。"),
    @("哑铃分腿蹲","4组","12-15次/侧","单边10-14kg（双手持哑铃）","75秒","【呼吸】下蹲吸气、起身呼气。【纠偏】躯干竖直、前脚全掌踩实、深度控制在膝盖无痛范围内。【发力感】前腿股四头主导发力，重心放脚跟。加组不加重的容量策略。"),
    @("哑铃臀推","4组","15-20次","单颗14-18kg放髋部","60秒","【呼吸】推起呼气、下放吸气。【纠偏】下巴微收，不要过度伸展腰椎。【发力感】顶峰臀大肌用力夹紧1-2秒。高次数臀推——追求代谢压力。"),
    @("绳索飞鸟+绳索下压 超级组","各4组","各15次","龙门架轻重量","飞鸟 下压 休息45秒","【飞鸟】手臂微屈，想象抱一棵大树，胸肌驱动而非手臂拉动。【下压】大臂固定身体两侧，压到底微微外旋手腕。容量周：超级组组数也+1。")
)
$r = Write-DayPlan $ws2 $r "全身A  推力主导" "胸大肌 / 三角肌前中束 / 肱三头肌 / 股四头肌 / 核心" "40-45" $actA

# --- 全身B (容量周) ---
$actB = @(
    @("高位下拉","5组","10-12次","插片选12-15RM","60秒","【呼吸】下拉呼气、还原吸气。【纠偏】启动时肩胛骨先下沉再用手臂拉——不要用手臂硬拽。【发力感】想象把杆拉向锁骨，肘向下后方驱动，顶峰背阔肌收紧。"),
    @("坐姿水平划船","4组","12-15次","插片选12-15RM","60秒","【呼吸】拉向身体呼气、还原吸气。【纠偏】身体不要过多后仰借力，拉到腹部而非胸口。【发力感】拉到顶端肩胛骨向中间夹紧，中背部收缩。"),
    @("哑铃罗马尼亚硬拉","4组","12-15次","单边14-18kg","75秒","【呼吸】下放吸气、拉起呼气。【纠偏】膝盖微屈不弯曲过多，哑铃贴大腿前侧下滑，全程直背。【发力感】下放时感受腘绳肌拉伸，臀向后推。"),
    @("龙门架面拉","4组","15次","轻重量","45秒","【呼吸】拉向面部呼气、还原吸气。【纠偏】拉向额头而非脖子——保护肩关节。【发力感】肩后束+上背发力，手肘向外打开。"),
    @("哑铃弯举+仰卧举腿 超级组","各4组","弯举12-15次 / 举腿力竭","弯举单边6-10kg","弯举 举腿 休息45秒","【弯举】大臂固定不摆动，顶峰微微外旋手腕。【举腿】用下腹带动骨盆上卷，不靠惯性摆动。")
)
$r = Write-DayPlan $ws2 $r "全身B  拉力主导" "背阔肌 / 菱形肌 / 三角肌后束 / 肱二头肌 / 腘绳肌 / 臀" "40-45" $actB

# --- 全身C (容量周, 6个动作) ---
$actC = @(
    @("哑铃上斜卧推","4组","12-15次","单边12-16kg","60秒","【呼吸】同平板卧推。【纠偏】凳角30-45度，太高会变成肩推。【发力感】上胸主导，下放至胸锁骨水平位置。"),
    @("高位下拉（反握/窄握变式）","4组","12-15次","比宽握轻1-2片","60秒","【变式目的】反握更偏向下背阔肌，不同角度刺激。【呼吸+纠偏】同宽握，但侧重感受下背阔肌的收缩。"),
    @("哑铃分腿蹲","4组","12-15次/侧","单边10-14kg","75秒","【要点同全身A】容量周特点：组数+1，次数取上限。深度在膝盖无痛范围内。"),
    @("单腿哑铃罗马尼亚硬拉","3组","10-12次/侧","单边8-12kg","60秒","【纠偏】支撑腿膝盖微屈，髋为轴后推。【发力感】腘绳肌+臀+踝关节稳定同时训练。单腿更考验控制力。"),
    @("绳索下压+绳索弯举 超级组","各4组","各15次","轻重量","下压 弯举 休息45秒","三头+二头同时充血，容量积累。如果当天时间紧，可削减此超级组。"),
    @("平板支撑","3组","40-50秒","自重","30秒","【纠偏】腰不要塌，骨盆微微后倾。【发力感】腹肌主动收紧，想象有人要打你肚子——全身绷紧。")
)
$r = Write-DayPlan $ws2 $r "全身C  综合补弱" "上胸 / 背阔肌不同角度 / 下肢容量 / 手臂 / 核心稳定性" "42-48" $actC

$ws2.Range("A:A").ColumnWidth = 6
$ws2.Range("B:B").ColumnWidth = 22
$ws2.Range("C:C").ColumnWidth = 8
$ws2.Range("D:D").ColumnWidth = 14
$ws2.Range("E:E").ColumnWidth = 22
$ws2.Range("F:F").ColumnWidth = 14
$ws2.Range("G:G").ColumnWidth = 58

Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 第2周 (强度周上)
# ============================================
$ws3 = $wb.Worksheets.Add()
$ws3.Name = "第2周-强度周上"

Set-Cell $ws3 1 1 "第2周  强度周（上） 训练计划" $true 16 $headerBlue 0 $true
Merge-Cells $ws3 1 1 7
$ws3.Range("A1:G1").RowHeight = 34

Set-Cell $ws3 2 1 "强度周目标：重量递增，次数减少，组数回归正常。第1周的容量积累为这一周做了准备——现在用更重的重量突破力量。上肢+1-2.5kg，下肢+2.5-5kg。组数从5组回到4组，把第1周积累的体能转化为力量。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws3 2 1 7
$ws3.Range("A2:G2").RowHeight = 34

$r = 4

# --- 全身A (强度周: 重量增加) ---
$actA2 = @(
    @("哑铃平板卧推","4组","8-10次","单边16-18kg（第1周基础上+1-2kg/边）","90秒","【呼吸】下放吸气、推起呼气。【纠偏】手肘与躯干45-60度。【发力感】推起时胸肌向中间挤压。强度周：重量优先，次数降回8-10。"),
    @("器械坐姿推肩","3组","10-12次","插片+1片（比第1周重）","75秒","【呼吸】推起呼气、下放吸气。【纠偏】下背贴紧靠背。【发力感】推到顶端三角肌收紧不锁死。【加重提示】上肢精细加重，+1片即可。"),
    @("哑铃分腿蹲","3组","8-10次/侧","单边14-18kg（第1周基础上+2.5-5kg/边）","90秒","【呼吸】下蹲吸气、起身呼气。【纠偏】躯干竖直、深度在膝盖无痛范围内。【发力感】前腿股四头主导。【加重提示】下肢可以加重幅度大一些。⚠️ 膝盖无痛才加。"),
    @("哑铃臀推","3组","10-12次","单颗18-20kg（可冲20kg上限）","90秒","【呼吸】推起呼气、下放吸气。【纠偏】下巴微收。【发力感】顶峰臀大肌夹紧保持2秒。强度周：臀推上大重量。"),
    @("绳索飞鸟+绳索下压 超级组","各3组","各10-12次","比第1周重1档","飞鸟 下压 休息60秒","强度周辅助动作加重不加重——用控制力和收缩感弥补。【飞鸟】胸肌驱动。【下压】大臂固定。")
)
$r = Write-DayPlan $ws3 $r "全身A  推力主导" "胸大肌 / 三角肌前中束 / 肱三头肌 / 股四头肌 / 核心" "38-42" $actA2

# --- 全身B (强度周) ---
$actB2 = @(
    @("高位下拉","4组","8-10次","插片+1-2片（比第1周重）","90秒","【呼吸】下拉呼气、还原吸气。【纠偏】先沉肩胛再用手臂拉。【发力感】肘向下后方驱动，顶峰背阔肌收紧。"),
    @("坐姿水平划船","3组","10-12次","插片+1片","75秒","【呼吸】拉向身体呼气、还原吸气。【纠偏】身体不要后仰借力。【发力感】拉到顶端肩胛骨向中间夹紧。"),
    @("哑铃罗马尼亚硬拉","3组","8-10次","单边18-20kg（第1周基础上+2.5-5kg/边）","90秒","【呼吸】下放吸气、拉起呼气。【纠偏】膝盖微屈，哑铃贴腿下滑，全程直背。【发力感】腘绳肌拉伸后爆发收紧。【加重提示】下肢RDL可以大胆加。"),
    @("龙门架面拉","3组","12-15次","轻重量","45秒","【呼吸】拉向面部呼气、还原吸气。【纠偏】拉向额头而非脖子。【发力感】肩后束+上背发力。面拉不求重量，求质量。"),
    @("哑铃弯举+仰卧举腿 超级组","各3组","弯举8-10次 / 举腿力竭","弯举单边10-14kg","弯举 举腿 休息60秒","【弯举】加重后严格控制，杜绝借力摆动。【举腿】下腹发力带动骨盆。")
)
$r = Write-DayPlan $ws3 $r "全身B  拉力主导" "背阔肌 / 菱形肌 / 三角肌后束 / 肱二头肌 / 腘绳肌 / 臀" "38-42" $actB2

# --- 全身C (强度周) ---
$actC2 = @(
    @("哑铃上斜卧推","3组","8-10次","单边16-18kg（+2-3kg/边）","75秒","【呼吸】同平板卧推。【纠偏】凳角30-45度。【发力感】上胸主导发力。强度周：上斜也比第1周重。"),
    @("高位下拉（反握/窄握）","3组","10-12次","比宽握轻1-2片但比第1周反握重","60秒","【要点】反握偏向下背阔肌。比第1周同动作加重。【发力感】下背阔肌收缩。"),
    @("哑铃分腿蹲","3组","8-10次/侧","单边14-18kg","75秒","【要点同全身A】强度周：减少次数，增加重量。膝盖无痛才加。"),
    @("单腿哑铃罗马尼亚硬拉","3组","8-10次/侧","单边12-16kg","75秒","单腿RDL在强度周也可以合理加重——单侧变式考验稳定性的同时也考验力量。"),
    @("绳索下压+绳索弯举 超级组","各3组","各10-12次","比第1周重1档","下压 弯举 休息60秒","强度周手臂辅助动作适度加重。如果当天时间紧，可削减此超级组。"),
    @("平板支撑","3组","30-45秒","自重","30秒","维持核心稳定性即可。")
)
$r = Write-DayPlan $ws3 $r "全身C  综合补弱" "上胸 / 背阔肌不同角度 / 下肢容量 / 手臂 / 核心稳定性" "38-42" $actC2

$ws3.Range("A:A").ColumnWidth = 6
$ws3.Range("B:B").ColumnWidth = 22
$ws3.Range("C:C").ColumnWidth = 8
$ws3.Range("D:D").ColumnWidth = 14
$ws3.Range("E:E").ColumnWidth = 22
$ws3.Range("F:F").ColumnWidth = 14
$ws3.Range("G:G").ColumnWidth = 58

Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 第3-4周 强度周下+减载周
# ============================================
$ws4 = $wb.Worksheets.Add()
$ws4.Name = "第3-4周-强度下-减载"

Set-Cell $ws4 1 1 "第3周强度周（下）  +  第4周减载周" $true 16 $headerBlue 0 $true
Merge-Cells $ws4 1 1 7
$ws4.Range("A1:G1").RowHeight = 34

# ----- 第3周 -----
$r = 3
Set-Cell $ws4 $r 1 "第3周  强度周（下）  冲击本周期力量上限" $true 14 $white $accentRed $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $accentRed
$r++

Set-Cell $ws4 $r 1 "第3周目标：继续加重至80-85%极限重量。这是本周期的高点——用第1-2周铺垫的体能和技术冲击力量上限。如果某些动作已经触及器械上限（哑铃20kg），切换到突破策略（慢离心/1.5次法/单侧变式）。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 30
$r++

Set-Cell $ws4 $r 1 "动作框架延续第2周，变化如下" $true 11 $headerBlue 0 $false
Merge-Cells $ws4 $r 1 7
$r++

Set-Cell $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "第2周重量" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 3 "第3周调整" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "组数x次数" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 5 "已达上限的备选策略" $true 10 $white $headerBlue $false
$r++

$week3Changes = @(
    @("哑铃平板卧推","16-18kg","+1-2kg/边","4x6-8","已达20kg→慢离心3-4秒"),
    @("器械坐姿推肩","+1片","再+1片","3x8-10","已达顶→缩短间歇45秒"),
    @("哑铃分腿蹲","14-18kg","+2.5-5kg/边","3x6-8","已达20kg→1.5次法或单腿蹲"),
    @("哑铃臀推","18-20kg","已近上限，换单腿臀推","3x8-10/侧","单腿臀推——轻重量双倍刺激"),
    @("高位下拉","+1-2片","再+1片","4x6-8","已达顶→慢离心4秒"),
    @("坐姿划船","+1片","再+1片或慢离心","3x8-10","已达顶→缩短休息"),
    @("哑铃RDL","18-20kg","+2.5-5kg或单腿变式","3x6-8","已达20kg→单腿RDL"),
    @("辅助动作","—","维持或微调","3x8-12","辅助动作不追求极限")
)

foreach ($c3 in $week3Changes) {
    Set-Cell $ws4 $r 1 $c3[0] $true 9 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $c3[1] $false 9 0 0 $false
    Set-Cell $ws4 $r 3 $c3[2] $false 9 0 0 $false
    Set-Cell $ws4 $r 4 $c3[3] $false 9 0 0 $false
    Set-Cell $ws4 $r 5 $c3[4] $false 9 0 0 $true
    $ws4.Range("A$($r):G$($r)").RowHeight = 30
    Set-Borders $ws4 $r 1 7
    $r++
}

# ----- 第4周 减载周 -----
$r += 2
Set-Cell $ws4 $r 1 "第4周  减载周  让身体超量恢复" $true 14 $white $accentGreen $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 26
$ws4.Range("A$($r):G$($r)").Font.Color = $white
$ws4.Range("A$($r):G$($r)").Interior.Color = $accentGreen
$r++

Set-Cell $ws4 $r 1 "减载周目标：主动恢复。重量降至50-60%极限，组数和次数均减少。让神经-肌肉系统、关节、内分泌系统全面修复。经过前3周的容量+强度刺激，这一周的'退'恰恰是为了下一个周期的'进'——超量恢复。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 30
$r++

Set-Cell $ws4 $r 1 "减载周不是完全休息——是'主动恢复'。完全不练会让身体脱敏，低强度低容量训练保持神经适应性，同时让身体修复。" $false 10 $accentOrange 0 $true
Merge-Cells $ws4 $r 1 7
$r++

Set-Cell $ws4 $r 1 "参数" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "减载周设置" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "说明" $true 10 $white $headerBlue $false
$r++

$deload = @(
    @("训练频率","仍为每周3练","不减训练日，减训练量——保持习惯节奏"),
    @("动作框架","沿用第1-3周的动作","每个动作都做，但全部降低强度"),
    @("重量","50-60%极限重量","大约第1周重量的70-80%。完全不冲。"),
    @("组数","每动作减1-2组","主项3组（原4-5组），辅助2组（原3-4组）"),
    @("次数","减少30-40%","主项6-8次（原8-12次），辅助8-10次（原12-15次）"),
    @("休息","延长至90-120秒","不赶时间，充分休息，专注动作质量和身体感受"),
    @("有氧","仍做10-15分钟","低强度——坡度8%代替10-12%，速度降到4km/h"),
    @("主观感受","RPE 5-6/10","练完应该感觉'没练够'——这正是减载周该有的感觉"),
    @("禁止事项","不冲重量、不力竭、不增加新动作","减载周不是'试试新纪录'的时候——让身体真正休息")
)

foreach ($d in $deload) {
    Set-Cell $ws4 $r 1 $d[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $d[1] $false 10 0 0 $true
    Merge-Cells $ws4 $r 2 3
    Set-Cell $ws4 $r 4 $d[2] $false 10 0 0 $true
    Merge-Cells $ws4 $r 4 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 30
    Set-Borders $ws4 $r 1 7
    $r++
}

$r++
Set-Cell $ws4 $r 1 "减载周训练示意（全身A为例，B/C同理缩放）" $true 11 $headerBlue 0 $false
Merge-Cells $ws4 $r 1 7
$r++

Set-Cell $ws4 $r 1 "动作" $true 9 $white $headerBlue $false
Set-Cell $ws4 $r 2 "原组数/重量" $true 9 $white $headerBlue $false
Set-Cell $ws4 $r 3 "减载组数/重量" $true 9 $white $headerBlue $false
Set-Cell $ws4 $r 5 "感受" $true 9 $white $headerBlue $false
$r++

$deloadExample = @(
    @("哑铃平板卧推","5组x10-12次 14-16kg","3组x8次 10-12kg","轻松完成，感受胸肌收缩节奏"),
    @("器械坐姿推肩","4组x12-15次","2组x10次 减2-3片","肩关节无压力，活动度优先"),
    @("哑铃分腿蹲","4组x12-15次/侧 10-14kg","2组x8次/侧 6-8kg","膝盖感受放在第一位"),
    @("哑铃臀推","4组x15-20次","2组x10次 10kg","臀肌激活即可"),
    @("绳索超级组","各4组x15次","各2组x10次 更轻","让关节滑液充分润滑")
)

foreach ($de in $deloadExample) {
    Set-Cell $ws4 $r 1 $de[0] $true 9 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $de[1] $false 9 0 0 $false
    Set-Cell $ws4 $r 3 $de[2] $false 9 0 0 $false
    Set-Cell $ws4 $r 5 $de[3] $false 9 0 0 $true
    Merge-Cells $ws4 $r 5 7
    $ws4.Range("A$($r):G$($r)").RowHeight = 26
    Set-Borders $ws4 $r 1 7
    $r++
}

$r++
Set-Cell $ws4 $r 1 " 减载周结束时，你应该感觉：精力充沛、关节轻松、'想练但没练够'——这就是超量恢复即将发生的最佳状态。第5周（即下个周期的容量周）你将带着更强的身体储备重新出发。" $false 10 $accentGreen 0 $true
Merge-Cells $ws4 $r 1 7
$ws4.Range("A$($r):G$($r)").RowHeight = 30

$ws4.Range("A:A").ColumnWidth = 22
$ws4.Range("B:B").ColumnWidth = 22
$ws4.Range("C:C").ColumnWidth = 20
$ws4.Range("D:G").ColumnWidth = 16

Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 有氧+热身+冷身+建议购置
# ============================================
$ws5 = $wb.Worksheets.Add()
$ws5.Name = "有氧热身冷身"

Set-Cell $ws5 1 1 "有氧安排  热身  冷身  建议购置" $true 16 $headerBlue 0 $true
Merge-Cells $ws5 1 1 7
$ws5.Range("A1:G1").RowHeight = 34

$r = 3
Set-Cell $ws5 $r 1 "有氧安排（力量训练后进行）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "时长" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "强度" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 5 "说明" $true 10 $white $headerBlue $false
$r++

$cardioPeriod = @(
    @("第1周 容量周","12-15分钟","坡度10% 速度4.5-5km/h","力量训练容量大——有氧也加时，增加总消耗"),
    @("第2周 强度周","12-15分钟","坡度10% 速度4.5-5km/h","维持有氧——强度周不下调有氧"),
    @("第3周 强度周","10-12分钟","坡度12% 速度5-5.5km/h","强度周下——适当缩短时间但提高坡度"),
    @("第4周 减载周","10分钟","坡度8% 速度4km/h","跟着力量一起减载——轻松走，不求消耗，求恢复")
)

foreach ($cp in $cardioPeriod) {
    Set-Cell $ws5 $r 1 $cp[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $cp[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 3 $cp[2] $false 10 0 0 $false
    Set-Cell $ws5 $r 5 $cp[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 5 7
    Set-Borders $ws5 $r 1 7
    $r++
}

$r++
Set-Cell $ws5 $r 1 "方式：跑步机爬坡走（首选）| 备选：酒店楼梯/户外快走 | 心率：120-140bpm（能断续说话不能流畅聊天）| 频率：每周3次（力量后）" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 24

# Warmup
$r += 2
Set-Cell $ws5 $r 1 "训练前热身（8分钟  必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "动作" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "时间/次数" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 6 "目的" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 6 7
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
    Set-Cell $ws5 $r 1 $w[0] $false 10 0 0 $false
    Set-Cell $ws5 $r 2 $w[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $w[2] $false 10 0 0 $false
    Set-Cell $ws5 $r 6 $w[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 6 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 26
    Set-Borders $ws5 $r 1 7
    $r++
}

# Cooldown
$r += 2
Set-Cell $ws5 $r 1 "训练后冷身（5-8分钟  必须做）" $true 14 $white $darkBg $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws5 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "动作" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 2 4
Set-Cell $ws5 $r 5 "时间" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 6 "针对" $true 10 $white $headerBlue $false
Merge-Cells $ws5 $r 6 7
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
    Set-Cell $ws5 $r 1 $c[0] $false 10 0 0 $false
    Set-Cell $ws5 $r 2 $c[1] $true 10 $headerBlue 0 $false
    Merge-Cells $ws5 $r 2 4
    Set-Cell $ws5 $r 5 $c[2] $false 10 0 0 $false
    Set-Cell $ws5 $r 6 $c[3] $false 10 0 0 $true
    Merge-Cells $ws5 $r 6 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 26
    Set-Borders $ws5 $r 1 7
    $r++
}

# Purchases
$r += 2
Set-Cell $ws5 $r 1 "建议购置：泡沫轴 + 弹力带（总花费约50-80元）" $true 14 $white $accentGreen $true
Merge-Cells $ws5 $r 1 7
$ws5.Range("A$($r):G$($r)").RowHeight = 26
$ws5.Range("A$($r):G$($r)").Font.Color = $white
$ws5.Range("A$($r):G$($r)").Interior.Color = $accentGreen
$r++

Set-Cell $ws5 $r 1 "物品" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "价格" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 3 "用途与使用方法" $true 10 $white $headerBlue $true
Merge-Cells $ws5 $r 3 7
$r++

$purchases = @(
    @("泡沫轴（光滑面/中等硬度）","30-50元","【训练前】滚股四头肌、髂胫束、腘绳肌各30秒——尤其下肢日，大腿前侧/外侧过紧常是膝痛的根源。【训练后】滚胸椎、上背各30秒。酸痛点停住深呼吸10-15秒，让肌肉松开。"),
    @("弹力带（15-25磅）","15-30元","【热身时】肩胛激活15次（双手拉开弹力带）、臀桥+膝盖外撑15次（弹力带绑膝上）。【拉伸时】仰卧腘绳肌拉伸（弹力带套脚底辅助）、或辅助引体向上。"),
    @("放进行李箱不占地方","总约50-80元","酒店和家里都能用。对恢复和关节健康的价值远超这个价格——几十块解决大问题。")
)

foreach ($p in $purchases) {
    Set-Cell $ws5 $r 1 $p[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $p[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 3 $p[2] $false 10 0 0 $true
    Merge-Cells $ws5 $r 3 7
    $ws5.Range("A$($r):G$($r)").RowHeight = 48
    Set-Borders $ws5 $r 1 7
    $r++
}

$ws5.Range("A:A").ColumnWidth = 10
$ws5.Range("B:B").ColumnWidth = 24
$ws5.Range("C:C").ColumnWidth = 16
$ws5.Range("D:D").ColumnWidth = 16
$ws5.Range("E:E").ColumnWidth = 16
$ws5.Range("F:F").ColumnWidth = 16
$ws5.Range("G:G").ColumnWidth = 16

Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 膝关节保护+安全+营养+规则+器械突破
# ============================================
$ws6 = $wb.Worksheets.Add()
$ws6.Name = "保护饮食规则突破"

Set-Cell $ws6 1 1 "膝关节保护  安全提示  器械突破  营养建议  进阶判断规则" $true 16 $headerBlue 0 $true
Merge-Cells $ws6 1 1 7
$ws6.Range("A1:G1").RowHeight = 34

# Knee
$r = 3
Set-Cell $ws6 $r 1 "膝关节保护专项   重要！" $true 14 $white $accentRed $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $accentRed
$r++

Set-Cell $ws6 $r 1 "级别" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "动作" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "原因" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 3 6
Set-Cell $ws6 $r 7 "怎么办" $true 10 $white $headerBlue $false
$r++

$knee = @(
    @("  禁止","坐姿腿屈伸机","膝关节前侧剪切力——你上次的不适就是相关训练后出现的","4周内完全不碰（含减载周）"),
    @("  禁止","坐姿腿弯举机","上次练完这个后膝盖出现不适","4周内完全不碰（含减载周）"),
    @("  谨慎","哑铃分腿蹲","从浅深度开始，每周试探性加深1-2cm，无痛才继续","哪个训练日膝盖感觉不好→当天换臀推+RDL组合"),
    @("  安全","臀推/臀桥","膝关节几乎不参与","放心练，所有周期都可以加重"),
    @("  安全","罗马尼亚硬拉","膝关节几乎不动（只微屈），压力在髋关节","腘绳肌+臀部的核心动作，放心练"),
    @("  康复","靠墙静蹲（每次热身前）","温和激活股四头肌，零膝关节压力","每次训练前30秒x2组"),
    @("  停止信号","膝盖出现锐痛/针刺感/肿胀","—","当天不再做任何下肢动作。酸痛/紧绷感→正常，可继续但要降强度")
)

foreach ($k in $knee) {
    Set-Cell $ws6 $r 1 $k[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $k[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws6 $r 3 $k[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 3 6
    Set-Cell $ws6 $r 7 $k[3] $false 10 0 0 $true
    $ws6.Range("A$($r):G$($r)").RowHeight = 38
    Set-Borders $ws6 $r 1 7
    $r++
}

# Safety
$r += 2
Set-Cell $ws6 $r 1 "安全提示" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "  你的优势" $true 10 $headerBlue $lightGray $false
Set-Cell $ws6 $r 2 "无伤病史（膝盖不适是近期可逆的） | 1.5年私教基础，动作模式已有基础 | 37岁尚有较好的恢复能力和睾酮水平" $false 10 0 0 $true
Merge-Cells $ws6 $r 2 7
Set-Borders $ws6 $r 1 7
$r++

Set-Cell $ws6 $r 1 "  需要注意" $true 10 $accentOrange $lightGray $false
Set-Cell $ws6 $r 2 "37岁恢复速度不等于20岁——保证7-8小时睡眠 | 体重80kg/BMI 28.3——跳跃类动作不加入，保护膝关节 | 应酬饮酒抑制蛋白质合成——应酬日前后训练强度下调 | 酒店健身房无人保护——哑铃最后一组不做到力竭 | 减载周必须认真执行——37岁的恢复能力尤其需要周期性的减载" $false 10 0 0 $true
Merge-Cells $ws6 $r 2 7
$ws6.Range("A$($r):G$($r)").RowHeight = 48
Set-Borders $ws6 $r 1 7
$r++

Set-Cell $ws6 $r 1 "  建议避免" $true 10 $accentRed $lightGray $false
Set-Cell $ws6 $r 2 "颈后推举/颈后下拉（肩关节不必要风险） | 坐姿腿屈伸/腿弯举（膝盖恢复前） | 空腹力量训练 | 哑铃模拟大重量杠铃深蹲（没深蹲架导致下背风险大） | 减载周冲重量（身体在修复期，不是突破期）" $false 10 0 0 $true
Merge-Cells $ws6 $r 2 7
$ws6.Range("A$($r):G$($r)").RowHeight = 42
Set-Borders $ws6 $r 1 7
$r++

# Equipment breakthrough strategies
$r += 2
Set-Cell $ws6 $r 1 "器械上限突破策略（哑铃20kg/器械插片到顶时）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "策略" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "操作" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "适用周期" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "适用动作" $true 10 $white $headerBlue $false
$r++

$btStrategies = @(
    @("慢离心训练","3-4秒下放+爆发推起","所有周期均可使用","卧推、下拉、分腿蹲"),
    @("1.5次法","下放→起一半→再下放→完全起身=1次","容量周和强度周","分腿蹲、臀推——代谢压力翻倍"),
    @("缩短组间休息","从90秒逐步缩到45秒","容量周优先使用","所有动作——增加训练密度"),
    @("增加次数","从8-10次进阶到15-20次","容量周——核心策略","所有动作——容量来源"),
    @("单侧变式","单腿RDL、单臂划船、单腿臀推","强度周首选","下肢、背部——轻重量高刺激"),
    @("弹力带叠加","哑铃+弹力带双阻力","强度周突破","卧推、臀推——变相增加负重"),
    @("预疲劳法","先用孤立动作力竭再用复合动作","容量周进阶","飞鸟→卧推，面拉→划船")
)

foreach ($b in $btStrategies) {
    Set-Cell $ws6 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $b[1] $false 10 0 0 $true
    Set-Cell $ws6 $r 3 $b[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $b[3] $false 10 0 0 $true
    Merge-Cells $ws6 $r 4 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 32
    Set-Borders $ws6 $r 1 7
    $r++
}

# Nutrition
$r += 2
Set-Cell $ws6 $r 1 "营养建议（通用方向，非处方）——按周期调整" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "当前状况：80kg，体脂25-30%，BMI 28.3。目标减脂增肌。过往从85kg减到72kg说明你能减，但回升到80kg说明维持策略需要调整。本方案不搞激进减脂，做可持续的身体重组。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 28
$r++

Set-Cell $ws6 $r 1 "周期" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "总热量" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "蛋白质" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 4 "碳水" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 6 "说明" $true 10 $white $headerBlue $false
$r++

$nutriPeriod = @(
    @("容量周","2100-2300 kcal","144-160g","230-260g","训练量大→热量和碳水要跟上，保证恢复"),
    @("强度周","2100-2300 kcal","144-160g","230-260g","保持——强度周需要能量支撑"),
    @("减载周","1800-2000 kcal","144-160g","160-180g","训练量减少→热量和碳水相应减少，蛋白质不减"),
    @("（脂肪全天周统一50-55g，饮水3-4L）","","","","")
)

foreach ($np in $nutriPeriod) {
    Set-Cell $ws6 $r 1 $np[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $np[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 3 $np[2] $false 10 0 0 $false
    Set-Cell $ws6 $r 4 $np[3] $false 10 0 0 $false
    Set-Cell $ws6 $r 6 $np[4] $false 10 0 0 $true
    Merge-Cells $ws6 $r 6 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 24
    Set-Borders $ws6 $r 1 7
    $r++
}

$r++
Set-Cell $ws6 $r 1 "三餐实操（生活化调整，不要求每天称克数）" $true 12 $headerBlue 0 $true
Merge-Cells $ws6 $r 1 7
$r++

Set-Cell $ws6 $r 1 "餐次" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "你现在的习惯" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 3 "可以这样调整" $true 10 $white $headerBlue $true
Merge-Cells $ws6 $r 3 7
$r++

$meals = @(
    @("早餐（居家）","麦片40g+牛奶250ml+1.5个蛋","  保持！换成2个整蛋，蛋白质量提到约30g。面包/馒头时优先选全麦。"),
    @("午餐（食堂围桌）","食堂炒菜","米饭控在一拳大小。夹菜次序：肉/鱼/豆腐→蔬菜→主食。油多的菜先沥一下再夹。吃到七分饱停筷。"),
    @("晚餐（居家）","跟家人一起吃","米饭减半或换杂粮。保证一份手掌大瘦肉/鱼。蔬菜不限量。家人做菜油多→自己盛出来前控油。"),
    @("练后加餐","一勺蛋白粉","  保持！加一根香蕉(便利店买)，补充肌糖原促进恢复。"),
    @("应酬（约每周2次）","—","当天午餐做减法。喝酒前吃蛋白质垫肚子。能不喝就不喝。应酬后多喝水。核心心态：80%执行率就是成功。")
)

foreach ($m in $meals) {
    Set-Cell $ws6 $r 1 $m[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws6 $r 2 $m[1] $false 10 0 0 $false
    Set-Cell $ws6 $r 3 $m[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 3 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 42
    Set-Borders $ws6 $r 1 7
    $r++
}

# Advance judgment rules
$r += 2
Set-Cell $ws6 $r 1 "进阶判断规则（教你自己做决策）" $true 14 $white $darkBg $true
Merge-Cells $ws6 $r 1 7
$ws6.Range("A$($r):G$($r)").RowHeight = 26
$ws6.Range("A$($r):G$($r)").Font.Color = $white
$ws6.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws6 $r 1 "信号" $true 10 $white $headerBlue $false
Set-Cell $ws6 $r 2 "判断条件" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 2 4
Set-Cell $ws6 $r 5 "做法" $true 10 $white $headerBlue $false
Merge-Cells $ws6 $r 5 7
$r++

$rules = @(
    @("  可以加重（强度周）","连续2次训练该动作所有正式组轻松完成目标次数上限","下次加重量：上肢+1-2.5kg/边，下肢+2.5-5kg/边"),
    @("  可以加组/加次（容量周）","动作质量稳定，恢复良好","下个容量周：主项+1组，或次数上限+2次"),
    @("  维持不动","能完成目标次数但最后1-2组比较吃力","维持当前重量继续打磨——这恰恰是最佳训练区间"),
    @("  需要减重","连续2次训练无法完成目标次数下限，或动作变形","减重5-10%，重建动作质量——退一步是为了进两步"),
    @("  该换动作了","某个动作连续4周没有进步，或训练时关节持续不适","换同肌群的替代动作"),
    @("  膝盖在分腿蹲时不适","训练中出现膝盖锐痛/针刺感","当天换成完全无膝压组合：臀推+RDL+站姿弹力带后踢腿"),
    @("  每周3练恢复不过来","连续2周训练后48小时+仍明显酸痛，精力下降","暂时降为每周2练（用精简备用方案），2周后加回来"),
    @("  减载周感觉'太轻松'","—","  这是对的！减载周就该轻松。不轻松说明你前三周没练到位。"),
    @("  体重没变但力量在涨","—","  这是最好的信号！身体重组正在发生——脂肪换肌肉。继续！"),
    @("  体重掉>0.5kg/周且力量也掉","—","  热量缺口太大了。训练日增加100-150kcal碳水。")
)

foreach ($ru in $rules) {
    Set-Cell $ws6 $r 1 $ru[0] $true 10 0 0 $false
    Set-Cell $ws6 $r 2 $ru[1] $false 10 0 0 $true
    Merge-Cells $ws6 $r 2 4
    Set-Cell $ws6 $r 5 $ru[2] $false 10 0 0 $true
    Merge-Cells $ws6 $r 5 7
    $ws6.Range("A$($r):G$($r)").RowHeight = 38
    Set-Borders $ws6 $r 1 7
    $r++
}

$ws6.Range("A:A").ColumnWidth = 18
$ws6.Range("B:B").ColumnWidth = 26
$ws6.Range("C:C").ColumnWidth = 14
$ws6.Range("D:D").ColumnWidth = 14
$ws6.Range("E:E").ColumnWidth = 24
$ws6.Range("F:F").ColumnWidth = 14
$ws6.Range("G:G").ColumnWidth = 14

Write-Host "Sheet 6 done"

# ============================================
# SHEET 7: 全年周期框架 + 备用方案 + 自主计划 + 复盘
# ============================================
$ws7 = $wb.Worksheets.Add()
$ws7.Name = "全年周期与复盘"

Set-Cell $ws7 1 1 "全年周期框架  精简备用方案  自主训练框架  4周复盘清单" $true 16 $headerBlue 0 $true
Merge-Cells $ws7 1 1 7
$ws7.Range("A1:G1").RowHeight = 34

# Yearly periodization
$r = 3
Set-Cell $ws7 $r 1 "全年大周期框架" $true 14 $white $darkBg $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 26
$ws7.Range("A$($r):G$($r)").Font.Color = $white
$ws7.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r++

Set-Cell $ws7 $r 1 "1个大周期 = 1年" $true 11 $headerBlue $lightGray $false
Set-Cell $ws7 $r 3 "围绕长期目标：减脂增肌身体重组  达到理想体重(约70-72kg) + 力量水平稳定提升" $false 10 0 0 $true
Merge-Cells $ws7 $r 3 7
Set-Borders $ws7 $r 1 7
$r++

Set-Cell $ws7 $r 1 "1个中周期 = 1季度(3个月)" $true 11 $headerBlue $lightGray $false
Set-Cell $ws7 $r 3 "月1=容量月(优先增加组数/次数)  月2=强度月(优先增加重量)  月3=整合月(维持+专项突破)  每4周末尾设减载周" $false 10 0 0 $true
Merge-Cells $ws7 $r 3 7
Set-Borders $ws7 $r 1 7
$r++

Set-Cell $ws7 $r 1 "1个小周期 = 1个月(4周)" $true 11 $headerBlue $lightGray $false
Set-Cell $ws7 $r 3 "容量周  强度周  强度周  减载周——你现在手里的方案就是这个结构" $false 10 0 0 $true
Merge-Cells $ws7 $r 3 7
Set-Borders $ws7 $r 1 7
$r++

$r++
Set-Cell $ws7 $r 1 "季度中周期示例（Q1：本季度）" $true 12 $headerBlue 0 $true
Merge-Cells $ws7 $r 1 7
$r++

Set-Cell $ws7 $r 1 "月份" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "阶段" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 3 "核心策略" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 5 "重点" $true 10 $white $headerBlue $false
$r++

$quarterly = @(
    @("第1个月","容量月","优先加组数(+1组)+加次数(+2-3次)，重量维持60-65%","建立动作储备和身体做功能力——为强度月蓄力"),
    @("第2个月","强度月","优先加重量（上肢+1-2.5kg，下肢+2.5-5kg），组数回归正常","突破力量平台——第1个月的容量转化为力量"),
    @("第3个月","整合月","维持强度月重量，适当加回一点容量（+1组）","整合容量+强度成果，月底评估→决定下季度方向"),
    @("（每个月内部都是4周：容量周→强度周→强度周→减载周）","","","")
)

foreach ($q in $quarterly) {
    Set-Cell $ws7 $r 1 $q[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $q[1] $true 10 $white $darkBg $false
    Set-Cell $ws7 $r 3 $q[2] $false 10 0 0 $true
    Set-Cell $ws7 $r 5 $q[3] $false 10 0 0 $true
    Merge-Cells $ws7 $r 5 7
    $ws7.Range("A$($r):G$($r)").RowHeight = 28
    Set-Borders $ws7 $r 1 7
    $r++
}

$r++
Set-Cell $ws7 $r 1 "上下肢加重差异原则" $true 12 $headerBlue 0 $true
Merge-Cells $ws7 $r 1 7
$r++

Set-Cell $ws7 $r 1 "  上肢（卧推/推肩/划船等）：每次+1-2.5kg" $false 10 0 0 $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 22
$r++

Set-Cell $ws7 $r 1 "  原因：上肢关节精细（肩、肘），小幅度加重保护关节，持续累积" $false 10 0 0 $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 22
$r++

Set-Cell $ws7 $r 1 "  下肢（深蹲/硬拉/RDL/臀推等）：每次+2.5-5kg" $false 10 0 0 $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 22
$r++

Set-Cell $ws7 $r 1 "  原因：下肢肌群大、力量潜力大、关节更粗壮，能承受更大的增幅" $false 10 0 0 $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 22
$r++

# Simplified plan
$r += 2
Set-Cell $ws7 $r 1 "2练精简备用方案（应酬/出差挤占训练日时使用）" $true 14 $white $accentOrange $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 26
$ws7.Range("A$($r):G$($r)").Font.Color = $white
$ws7.Range("A$($r):G$($r)").Interior.Color = $accentOrange
$r++

Set-Cell $ws7 $r 1 "当某周只能练2次时，用这个版本。精简原则：保留核心复合动作，压缩辅助动作，去掉单独核心训练。" $false 10 $headerBlue $lightGray $true
Merge-Cells $ws7 $r 1 7
$r++

Set-Cell $ws7 $r 1 "序号" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "精简第1练（推力+下肢混合）" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 2 4
Set-Cell $ws7 $r 5 "精简第2练（拉力+下肢混合）" $true 10 $white $headerBlue $false
Merge-Cells $ws7 $r 5 7
$r++

$simple = @(
    @("  ","哑铃卧推  根据所在周调整组数重量","高位下拉  根据所在周调整组数重量"),
    @("  ","器械推肩","坐姿划船"),
    @("  ","哑铃分腿蹲","哑铃RDL"),
    @("  ","哑铃臀推","龙门架面拉"),
    @("  ","绳索飞鸟+下压超级组","哑铃弯举+仰卧举腿超级组"),
    @("有氧","跑步机爬坡走（按所在周时长）","跑步机爬坡走（按所在周时长）"),
    @("总时长","约40-45min力量 + 10-15min有氧","约40-45min力量 + 10-15min有氧")
)

foreach ($si in $simple) {
    Set-Cell $ws7 $r 1 $si[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $si[1] $false 10 0 0 $false
    Merge-Cells $ws7 $r 2 4
    Set-Cell $ws7 $r 5 $si[2] $false 10 0 0 $false
    Merge-Cells $ws7 $r 5 7
    $ws7.Range("A$($r):G$($r)").RowHeight = 26
    Set-Borders $ws7 $r 1 7
    $r++
}

# Self-training framework
$r += 2
Set-Cell $ws7 $r 1 "学会自己做周期计划  4步框架" $true 14 $white $accentGreen $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 26
$ws7.Range("A$($r):G$($r)").Font.Color = $white
$ws7.Range("A$($r):G$($r)").Interior.Color = $accentGreen
$r += 2

$framework = @(
    @("第1步","定分化方式","根据每周能练几天","2练→全身x2  |  3练→全身x3  |  4练→上下肢分化  |  5-6练→推拉腿"),
    @("第2步","定周期节奏","4周一个循环","容量周→强度周→强度周→减载周。容量月优先加组加次，强度月优先加重。"),
    @("第3步","每次选动作","按优先级从高到低"," 主项复合1-2个 →  辅助变式1-2个 →  孤立动作1-2个 →  核心动作1个"),
    @("第4步","每4周评估一次","用进阶判断规则","会看信号→知道该加容量还是加重量→下一个4周就有了明确方向")
)

Set-Cell $ws7 $r 1 "步骤" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 2 "做什么" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 3 "核心问题" $true 10 $white $headerBlue $false
Set-Cell $ws7 $r 4 "关键信息" $true 10 $white $headerBlue $true
Merge-Cells $ws7 $r 4 7
$r++

foreach ($fw in $framework) {
    Set-Cell $ws7 $r 1 $fw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws7 $r 2 $fw[1] $true 11 $headerBlue 0 $false
    Set-Cell $ws7 $r 3 $fw[2] $false 10 0 0 $false
    Set-Cell $ws7 $r 4 $fw[3] $false 10 0 0 $true
    Merge-Cells $ws7 $r 4 7
    $ws7.Range("A$($r):G$($r)").RowHeight = 30
    Set-Borders $ws7 $r 1 7
    $r++
}

# 4-week review
$r += 2
Set-Cell $ws7 $r 1 "4周后复盘清单（建议和教练一起过一遍）" $true 14 $white $darkBg $true
Merge-Cells $ws7 $r 1 7
$ws7.Range("A$($r):G$($r)").RowHeight = 26
$ws7.Range("A$($r):G$($r)").Font.Color = $white
$ws7.Range("A$($r):G$($r)").Interior.Color = $darkBg
$r += 2

$reviews = @(
    @("  ","力量变化","容量周基准 vs 强度周上限——容量涨了多少？重量涨了多少？"),
    @("  ","体重+腰围","晨起空腹称重趋势（每周固定2-3次），腰带松了几个孔？"),
    @("  ","膝盖恢复","是否已无不适？能否逐步恢复腿屈伸/腿弯举？分腿蹲深度有没有加深？"),
    @("  ","减载周感受","减载周结束是否感觉精力充沛、'想练但没练够'？——这是超量恢复的信号。"),
    @("  ","应酬影响评估","哪些应对策略有效？哪些需要调整？应酬后第二天训练质量如何？"),
    @("  ","下一周期方向","继续容量月？还是进入强度月？身体重组趋势如何？"),
    @("  ","自主做计划能力","能不能看懂这套周期框架？有没有信心自己调整下个4周？学到了什么？")
)

foreach ($rv in $reviews) {
    Set-Cell $ws7 $r 1 $rv[0] $false 14 0 0 $false
    Set-Cell $ws7 $r 2 $rv[1] $true 12 $headerBlue 0 $false
    Set-Cell $ws7 $r 3 $rv[2] $false 11 0 0 $true
    Merge-Cells $ws7 $r 3 7
    $ws7.Range("A$($r):G$($r)").RowHeight = 32
    $r++
}

$ws7.Range("A:A").ColumnWidth = 10
$ws7.Range("B:B").ColumnWidth = 22
$ws7.Range("C:C").ColumnWidth = 16
$ws7.Range("D:G").ColumnWidth = 14

Write-Host "Sheet 7 done"

# ============================================
# Save
# ============================================
$savePath = "D:\Codex\Joe何_训练方案_V3.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws2) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws3) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws4) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws5) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws6) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws7) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()

Write-Host "Saved to: $savePath"
Write-Host "DONE"
