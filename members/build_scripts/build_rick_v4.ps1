$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $wb = $excel.Workbooks.Add()

$darkBg=0x1a1a2e;$gold=0xd4a574;$white=0xffffff;$lightGray=0xf5f5f5
$headerBlue=0x2c3e50;$accentGreen=0x27ae60;$accentRed=0xe74c3c
$accentOrange=0xf39c12;$lightBlue=0xd6eaf8;$softPink=0xe91e63
$warmPink=0xff6b81;$softPurple=0x8e44ad;$softBlue=0x2980b9

function SetC($ws,$r,$c,$v,$b,$s,$cl,$bg,$wr){
    $cell=$ws.Cells.Item($r,$c);$cell.Value="$v";$cell.Font.Bold=$b
    $cell.Font.Size=$s;if($cl){$cell.Font.Color=$cl}
    if($bg){$cell.Interior.Color=$bg};$cell.WrapText=$wr
}
function MC($ws,$r,$c1,$c2){$ws.Range($ws.Cells.Item($r,$c1),$ws.Cells.Item($r,$c2)).Merge()}
function SB($ws,$r,$c1,$c2){$rng=$ws.Range($ws.Cells.Item($r,$c1),$ws.Cells.Item($r,$c2));$rng.Borders.LineStyle=1;$rng.Borders.Weight=2}

Write-Host "Setup done"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "瑞克  力型兼备方案 V4" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  168cm / 67kg  |  力型兼备  高翻高频  午间计时  V4=国家力量教练逻辑" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4

# 术语速查
SetC $ws0 $row 1 "  专业术语速查" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

SetC $ws0 $row 1 "术语" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "全称" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "白话解释" $true 10 $white $headerBlue $false; MC $ws0 $row 3 5; $row++

$terms = New-Object 'System.Collections.Generic.List[Object]'
$terms.Add(@("DUP","每日波动周期","每天练的不一样。周一重、周三重+技能、周五快+容量。不让身体适应同一种节奏。"))
$terms.Add(@("AMRAP","尽可能多做一次","最后一组不限次数 做到动作快变形了就停。用来测试你今天真实的极限。"))
$terms.Add(@("RPE","自感用力程度","做完一组凭感觉打分。1=躺着 10=拼老命。8-9=做完还剩1-2次力竭。"))
$terms.Add(@("高翻","Power Clean","举重动作——杠铃从地面翻到锁骨前。最难举重技术之一。50%指高翻1RM的50% 没测过=流畅5次不砸锁骨的重量x0.5。"))
$terms.Add(@("三倍伸展","Triple Extension","高翻的核心发力——髋、膝、踝三个关节同时伸展。不是用手臂拉 是用腿和髋爆发。"))
$terms.Add(@("超级组","Superset","两个动作连着做 中间不休息。A做完直接做B——压缩时间 增加代谢压力。"))
$terms.Add(@("速度日","Dynamic Effort","轻重量爆发力训练(50-65%1RM)。练神经驱动 不累积疲劳。速度变慢=第3次明显比第1次慢 或需要硬挤。减慢减重5%。"))

foreach($t in $terms){
    SetC $ws0 $row 1 $t[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $t[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $t[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  V4 升级了什么——从好计划到国家级力量教练逻辑" $true 16 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades = New-Object 'System.Collections.Generic.List[Object]'
$upgrades.Add(@("  高翻 每周1次 每次训练前轻技术","放在每次训练热身之后 主项之前。5-8分钟 5x3@50-60%——高频低强度练神经适应。不影响后面主项。"))
$upgrades.Add(@("  周一卧推 4x5 3x5+1xAMRAP","前3组保证质量 第4组探底。AMRAP次数=每周进步的标尺。更安全 更有效。"))
$upgrades.Add(@("  周五 纯容量 速度日+容量混合","主项速度卧推8x3@60-65%爆发力——练神经驱动 不累积疲劳。辅助做容量——补同一肌群。"))
$upgrades.Add(@("  深蹲 固定 4周轮换变式","第1-4周=高杠 第5-8周=低杠 第9-12周=前蹲 第13-16周=箱式。每4周换一次 避免适应性停滞。"))
$upgrades.Add(@("  午间策略 加精确组间计时","主项120s 辅助90s 孤立45-60s 超级组60s。手机计时器严格到秒——45min内完成所有内容。"))
$upgrades.Add(@("  减载周 换动作 不是做轻一点","同样动作做轻量=关节应力路径相同。换动作=给关节不同应力分布 真正的休息。"))
$upgrades.Add(@("  营养 加力型兼备热量决策标准","体重每周涨0.25-0.5kg=热量合适。涨超0.5kg=减200kcal。不动=加200kcal。力量连续2周不涨=查热量和睡眠。"))
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  DUP 3天分化——V4版 每次训练前都有高翻技术练习" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "V4核心方法" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "RPE" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$dup = New-Object 'System.Collections.Generic.List[Object]'
$dup.Add(@("周一","上体力(大重量)","卧推3x5+1AMRAP 引体 推举 划船","8-9","前3组标准 第4组探底——重而可控"))
$dup.Add(@("周三","下肢力+高翻(技能+力量)","高翻5x3技术+深蹲4x5+RDL+核心","7-8","先磨技术 再冲力量——高翻在精力最好时"))
$dup.Add(@("周五","全身速度+容量","速度卧推8x3@60-65%+上斜+下拉+前蹲+手臂","7-8","爆发力推 轻快有力——不累积疲劳"))
foreach($d in $dup){
    SetC $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $d[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $d[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $d[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周周期——积累 强化 峰值 减载" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "上体力(周一)" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "下肢力(周三)" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "速度+容量(周五)" $true 10 $white $headerBlue $false; $row++

$o4 = New-Object 'System.Collections.Generic.List[Object]'
$o4.Add(@("1","积累","85% 3x5+AMRAP 建基准","高翻50%技术 深蹲80%","速度60% 建基准 容量70%"))
$o4.Add(@("2","强化","+2.5-5kg AMRAP超第1周","高翻55% 深蹲+5kg","速度65% 容量缩短休息或+1组"))
$o4.Add(@("3","峰值","再+2.5-5kg 3x3+AMRAP","高翻60% 深蹲87.5% 4x3","速度65%挑战更快速 容量再缩短休息"))
$o4.Add(@("4","减载","60% 3x5 换动作","高翻50%练速度 深蹲60%","50%轻量 换俯卧撑/高脚杯/弹力带"))
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  午间训练——V4精确组间计时 45分钟必须完成" $true 14 $white $accentOrange $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentOrange
$row++

SetC $ws0 $row 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "时间" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "内容" $true 10 $white $headerBlue $false; MC $ws0 $row 3 5; $row++

$timer = New-Object 'System.Collections.Generic.List[Object]'
$timer.Add(@("热身(含高翻技术)","10min","划船机3min+动态拉伸2min+神经激活1min+高翻技术5x3@50-60% 4min"))
$timer.Add(@("主项","12min","卧推/深蹲 4组x120s=8min+前面热身组4min"))
$timer.Add(@("辅助项","13min","3个辅助动作 各3-4组x90s 加切换时间"))
$timer.Add(@("孤立项+拉伸","10min","手臂超级组/侧平举/核心+拉伸5min"))
$timer.Add(@("总时长","45min+拉伸5min=50min","午休一小时内完成 下午上班不迟到"))
foreach($tm in $timer){
    SetC $ws0 $row 1 $tm[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $tm[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $tm[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx = New-Object 'System.Collections.Generic.List[Object]'
$idx.Add(@("完整方案逻辑+DUP框架+V4升级详解"," Sheet 方案概览"))
$idx.Add(@("上体力 下肢力+高翻 速度+容量"," Sheet 训练计划"))
$idx.Add(@("渐进决策树+高翻进阶+深蹲变式轮换"," Sheet 渐进决策树"))
$idx.Add(@("午间热身+主动恢复+减载换动作"," Sheet 恢复策略"))
$idx.Add(@("力型兼备营养+热量决策+判断规则"," Sheet 饮食与规则"))
$idx.Add(@("第一周示例+每周检查清单+FAQ"," Sheet 示例检查FAQ"))
$idx.Add(@("精简版+力型兼备复盘+自主调节"," Sheet 备用与复盘"))
foreach($ix in $idx){
    SetC $ws0 $row 1 $ix[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $ix[1] $false 10 0 0 $false; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; SB $ws0 $row 1 5; $row++
}

$ws0.Range("A:A").ColumnWidth=28; $ws0.Range("B:B").ColumnWidth=20; $ws0.Range("C:E").ColumnWidth=20
Write-Host "Cover done"

# ============================================
# SHEET 1: 方案概览
# ============================================
$ws1=$wb.Worksheets.Add();$ws1.Name="方案概览"

SetC $ws1 1 1 "瑞克  力型兼备  方案概览 V4" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "男 31-35岁 168cm/67kg 体脂20-25% 3-4年经验 力型兼备 高翻高频 午间计时 V4" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览 V4" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ovl = New-Object 'System.Collections.Generic.List[Object]'
$ovl.Add(@("训练频率","3次/周 DUP波动 每次含高翻技术练习(5-8min)","上体力(重)+下肢力+高翻(技能)+速度+容量(爆发+增肌)"))
$ovl.Add(@("核心升级","高翻每练必做 卧推AMRAP 速度日 深蹲变式轮换 精确计时 减载换动作 热量决策","7大升级从好计划到国家力量教练逻辑"))
$ovl.Add(@("周期结构","4周一循环(积累 强化 峰值 减载) 深蹲每4周轮换变式","每3个月全休1周"))
$ovl.Add(@("午间策略","练前轻食+高翻技术10min+主项12min+辅助13min+孤立拉伸10min=45+5min","手机计时器严格到秒"))
$ovl.Add(@("方案特点","力型兼备 DUP波动 高翻高频 速度日加持 午间精确计时 示例+清单+FAQ"))
$r=5
foreach($o in $ovl){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "V4七大升级详解" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$v4why = @(
    "1. 高翻每次训练前都做（5-8min 5x3@50-60%）。高翻是技能——每周1次不够。每周3次轻技术比每周1次重技术进步更快。放在热身之后主项之前 神经系统最清醒。轻重量不累积疲劳 不影响后面主项。",
    "2. 卧推3x5+1AMRAP。前3组保证质量 第4组探底——比4x5更安全（少一组疲劳）更有效（知道今天真实的次数极限）。AMRAP次数是每周进步的标尺。",
    "3. 周五速度日+容量混合。速度卧推8x3@60-65%爆发力——练神经驱动 不累积疲劳 周末能恢复。辅助做容量——上斜卧推+下拉+前蹲 补同一肌群不同角度。",
    "4. 深蹲每4周轮换变式。高杠 低杠 前蹲 箱式。同一变式做8周以上神经系统完全适应 进步变慢。前蹲还能直接帮高翻的接杠姿势。",
    "5. 午间精确计时。主项120s 辅助90s 孤立45-60s 超级组60s。热身10min+主项12min+辅助13min+孤立拉伸10min=45+5min。",
    "6. 减载周换动作。同样动作做轻量=关节应力路径相同 不是真休息。换俯卧撑/高脚杯/弹力带=给关节不同应力分布。",
    "7. 力型兼备热量决策标准。体重每周涨0.25-0.5kg=合适。涨超0.5kg=减200kcal。不动=加200kcal。力量连续2周不涨=查热量和睡眠。"
)
foreach($w in $v4why){SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++}

$r++
SetC $ws1 $r 1 "RPE参考——3年基础 统一标准" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++
SetC $ws1 $r 1 "RPE" $true 10 $white $headerBlue $false
SetC $ws1 $r 2 "感觉" $true 10 $white $headerBlue $false
SetC $ws1 $r 4 "剩余次数" $true 10 $white $headerBlue $false
SetC $ws1 $r 6 "用在哪天" $true 10 $white $headerBlue $false; $r++

$rpeL = New-Object 'System.Collections.Generic.List[Object]'
$rpeL.Add(@("7-8","中等偏强——做完还能做3-4次","周五速度+容量——泵感和爆发力"))
$rpeL.Add(@("8-9","很重——做完还能做1-2次","周一上体力AMRAP——探到真实极限"))
$rpeL.Add(@("技术(不评RPE)","不是力竭——是动作质量","高翻技术——磨动作 不追求RPE"))
$rpeL.Add(@("10","极限力竭","第3周峰值周AMRAP最后一组的最后1次"))
foreach($rp in $rpeL){
    SetC $ws1 $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    SetC $ws1 $r 2 $rp[1] $false 10 0 0 $false
    SetC $ws1 $r 4 $rp[2] $false 10 0 0 $false
    SetC $ws1 $r 6 $rp[3] $false 10 0 0 $true; MC $ws1 $r 6 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$pr = New-Object 'System.Collections.Generic.List[Object]'
$pr.Add(@("1","高翻是技能 高频低强度学习","每次训练前5-8分钟磨技术。不是练到累——是练到准。技术对了重量自然来。"))
$pr.Add(@("2","力型兼备 两个维度一起走","周一冲大重量(力) 周五做速度+容量(型)。同一个人 两种刺激 两个维度都在涨。"))
$pr.Add(@("3","午间效率=质量不妥协","45分钟严格计时——但热身(含高翻)不跳过 动作质量不降。午休不等人。"))
foreach($p in $pr){
    SetC $ws1 $r 1 $p[0] $false 10 0 0 $false
    SetC $ws1 $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws1 $r 3 $p[2] $false 10 0 0 $true; MC $ws1 $r 3 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$ws1.Range("A:A").ColumnWidth=30; $ws1.Range("B:I").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划 V4
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划 V4——DUP 3天 每次含高翻技术 速度日+AMRAP" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "每次训练前做高翻技术5x3@50-60%（热身之后 主项之前）。周一=上体力(3x5+AMRAP) 周三=下肢力+高翻 周五=速度+容量。组间严格计时。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# Helper function
function Write-Day($ws,$r0,$dayTitle,$daySub,$acts){
    $r=$r0
    SetC $ws $r 1 $dayTitle $true 13 $white $darkBg $true
    MC $ws $r 1 9; $ws.Range("A$($r):I$($r)").RowHeight=28
    $ws.Range("A$($r):I$($r)").Font.Color=$white; $ws.Range("A$($r):I$($r)").Interior.Color=$darkBg
    $r++
    SetC $ws $r 1 $daySub $false 9 $headerBlue $lightGray $true
    MC $ws $r 1 9; $ws.Range("A$($r):I$($r)").RowHeight=20; $r++

    SetC $ws $r 1 "序" $true 10 $white $headerBlue $false
    SetC $ws $r 2 "动作" $true 10 $white $headerBlue $false
    SetC $ws $r 3 "组数" $true 10 $white $headerBlue $false
    SetC $ws $r 4 "次数" $true 10 $white $headerBlue $false
    SetC $ws $r 5 "负荷/RPE" $true 10 $white $headerBlue $false
    SetC $ws $r 6 "休息" $true 10 $white $headerBlue $false
    SetC $ws $r 7 "计时" $true 10 $white $headerBlue $false
    SetC $ws $r 8 "力/型" $true 10 $white $headerBlue $false
    SetC $ws $r 9 "技术要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($r):I$($r)").RowHeight=26; $r++

    $i=1
    foreach($a in $acts){
        SetC $ws $r 1 $i $false 10 0 $lightGray $false
        SetC $ws $r 2 $a[0] $true 10 $headerBlue 0 $false
        SetC $ws $r 3 $a[1] $false 10 0 0 $false
        SetC $ws $r 4 $a[2] $false 10 0 0 $false
        SetC $ws $r 5 $a[3] $false 10 0 0 $false
        SetC $ws $r 6 $a[4] $false 10 0 0 $false
        SetC $ws $r 7 $a[5] $false 9 0 0 $false
        SetC $ws $r 8 $a[6] $false 9 0 0 $false
        SetC $ws $r 9 $a[7] $false 9 0 0 $true
        $ws.Range("A$($r):I$($r)").RowHeight=50; SB $ws $r 1 9; $r++; $i++
    }
    return $r+2
}

# 周一
$monActs = New-Object 'System.Collections.Generic.List[Object]'
$monActs.Add(@("高翻技术(先做!)","5组","3次","50%极限 RPE=技术不评","60s","4min","力+技能","髋爆发+耸肩+展体。不是用手臂拉——用髋和三倍伸展。轻重量练速度。"))
$monActs.Add(@("杠铃卧推(主项)","3组+1AMRAP","5次","85%1RM RPE8-9 第4组AMRAP","120s","8min","力的标尺","起桥 肩胛收紧 瓦式。前3组标准5次 第4组做到技术变形前停。记录AMRAP次数。"))
$monActs.Add(@("负重引体向上","3组","5-8次","自重+负重 RPE8","120s","6min","型的宽度","全幅度——底部伸直 顶部下巴过杠。做不了8次=减负重 能做8+=加负重。"))
$monActs.Add(@("站姿杠铃推举","3组","6-8次","RPE7-8","90s","5min","力的辅助","核心收紧 不塌腰。杠铃从锁骨前推至头顶 不锁肘。"))
$monActs.Add(@("杠铃俯身划船","3组","8-10次","RPE7-8","90s","5min","型的厚度","俯身45-60度 杠铃贴大腿前侧拉。引体=垂直拉 划船=水平拉 全角度。"))
$monActs.Add(@("下压+弯举超级组","各3组","各12-15次","轻-中 RPE7","组内不休息 组间60s","5min","型的手臂","下压做完直接弯举 休息60秒。大臂固定不借力。"))
$r=Write-Day $ws2 4 "周一  上体力（大重量 3x5+1AMRAP）" "先练高翻技术5x3@50%(4min)  主项卧推3x5+1AMRAP  辅助90s  孤立45-60s  组间严格计时。" $monActs

# 周三
$wedActs = New-Object 'System.Collections.Generic.List[Object]'
$wedActs.Add(@("高翻技术(先做! 比周一稍重)","5组","3次","50-55%极限 技术优先","60-90s","5min","力+技能","比周一稍重 但技术仍然是第一位。接杠要稳 杠铃轨迹要直。"))
$wedActs.Add(@("杠铃深蹲(主项 高杠)","4组","5次","80%1RM RPE7-8","120s","8min","力的核心","深度平行地面 膝盖不内扣。当前周期=高杠 躯干更竖直 股四主导。"))
$wedActs.Add(@("杠铃罗马尼亚硬拉","3组","8-10次","RPE7-8","90s","5min","型的后侧","膝盖微屈 全程直背。腘绳肌+臀=下肢后侧线条。和深蹲互补。"))
$wedActs.Add(@("悬垂举腿","3组","12-15次","自重","45s","3min","核心地基","下腹带动骨盆上卷 不靠惯性。高翻和深蹲都需要核心稳定。"))
$wedActs.Add(@("站姿提踵","3组","15-20次","RPE7","30s","2min","小腿收尾","顶峰停1秒 慢放。"))
$r=Write-Day $ws2 $r "周三  下肢力+高翻（技能+力量）" "高翻技术5x3@50-55%(5min)  深蹲4x5@80%120s  辅助90s  孤立45s  组间严格计时。" $wedActs

# 周五
$friActs = New-Object 'System.Collections.Generic.List[Object]'
$friActs.Add(@("高翻技术(先做!)","5组","3次","50%极限 练速度","60s","4min","力+技能","轻重量 追求杠铃速度快。和三倍伸展的爆发力感觉。"))
$friActs.Add(@("速度杠铃卧推(主项)","8组","3次","60-65%1RM 爆发力推","60s","8min","力的神经驱动","每组3次 爆发推起 控制下放。不累积疲劳——练神经 不是练肌肉。速度变慢=太重了 减重。"))
$friActs.Add(@("上斜哑铃卧推","3组","10-12次","RPE7-8","90s","5min","型的上胸","凳角30-45度。和速度卧推互补——速度练神经 上斜补容量。"))
$friActs.Add(@("高位下拉(宽握)","3组","12-15次","RPE7","60s","4min","型的宽度","宽握 杆拉至锁骨。和周一引体互补——轻量高次数 背阔肌外沿。"))
$friActs.Add(@("杠铃前蹲","3组","8-10次","RPE7-8","90s","5min","力+型+高翻辅助","杠铃架在锁骨前 躯干更竖直。股四头主导。同时强化高翻接杠姿势。"))
$friActs.Add(@("侧平举+弯举超级组","各3组","15次+12次","轻-中 RPE7","组内不休息 组间45s","4min","型的手臂+肩宽","侧平举做完直接弯举 休息45秒。短间歇=泵感拉满。"))
$r=Write-Day $ws2 $r "周五  全身速度+容量（Dynamic Effort + Hypertrophy）" "高翻技术5x3@50%(4min)  速度卧推8x3@60-65%60s  容量辅助90s  孤立45s  组间严格计时。" $friActs

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=17
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=9
$ws2.Range("E:E").ColumnWidth=15; $ws2.Range("F:F").ColumnWidth=9
$ws2.Range("G:G").ColumnWidth=8; $ws2.Range("H:H").ColumnWidth=14
$ws2.Range("I:I").ColumnWidth=46
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进决策树 V4
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进决策树"

SetC $ws3 1 1 "渐进决策树  高翻进阶  深蹲变式轮换  器械突破" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34
SetC $ws3 2 1 "训练像做实验——每次都是数据点。V4新增：深蹲每4周轮换变式 高翻技术判断树。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 9

$r=4
SetC $ws3 $r 1 "4周渐进（每次含高翻技术 第1-3周渐进 第4周减载）" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "高翻(每练必做)" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "上体力(周一)" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "下肢力(周三)" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "速度+容量(周五)" $true 10 $white $headerBlue $false; $r++

$progL = New-Object 'System.Collections.Generic.List[Object]'
$progL.Add(@("1","50% 5x3 技术","85% 3x5+AMRAP 建基准","80% 4x5 高杠","速度60% 8x3 容量70%"))
$progL.Add(@("2","55% 5x3 稍重","+2.5-5kg AMRAP超第1周","+5kg","速度65% 容量缩短休息或+1组"))
$progL.Add(@("3","60% 4x3 挑战","再+2.5-5kg 3x3+AMRAP","87.5% 4x3","速度65%挑战更快速 容量再缩短"))
$progL.Add(@("4","50% 3x3 练速度","60% 3x5 换俯卧撑","60% 3x5 高脚杯","50%轻量 换弹力带"))
foreach($pw in $progL){
    SetC $ws3 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pw[1] $false 9 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pw[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $pw[3] $false 9 0 0 $false
    SetC $ws3 $r 6 $pw[4] $false 9 0 0 $true; MC $ws3 $r 6 9
    $ws3.Range("A$($r):I$($r)").RowHeight=30; SB $ws3 $r 1 9; $r++
}

$r+=2
SetC $ws3 $r 1 "深蹲变式轮换——每4周一换 避免适应性停滞" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "周期" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "变式" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "特点" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "适应期" $true 10 $white $headerBlue $false
SetC $ws3 $r 8 "对高翻的帮助" $true 10 $white $headerBlue $false; $r++

$squatL = New-Object 'System.Collections.Generic.List[Object]'
$squatL.Add(@("第1-4周(本周期)","高杠深蹲","躯干更竖直 股四头肌主导","你现在用的——已适应 正是最佳期","和接杠姿势相似——直接帮助高翻"))
$squatL.Add(@("第5-8周","低杠深蹲","躯干更前倾 后侧链主导","新刺激——前2周适应 后2周冲","强化后侧链——高翻启动更有力"))
$squatL.Add(@("第9-12周","前蹲","杠铃架锁骨前 核心要求最高","最接近高翻接杠姿势——你的弱项","直接强化高翻接杠——最值得练"))
$squatL.Add(@("第13-16周","箱式深蹲","坐箱起 练底部爆发力","新刺激——突破粘滞点","练底部爆发力——高翻的启动阶段"))
foreach($sc in $squatL){
    SetC $ws3 $r 1 $sc[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $sc[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $sc[2] $false 10 0 0 $false; MC $ws3 $r 4 5
    SetC $ws3 $r 6 $sc[3] $false 10 0 0 $false; MC $ws3 $r 6 7
    SetC $ws3 $r 8 $sc[4] $false 10 0 0 $true; MC $ws3 $r 8 9
    $ws3.Range("A$($r):I$($r)").RowHeight=28; SB $ws3 $r 1 9; $r++
}

$r+=2
SetC $ws3 $r 1 "高翻进阶——每次训练前5x3 技术优先" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$cleanL = New-Object 'System.Collections.Generic.List[Object]'
$cleanL.Add(@("动作流畅 杠铃轨迹直 接杠稳","技术对了——可以尝试+2.5kg","但只在周三这天稍加 周一和周五保持轻量练速度"))
$cleanL.Add(@("手臂过早发力弯肘 用二头拉","技术不对——三倍伸展没出来","减重练髋爆发+耸肩。找腿发力 手只是引导的感觉"))
$cleanL.Add(@("接杠时杠铃砸锁骨 手腕疼","接杠位置不对 前架不稳","练前蹲+杠铃耸肩 找锁骨架杠的感觉。不加重。"))
$cleanL.Add(@("连续6次训练(2周) 技术流畅无卡顿","可以进入第2周——所有训练日都稍加重","周一50% 55% 周三55% 60% 周五50% 55%"))
$cleanL.Add(@("今天状态差/下午累了/午休不够","不练高翻——改练杠铃耸肩+前蹲轻量","高翻是最高神经需求——状态不好练了反而退步"))
foreach($ct in $cleanL){
    SetC $ws3 $r 1 $ct[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $ct[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $ct[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=28; SB $ws3 $r 1 9; $r++
}

$r+=2
SetC $ws3 $r 1 "各训练日渐进决策树" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$treeL = New-Object 'System.Collections.Generic.List[Object]'
$treeL.Add(@("周一 3x5轻松+AMRAP比上周+2次","变强了","下次+2.5-5kg"))
$treeL.Add(@("周一 AMRAP次数比上周少","累积疲劳或状态差","维持重量 不加重——不退步"))
$treeL.Add(@("周三 高翻+深蹲后还有余力","可以加重","下周深蹲+2.5-5kg"))
$treeL.Add(@("周五 速度卧推速度变慢","太重了","减重5%——速度优先于重量"))
$treeL.Add(@("周五 容量部分休息觉得太长","代谢压力不够","缩短10-15秒"))
$treeL.Add(@("某动作连续4周没进步","需要换刺激","换变式——同肌群不同角度"))
$treeL.Add(@("连续2周睡不好+训练状态差","该减载了","提前进入第4周减载"))
foreach($ta in $treeL){
    SetC $ws3 $r 1 $ta[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $ta[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $ta[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=30; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "午间热身  主动恢复  减载换动作" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

$r=3
SetC $ws4 $r 1 "午间热身——含高翻技术 10分钟 不跳过" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "时间" $true 10 $white $headerBlue $false; $r++

$wuL = New-Object 'System.Collections.Generic.List[Object]'
$wuL.Add(@("一般热身","划船机/单车——轻度有氧升温","3min"))
$wuL.Add(@("动态拉伸","猫式+最伟大伸展+摆腿+肩环绕","2min"))
$wuL.Add(@("神经激活","药球砸地/跳箱(低)/弹力带侧走","1min"))
$wuL.Add(@("高翻技术(每次必做!)","5x3@50-60% 轻重量练三倍伸展+接杠","4min"))
foreach($w in $wuL){
    SetC $ws4 $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $w[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $w[2] $false 10 0 0 $false; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=24; SB $ws4 $r 1 9; $r++
}

$r+=2
SetC $ws4 $r 1 "休息日主动恢复——打工人 不用复杂" $true 14 $white $softBlue $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$activeV4 = @(
    "下班后不需要专门做恢复。做这几件事就够了：",
    "  泡沫轴10min：小腿 股四 腘绳 臀 背 胸 每部位30秒",
    "  散步20-30min：心率<120bpm 不是训练 是恢复",
    "  睡够7-8h：午间训练表现直接和睡眠质量挂钩——这是最重要的恢复手段"
)
foreach($ar in $activeV4){SetC $ws4 $r 1 $ar $false 10 0 0 $true; MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=20; $r++}

$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 不是做轻一点" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 9; $r++

$deloadL = New-Object 'System.Collections.Generic.List[Object]'
$deloadL.Add(@("高翻(重)","杠铃耸肩+前蹲(轻量 练分解)","2组x8-10"))
$deloadL.Add(@("杠铃卧推","俯卧撑(3秒下+爆发推)","3组x力竭60%"))
$deloadL.Add(@("杠铃深蹲","高脚杯深蹲(轻哑铃)","2组x10-12"))
$deloadL.Add(@("杠铃划船","哑铃单臂划船(轻量)","2组x10/侧"))
$deloadL.Add(@("负重引体","高位下拉(轻量)","2组x10-12"))
$deloadL.Add(@("有氧","低强度散步/休闲骑","20-30min <120bpm"))
foreach($dl in $deloadL){
    SetC $ws4 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $dl[2] $false 10 0 0 $true; MC $ws4 $r 4 9
    $ws4.Range("A$($r):I$($r)").RowHeight=26; SB $ws4 $r 1 9; $r++
}

$ws4.Range("A:A").ColumnWidth=22; $ws4.Range("B:I").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食与规则 V4
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与规则"

SetC $ws5 1 1 "力型兼备营养  热量决策标准  判断规则" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

$r=3
SetC $ws5 $r 1 "午间训练——营养窗口策略" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "时机" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws5 $r 2 5
SetC $ws5 $r 6 "逻辑" $true 10 $white $headerBlue $false; MC $ws5 $r 6 9; $r++

$periL = New-Object 'System.Collections.Generic.List[Object]'
$periL.Add(@("练前 11:30","香蕉+蛋白粉/全麦吐司2片+2蛋","轻食不撑胃 提供训练能量"))
$periL.Add(@("练后 13:00前","米饭+鸡胸肉/蛋白粉2勺+香蕉+牛奶","下午上班前关恢复窗口"))
$periL.Add(@("下午 15:00","坚果/水果/酸奶(饿了吃)","不饿可以不吃"))
$periL.Add(@("晚餐","正常吃：碳水一拳+蛋白质+蔬菜","和平时一样"))
foreach($pw in $periL){
    SetC $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pw[1] $false 10 0 0 $true; MC $ws5 $r 2 5
    SetC $ws5 $r 6 $pw[2] $false 10 0 0 $true; MC $ws5 $r 6 9
    $ws5.Range("A$($r):I$($r)").RowHeight=30; SB $ws5 $r 1 9; $r++
}

$r++
SetC $ws5 $r 1 "全天：总热量2500-2700 蛋白质130-150g(2.0-2.2g/kg) 碳水300-340g 脂肪55-65g 水3L。干净增肌——不做脏增肌。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=22; $r++

$r+=2
SetC $ws5 $r 1 "力型兼备——热量调整决策标准 V4" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "观察指标" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "信号" $true 10 $white $headerBlue $false; MC $ws5 $r 2 4
SetC $ws5 $r 5 "判断" $true 10 $white $headerBlue $false
SetC $ws5 $r 7 "调整" $true 10 $white $headerBlue $false; MC $ws5 $r 7 9; $r++

$calL = New-Object 'System.Collections.Generic.List[Object]'
$calL.Add(@("体重趋势","每周涨0.25-0.5kg","合适","继续"))
$calL.Add(@("体重趋势","每周涨0.5kg以上","偏高","减200kcal(碳水减半碗饭)"))
$calL.Add(@("体重趋势","体重不动","偏低","加200kcal(碳水加半碗饭)"))
$calL.Add(@("体重趋势","体重下降","不够","加300kcal(碳水加一碗饭)"))
$calL.Add(@("力量趋势","卧推/深蹲每周能+2.5-5kg","合适","继续"))
$calL.Add(@("力量趋势","力量连续2周不涨","查原因","先查睡眠 再加碳水"))
$calL.Add(@("体脂趋势","腹肌线条还在","合适","继续"))
$calL.Add(@("体脂趋势","腹肌线条模糊","偏高","减200kcal"))
$calL.Add(@("体脂趋势","更清晰但力量不涨","偏低","加200kcal"))
foreach($cd in $calL){
    SetC $ws5 $r 1 $cd[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $cd[1] $false 10 0 0 $true; MC $ws5 $r 2 4
    SetC $ws5 $r 5 $cd[2] $false 10 0 0 $false
    SetC $ws5 $r 7 $cd[3] $false 10 0 0 $true; MC $ws5 $r 7 9
    $ws5.Range("A$($r):I$($r)").RowHeight=26; SB $ws5 $r 1 9; $r++
}

$r+=2
SetC $ws5 $r 1 "训练与身体信号" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "信号" $true 10 $white $headerBlue $false; MC $ws5 $r 1 2
SetC $ws5 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$rulesL = New-Object 'System.Collections.Generic.List[Object]'
$rulesL.Add(@("AMRAP次数比上周+2次以上","变强了——下次+2.5-5kg"))
$rulesL.Add(@("AMRAP次数没变或下降","维持重量 不退步"))
$rulesL.Add(@("高翻连续6次训练流畅","+2.5kg——但只在周三这天加"))
$rulesL.Add(@("速度卧推时速度变慢","减重5%——速度优先于重量"))
$rulesL.Add(@("午间训练没力气","查11:30那顿——是不是没吃够"))
$rulesL.Add(@("中午练完下午犯困","练后碳水减半碗饭 蛋白不变"))
$rulesL.Add(@("状态极差","降级：热身+高翻技术+主项3组轻量+拉伸"))
foreach($rv in $rulesL){
    SetC $ws5 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $rv[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=26; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用与复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  力型兼备复盘  自主调节" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

$r=3
SetC $ws6 $r 1 "如果这周只能2练——精简版" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "选项" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "训练日1" $true 10 $white $headerBlue $false; MC $ws6 $r 2 4
SetC $ws6 $r 5 "训练日2" $true 10 $white $headerBlue $false; MC $ws6 $r 5 9; $r++

$simL = New-Object 'System.Collections.Generic.List[Object]'
$simL.Add(@("2练","上体力(精简)：高翻+卧推3x5+1AMRAP+引体+推举+划船+手臂","下肢力+全身(精简)：高翻+深蹲4x5+RDL+前蹲3x10+下拉+侧平举"))
$simL.Add(@("原则","高翻不跳——每练必做。保留核心复合 压缩辅助。",""))
foreach($so in $simL){
    SetC $ws6 $r 1 $so[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $so[1] $false 10 0 0 $true; MC $ws6 $r 2 4
    SetC $ws6 $r 5 $so[2] $false 10 0 0 $true; MC $ws6 $r 5 9
    $ws6.Range("A$($r):I$($r)").RowHeight=42; SB $ws6 $r 1 9; $r++
}

$r+=2
SetC $ws6 $r 1 "4周后复盘——力型兼备双重指标" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "第1周vs第3周" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 6 9; $r++

$revL = New-Object 'System.Collections.Generic.List[Object]'
$revL.Add(@("力-卧推AMRAP","第1周做了几次 第3周做了几次？","涨=加重。平/降=维持或换变式"))
$revL.Add(@("力-深蹲4x5","重量涨了多少？","涨=加重。没涨=查恢复"))
$revL.Add(@("力-高翻技术","流畅了吗？能+2.5kg吗？","流畅=继续。仍卡=继续磨"))
$revL.Add(@("型-体重","涨0.5-1kg净肌肉？","涨=继续。没涨=加碳水"))
$revL.Add(@("型-照片","第1天vs第28天","镜子比秤真实"))
$revL.Add(@("力型兼备","力量涨+体型变=正在发生","继续——这套框架可以一直用"))
foreach($rv in $revL){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 5
    SetC $ws6 $r 6 $rv[2] $false 10 0 0 $true; MC $ws6 $r 6 9
    $ws6.Range("A$($r):I$($r)").RowHeight=26; SB $ws6 $r 1 9; $r++
}

$r+=2
SetC $ws6 $r 1 "自主调节框架——3年基础 你可以自己迭代" $true 14 $white $accentGreen $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

$selfV4 = @(
    "1. 定分化：每周3天 DUP(上体力+下肢力+高翻+速度+容量)",
    "2. 高翻每练必做：5-8分钟 5x3@50-60%。技能训练 高频低强度",
    "3. 定周期：3周渐进+1周减载(换动作) 4周一循环。深蹲每4周轮换变式",
    "4. 定RPE：周一8-9(重) 周三7-8(重+技能) 周五不评RPE(速度日看速度不看RPE)",
    "5. 力型兼备热量决策：体重趋势+力量趋势+体脂趋势 三个指标联动调整",
    "你已经有了一套能跑一辈子的训练框架——它会跟着你一起进化。"
)
foreach($s in $selfV4){SetC $ws6 $r 1 $s $false 10 0 0 $true; MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=22; $r++}

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# ============================================
# SHEET 7: 示例检查FAQ V4
# ============================================
$ws7=$wb.Worksheets.Add();$ws7.Name="示例检查FAQ"

SetC $ws7 1 1 "第一周训练示例  每周检查清单  常见问题FAQ" $true 16 $headerBlue 0 $true
MC $ws7 1 1 8; $ws7.Range("A1:H1").RowHeight=34
SetC $ws7 2 1 "拿到计划不知道怎么开始？直接照着下面做。练完一周花1分钟填检查清单。遇到问题翻FAQ。" $false 10 $headerBlue $lightGray $true
MC $ws7 2 1 8

# 第一周 周一示例
$r=4
SetC $ws7 $r 1 "第一周训练示例——直接照着做（重量按你自己的实际调整）" $true 14 $white $accentGreen $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws7 $r 1 "周一 上体力（大重量）——示例" $true 12 $headerBlue $lightBlue $true; MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws7 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws7 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws7 $r 2 3
SetC $ws7 $r 4 "怎么做" $true 9 $white $headerBlue $false; MC $ws7 $r 4 6
SetC $ws7 $r 7 "休息" $true 9 $white $headerBlue $false
SetC $ws7 $r 8 "感觉" $true 9 $white $headerBlue $false; $r++

$monEx = New-Object 'System.Collections.Generic.List[Object]'
$monEx.Add(@("1","热身","划船机3min 拉伸2min 开合跳1min","—","—"))
$monEx.Add(@("2","高翻技术","5组x3次@50%1RM(比如你能翻40kg 用20kg)","60s","轻快——磨三倍伸展"))
$monEx.Add(@("3","杠铃卧推","40kgx5次x3组+第4组做到力竭(记次数!)","120s","重而可控"))
$monEx.Add(@("4","负重引体","自重+5kgx5次x3组","120s","能做8次=下次加重"))
$monEx.Add(@("5","杠铃推举","25kgx6次x3组","90s","核心收紧"))
$monEx.Add(@("6","杠铃划船","30kgx8次x3组","90s","俯身45度 肩胛夹紧"))
$monEx.Add(@("7","下压+弯举超级组","各15次x3组 做完下压直接弯举 休息60s","组内不休息","手臂泵感"))
$monEx.Add(@("8","拉伸","胸+背+肩+髋屈肌+腘绳肌 各30s","—","—"))
foreach($me in $monEx){
    SetC $ws7 $r 1 $me[0] $true 9 $headerBlue $lightGray $false
    SetC $ws7 $r 2 $me[1] $true 9 $white $darkBg $false; MC $ws7 $r 2 3
    SetC $ws7 $r 4 $me[2] $false 9 0 0 $true; MC $ws7 $r 4 6
    SetC $ws7 $r 7 $me[3] $false 9 0 0 $false
    SetC $ws7 $r 8 $me[4] $false 9 0 0 $false
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

$r++
SetC $ws7 $r 1 "周三 下肢力+高翻——示例" $true 12 $headerBlue $lightBlue $true; MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws7 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws7 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws7 $r 2 3
SetC $ws7 $r 4 "怎么做" $true 9 $white $headerBlue $false; MC $ws7 $r 4 6
SetC $ws7 $r 7 "休息" $true 9 $white $headerBlue $false
SetC $ws7 $r 8 "感觉" $true 9 $white $headerBlue $false; $r++

$wedEx = New-Object 'System.Collections.Generic.List[Object]'
$wedEx.Add(@("1","热身","划船3min+拉伸2min+神经激活1min","—","—"))
$wedEx.Add(@("2","高翻技术","5组x3次@50-55%(稍重 比如25kg)","60-90s","技术仍第一位——接杠要稳"))
$wedEx.Add(@("3","杠铃深蹲(高杠)","50kgx5次x4组","120s","深度平行 膝盖不内扣"))
$wedEx.Add(@("4","杠铃RDL","35kgx8次x3组","90s","全程直背 腘绳肌拉伸"))
$wedEx.Add(@("5","悬垂举腿","15次x3组","45s","下腹发力 不靠惯性"))
$wedEx.Add(@("6","站姿提踵+拉伸","提踵20次x3组30s休息+拉伸5min","—","小腿收尾"))
foreach($we in $wedEx){
    SetC $ws7 $r 1 $we[0] $true 9 $headerBlue $lightGray $false
    SetC $ws7 $r 2 $we[1] $true 9 $white $darkBg $false; MC $ws7 $r 2 3
    SetC $ws7 $r 4 $we[2] $false 9 0 0 $true; MC $ws7 $r 4 6
    SetC $ws7 $r 7 $we[3] $false 9 0 0 $false
    SetC $ws7 $r 8 $we[4] $false 9 0 0 $false
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

$r++
SetC $ws7 $r 1 "周五 速度+容量——示例" $true 12 $headerBlue $lightBlue $true; MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws7 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws7 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws7 $r 2 3
SetC $ws7 $r 4 "怎么做" $true 9 $white $headerBlue $false; MC $ws7 $r 4 6
SetC $ws7 $r 7 "休息" $true 9 $white $headerBlue $false
SetC $ws7 $r 8 "感觉" $true 9 $white $headerBlue $false; $r++

$friEx = New-Object 'System.Collections.Generic.List[Object]'
$friEx.Add(@("1","热身+高翻","高翻5组x3次@50%20kg 4min","60s","轻量练速度"))
$friEx.Add(@("2","速度杠铃卧推","25kgx3次x8组(60-65%1RM)","60s","爆发推起 轻快——速度变慢=太重了"))
$friEx.Add(@("3","上斜哑铃卧推","单边12kgx10次x3组","90s","上胸+肩前束"))
$friEx.Add(@("4","高位下拉(宽握)","插片选10-12RMx12次x3组","60s","宽握 背阔外沿"))
$friEx.Add(@("5","杠铃前蹲","30kgx8次x3组","90s","股四头主导 帮高翻接杠"))
$friEx.Add(@("6","侧平举+弯举超级组","侧平举15次+弯举12次x3组 组内不休息","组间45s","手臂+肩收尾"))
$friEx.Add(@("7","拉伸","全身拉伸5min","—","—"))
foreach($fe in $friEx){
    SetC $ws7 $r 1 $fe[0] $true 9 $headerBlue $lightGray $false
    SetC $ws7 $r 2 $fe[1] $true 9 $white $darkBg $false; MC $ws7 $r 2 3
    SetC $ws7 $r 4 $fe[2] $false 9 0 0 $true; MC $ws7 $r 4 6
    SetC $ws7 $r 7 $fe[3] $false 9 0 0 $false
    SetC $ws7 $r 8 $fe[4] $false 9 0 0 $false
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

$r++
SetC $ws7 $r 1 "  重量按你自己的实际调整——上面写的是示例。你练了3-4年 知道自己的卧推/深蹲大概多少。如果不知道85%1RM是多少：用你能做5次但第5次很吃力的重量就行。" $false 10 $accentGreen 0 $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26; $r++

# 每周检查清单
$r+=2
SetC $ws7 $r 1 "每周检查清单——练完周五后花1分钟填" $true 14 $white $accentOrange $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws7 $r 1 "检查项" $true 10 $white $headerBlue $false; MC $ws7 $r 1 4
SetC $ws7 $r 5 "第1周" $true 10 $white $headerBlue $false
SetC $ws7 $r 6 "第2周" $true 10 $white $headerBlue $false
SetC $ws7 $r 7 "第3周" $true 10 $white $headerBlue $false
SetC $ws7 $r 8 "第4周" $true 10 $white $headerBlue $false; $r++

$ckl = New-Object 'System.Collections.Generic.List[Object]'
$ckl.Add(@("本周体重涨了还是跌了？","___kg","___kg","___kg","___kg"))
$ckl.Add(@("周一卧推AMRAP比上周多了几次？","___次","___次","___次","___次"))
$ckl.Add(@("周三深蹲重量涨了吗？","___kg","___kg","___kg","___kg"))
$ckl.Add(@("高翻技术流畅吗？有没有砸锁骨？","是/否","是/否","是/否","是/否"))
$ckl.Add(@("本周睡眠平均几小时？","___h","___h","___h","___h"))
$ckl.Add(@("本周哪次训练感觉特别好/特别差？","—","—","—","—"))
foreach($cl in $ckl){
    SetC $ws7 $r 1 $cl[0] $true 10 $headerBlue $lightGray $false; MC $ws7 $r 1 4
    SetC $ws7 $r 5 $cl[1] $false 10 0 0 $false
    SetC $ws7 $r 6 $cl[2] $false 10 0 0 $false
    SetC $ws7 $r 7 $cl[3] $false 10 0 0 $false
    SetC $ws7 $r 8 $cl[4] $false 10 0 0 $false
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

# FAQ
$r+=2
SetC $ws7 $r 1 "常见问题FAQ——练到一半遇到问题 翻这里" $true 14 $white $darkBg $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws7 $r 1 "问题" $true 10 $white $headerBlue $false; MC $ws7 $r 1 2
SetC $ws7 $r 3 "答案" $true 10 $white $headerBlue $false; MC $ws7 $r 3 8; $r++

$faqL = New-Object 'System.Collections.Generic.List[Object]'
$faqL.Add(@("Q1:周一卧推AMRAP做到第几次算好？","第1周能多做2-3次就算好。比如3x5做完 第4组做了7次 那下次可以+2.5kg。"))
$faqL.Add(@("Q2:周三深蹲4x5 最后一组做不完怎么办？","做不完就做3次 记录4x5(3)。下次维持重量 不加重。连续2周能做完 再+2.5kg。"))
$faqL.Add(@("Q3:周五速度卧推8x3 做到第6组没力了？","减重5% 继续做完。速度日的核心是完成8组 不是用多重。"))
$faqL.Add(@("Q4:中午练完下午犯困怎么办？","练后碳水减半碗。如果还困 把咖啡从练前移到练后喝。"))
$faqL.Add(@("Q5:高翻总是砸锁骨怎么办？","减重 练前蹲+耸肩。高翻是技能不是力量——砸锁骨说明你接杠时机不对。"))
$faqL.Add(@("Q6:如果不知道85%1RM是多少？","用你能做5次 但第5次很吃力的重量。那个重量大概就是你的85%1RM。"))
$faqL.Add(@("Q7:减载周做完感觉根本没练到——正常吗？","完全正常。减载周就该轻松——不轻松说明你前三周没练到位。"))
foreach($fq in $faqL){
    SetC $ws7 $r 1 $fq[0] $true 10 $headerBlue $lightGray $false; MC $ws7 $r 1 2
    SetC $ws7 $r 3 $fq[1] $false 10 0 0 $true; MC $ws7 $r 3 8
    $ws7.Range("A$($r):H$($r)").RowHeight=32; SB $ws7 $r 1 8; $r++
}

$ws7.Range("A:A").ColumnWidth=26; $ws7.Range("B:H").ColumnWidth=18
Write-Host "Sheet 7 done"

# Save
$savePath="D:\Codex\members\瑞克\瑞克_力型兼备方案_V4.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
