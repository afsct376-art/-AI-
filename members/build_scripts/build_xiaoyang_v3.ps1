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

function Write-TrainingBlock($ws,$ref,$title,$focus,$tips,$time,$acts){
    SetC $ws $ref 1 $title $true 13 $white $darkBg $true; MC $ws $ref 1 8
    $ws.Range("A$($ref):H$($ref)").RowHeight=26
    $ws.Range("A$($ref):H$($ref)").Font.Color=$white
    $ws.Range("A$($ref):H$($ref)").Interior.Color=$darkBg; $ref++
    SetC $ws $ref 1 "$focus  |  $time 分钟  |  $tips" $false 10 $headerBlue $lightGray $true
    MC $ws $ref 1 8; $ref++
    SetC $ws $ref 1 "序" $true 10 $white $headerBlue $false
    SetC $ws $ref 2 "动作" $true 10 $white $headerBlue $false
    SetC $ws $ref 3 "基础版(固械)" $true 10 $white $headerBlue $false
    SetC $ws $ref 4 "进阶版(自由重量)" $true 10 $white $headerBlue $false
    SetC $ws $ref 5 "重量/RPE" $true 10 $white $headerBlue $false
    SetC $ws $ref 6 "休息" $true 10 $white $headerBlue $false
    SetC $ws $ref 7 "感觉" $true 10 $white $headerBlue $false
    SetC $ws $ref 8 "动作要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):H$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        SetC $ws $ref 1 $i $false 10 0 $lightGray $false
        SetC $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 6;$j++){SetC $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 6)}
        $ws.Range("A$($ref):H$($ref)").RowHeight=58
        SB $ws $ref 1 8; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "蜡笔小杨  减脂塑形方案 V3" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 72kg  |  继续上次的旅程  |  杠铃回归" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你不是重新开始——你是继续" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "半年前你40天瘦了7.5kg。那不是运气——是你做对了一系列事情。体脂从43.1%降到39.5% 腰围小了7cm。现在你72kg 和当时最低点差不多——说明你没反弹 身体记住了那个状态。这次不是重新开始 是继续。上次你学会了怎么瘦下来。这次——杠铃回来了。你之前跟我学的硬拉和划船 身体还记得。" $false 10 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=52; $row++

# V3: 这次升级了什么
$row+=2
SetC $ws0 $row 1 "  V3 这次升级了什么" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  杠铃回归","你之前跟我学过杠铃硬拉、杠铃划船——动作模式还在。这次每个训练日有两条路：基础版(固定器械)和进阶版(自由重量)。状态好就冲 状态差就稳。"),
    @("  夜班方案","你在火车站上班 有夜班——不需要按周一三五硬练。按你的生物钟：白班正常练 夜班前练基础版 夜班后先睡觉 睡醒再说。"),
    @("  强度提升","你有基础——强度可以比零基础高。进阶版用RPE 7-8(有点挑战但不力竭)。杠铃硬拉和划船的技术你之前学过 重新唤醒就行。"),
    @("  饮食从控制转向优化窗口","你饮食已经很自律了——这次不需要'控制' 需要的是'优化'。夜班那顿不吃碳水大餐 训练前后窗口更精确。")
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

# 上次做对了什么 + 踩坑
$row+=2
SetC $ws0 $row 1 "  上次做对了什么  继续" $true 14 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$didRight=@(
    @("  饮食控制","你当时怎么吃的——这次微调就行 不需要重学"),
    @("  规律训练","一周好几次——身体记得。这次杠铃加回来 效果更好"),
    @("  坚持记录","上次拍了照 量了腰围——所以看到7cm变化。这次也拍")
)
foreach($dr in $didRight){
    SetC $ws0 $row 1 $dr[0] $true 10 $accentGreen 0 $false
    SetC $ws0 $row 2 $dr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}

$row++
SetC $ws0 $row 1 "  上次踩过的坑  这次避开" $true 14 $white $accentRed $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentRed
$row++

$didWrong=@(
    @("  太激进","40天7.5kg偏快——这次稳一点 0.5-1kg/周"),
    @("  恢复不够","上次频率偏高——这次3次/周+减载周"),
    @("  停了就全停","上次没过渡——这次有减载+降级+4周后路线图")
)
foreach($dw in $didWrong){
    SetC $ws0 $row 1 $dw[0] $true 10 $accentRed 0 $false
    SetC $ws0 $row 2 $dw[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}

# 四周总览
$row+=2
SetC $ws0 $row 1 "  四周总览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "做什么" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "自由重量" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周目标" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","唤醒+基准","3次/周 全身","基础版为主 唤醒杠铃动作记忆","记录现在的位置——第4周对比"),
    @("2","容量积累","3次/周 全身","状态好→进阶版杠铃","身体适应——同样重量做更多组"),
    @("3","挑战周","3次/周 全身","杠铃硬拉/划船加重量","加到就加 加不到不勉强"),
    @("4","恢复+复盘","3次/周 减载","换动作 轻重量 不碰杠铃","对比基准数据 看涨了多少")
)
foreach($o4i in $o4){
    SetC $ws0 $row 1 $o4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $o4i[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $o4i[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $o4i[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $o4i[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=30; SB $ws0 $row 1 5; $row++
}

# V3: 夜班速查
$row+=2
SetC $ws0 $row 1 "  夜班怎么练（你在火车站上班——按生物钟来）" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "班次" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "什么时候练" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "练什么" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "注意" $true 10 $white $headerBlue $false; $row++

$nightShift=@(
    @("白班","下班后/晚饭前","正常训练 按计划","—"),
    @("夜班(晚上上班)","上班前(下午)","基础版 轻重量","夜班消耗大 强度降一档"),
    @("夜班后(早上下班)","先睡觉——睡醒再练","基础版或休息","不要刚下班就去练！"),
    @("换班日","休息或只做有氧","散步/拉伸","生物钟切换时身体压力最大")
)
foreach($ns in $nightShift){
    SetC $ws0 $row 1 $ns[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $ns[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $ns[2] $false 10 0 0 $false
    SetC $ws0 $row 5 $ns[3] $false 10 0 0 $true; MC $ws0 $row 5 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+设计思路"," Sheet 方案概览"),
    @("每次训练练什么(基础版+进阶版)"," Sheet 训练计划"),
    @("夜班方案(训练+饮食)+拍照"," Sheet 夜班与热身"),
    @("饮食优化+经期情绪管理"," Sheet 饮食与规则"),
    @("状态差/来例假/加班 兜底"," Sheet 备用与复盘"),
    @("4周后怎么看效果+三个选择"," Sheet 备用与复盘")
)
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
$ws=$wb.Worksheets.Add();$ws.Name="方案概览"

SetC $ws 1 1 "蜡笔小杨  减脂塑形方案  概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 8; $ws.Range("A1:H1").RowHeight=36

SetC $ws 2 1 "女  25岁以下  160cm/72kg  BMI 28.1  体脂35%+  40天私教基础(减7.5kg/腰围-7cm)  能硬拉和划船  火车站上班(有夜班)  饮食自律" $false 11 $gold $darkBg $true
MC $ws 2 1 8; $ws.Range("A2:H2").Font.Color=$white; $ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 8

$ov=@(
    @("训练频率","每周3练 全身训练","基础版+进阶版双路径——每天自选"),
    @("训练工具","固定器械 + 杠铃/哑铃自由重量","你之前学过硬拉和划船——杠铃回归"),
    @("周期结构","唤醒+基准  容量积累  挑战周  恢复+复盘","进阶版RPE 7-8(有点挑战但不力竭)"),
    @("特殊适配","夜班作息方案","白班正常练 夜班前练基础版 夜班后先睡"),
    @("方案特点","杠铃回归+双路径+夜班适配+饮食窗口优化+4周后三个选择")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 5
    SetC $ws $r 6 $o[2] $false 10 0 0 $true; MC $ws $r 6 8
    SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "为什么这样设计——杠铃回归 + 双路径" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$design=@(
    "你和零基础不一样——你之前跟我学过杠铃硬拉、杠铃俯身划船。动作模式在你身体里 只是停了一段时间需要唤醒。",
    "所以V3给每个训练日设了两条路：基础版(固定器械)和进阶版(自由重量)。状态好就冲杠铃 夜班后累了就走固定器械。两条路都能到终点 区别只是速度。",
    "你在火车站上班 有夜班——不需要按周一三五硬练。白班正常练 夜班前练基础版 夜班后先睡觉 睡醒再说。",
    "你饮食已经很自律了——这次不需要'控制'，需要的是'优化'：夜班那顿不吃碳水大餐 训练前后窗口更精确。",
    "第1周以固定器械为主——同时用轻重量唤醒杠铃动作记忆。第2-3周逐步加入进阶版。第4周减载——换动作 不碰杠铃。"
)
foreach($d in $design){
    SetC $ws $r 1 $d $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws $r 1 "四周节奏" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

SetC $ws $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws $r 3 "训练日" $true 10 $white $headerBlue $false
SetC $ws $r 4 "自由重量" $true 10 $white $headerBlue $false
SetC $ws $r 5 "组数/重量" $true 10 $white $headerBlue $false
SetC $ws $r 7 "目标" $true 10 $white $headerBlue $false; $r++

$w4=@(
    @("1","唤醒+基准","3次","基础版为主 轻重量重温杠铃动作","3组 中等","找到现在位置"),
    @("2","容量积累","3次","状态好上进阶版杠铃","3 4组 次数上限","身体适应 更强了"),
    @("3","挑战周","3次","杠铃硬拉/划船尝试加重","3-4组 RPE 7-8","试了就比不试强"),
    @("4","恢复+复盘","3次","换动作 轻量 不碰杠铃","2-3组 RPE 4-5","对比基准 看变化")
)
foreach($w4i in $w4){
    SetC $ws $r 1 $w4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $w4i[1] $true 10 $white $darkBg $false
    SetC $ws $r 3 $w4i[2] $false 10 0 0 $false
    SetC $ws $r 4 $w4i[3] $false 10 0 0 $false
    SetC $ws $r 5 $w4i[4] $false 10 0 0 $false
    SetC $ws $r 7 $w4i[5] $false 10 0 0 $true
    $ws.Range("A$($r):H$($r)").RowHeight=28; SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$pr=@(
    @("1","杠铃回来了","杠铃硬拉和划船的技术你之前学过——身体还记得。第1周用轻重量唤醒 第2-3周逐步加。进阶版和基础版——每天自己选。"),
    @("2","按你的生物钟来","你在火车站上班 有夜班。不需要硬套周一三五——白班正常练 夜班前练基础版 夜班后先睡觉。换班日休息。"),
    @("3","上次的经验是这次的武器","上次做对的继续 上次踩坑的避开。这次有杠铃 有双路径 有夜班方案——上次没有这些。你不是从头摸索 是带着经验升级。")
)
foreach($p in $pr){
    SetC $ws $r 1 $p[0] $false 10 0 0 $false
    SetC $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws $r 3 $p[2] $false 10 0 0 $true; MC $ws $r 3 8
    $ws.Range("A$($r):H$($r)").RowHeight=26; SB $ws $r 1 8; $r++
}

$ws.Range("A:A").ColumnWidth=30; $ws.Range("B:H").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划（双路径）
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划  第1-4周  基础版 + 进阶版" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "每周3练（周一/三/五或按夜班调整）。每个动作有两条路径：基础版=固定器械(状态一般时)  进阶版=自由重量(状态好时)。每天训练前自选。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# Week 1
$r=4
SetC $ws2 $r 1 "第1周  唤醒+基准  3组 12-15次  记录所有数据" $true 14 $white $accentGreen $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws2 $r 1 "第1周以基础版为主——同时用轻重量重温杠铃动作记忆。每个动作记录重量+次数（  =基准数据）。" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; $r++

# 全身A
$actsA=@(
    @("推胸","坐姿推胸","哑铃/杠铃卧推","第1周固定器械为主 3片左右  RPE6-7","60-90s","轻松-中等","【呼吸】推出呼气收回。【要点】后背贴紧靠背(固械)/肩胛收紧(杠铃)。推起时胸肌向中间挤。  基准：记录了没？"),
    @("推肩","坐姿推肩","哑铃推肩","最轻片/哑铃4-6kg  RPE6","60s","轻松-中等","【呼吸】推起呼气下放吸气。【要点】下背贴紧。推到上面不锁肘。  基准。"),
    @("腿举/深蹲","坐姿腿举","杠铃深蹲/高脚杯深蹲","空机或最轻片/空杆或轻哑铃","60-90s","中等","【呼吸】下蹲吸气起身呼气。【要点】膝盖不内扣 深度控制在舒服范围。杠铃深蹲技术你学过——用空杆重温。"),
    @("髋外展","坐姿髋外展","坐姿髋外展","轻重量","45s","轻松","【要点】身体贴紧 不晃。臀外侧发力。"),
    @("下拉/划船","高位下拉","杠铃俯身划船","最轻片/空杆或轻哑铃","60-90s","轻松-中等","【呼吸】下拉呼气还原。【要点】先沉肩再拉。杠铃划船技术你学过——用轻重量重温背阔肌发力感。  基准。"),
    @("核心","仰卧举腿","仰卧举腿","自重 2组x10-15次","45s","中等","【要点】下腹带动腿上举 不靠惯性。")
)
$r=Write-TrainingBlock $ws2 $r "全身A  推力+下肢主导（每周第1练）" "胸+肩+大腿前侧+臀外侧+背阔肌+核心  第1周=3组/记录基准" "基础版=固械  进阶版=自由重量  每天自选" "38-42" $actsA

# 全身B
$actsB=@(
    @("划船","坐姿划船","杠铃俯身划船","轻-中重量  RPE6-7","60-90s","轻松-中等","【呼吸】拉向身体呼气还原。【要点】身体不后仰 肩胛往中间夹。杠铃划船：俯身45-60度 核心收紧 杠铃贴大腿前侧拉。  基准。"),
    @("下拉","高位下拉(宽握)","引体向上(辅助)/哑铃划船","最轻片/弹力带辅助或哑铃8-10kg","60-90s","中等","【要点】宽握=上背阔肌。引体向上做不了就弹力带辅助——能做几个做几个。"),
    @("腿弯举/RDL","坐姿腿弯举","杠铃/哑铃罗马尼亚硬拉","轻重量/空杆或哑铃8-12kg","60-90s","中等","【呼吸】下放吸气拉起呼气。【要点】膝盖微屈 哑铃贴大腿前侧下滑 全程直背。RDL技术你学过——轻重量重温。  基准。"),
    @("臀桥/臀推","哑铃臀桥","杠铃臀推","哑铃6-10kg/空杆或轻杠铃","60-90s","中等","【呼吸】推起呼气下放吸气。【要点】下巴微收 顶峰夹紧1-2秒。杠铃臀推：上背靠凳 杠铃放髋部。"),
    @("髋内收","坐姿髋内收","坐姿髋内收","轻重量","45s","轻松","【要点】大腿内侧发力——改善腿部线条。"),
    @("核心","平板支撑","平板支撑","自重 2组x20-30秒","30s","力竭前","【要点】手肘正下方撑地 腰不塌。  基准：能撑多少秒？")
)
$r=Write-TrainingBlock $ws2 $r "全身B  拉力+下肢后侧主导（每周第2练）" "背+背阔肌+大腿后侧+臀部+大腿内侧+核心  第1周=3组/基准" "杠铃RDL和杠铃划船技术你学过——轻重量重温" "38-42" $actsB

# 全身C
$actsC=@(
    @("推胸(上斜)","坐姿推胸","哑铃上斜卧推","比全身A轻1片/哑铃比平板轻2-4kg","60-90s","中等","综合日不冲重量——让身体在不同疲劳状态下感受动作。上斜=上胸 和平板不同。"),
    @("下拉(窄握)","高位下拉(窄握)","哑铃单臂划船","轻-中重量","60-90s","中等","窄握=下背阔肌。单臂划船=左右平衡+稳定肌参与。三种握法 背阔肌全面发展。"),
    @("腿举/硬拉","坐姿腿举","杠铃硬拉(传统)","空机/空杆或轻重量","90-120s","中等偏强","  V3新增：杠铃传统硬拉回归！你之前学过——空杆重温技术：杠铃贴小腿 臀和肩同步上升 脚后跟蹬地。  基准：这个很重要 记录了没？"),
    @("侧平举","哑铃侧平举","哑铃侧平举","最轻2-4kg","45s","中等","【要点】不耸肩——肩膀外侧发力 手臂微屈。"),
    @("手臂/有氧","绳索下压+跑步机","窄距俯卧撑+跑步机","轻重量/自重","60s","中等","下压或俯卧撑做2组。然后跑步机多走5-10分钟——本周最后一次训练 多消耗点。")
)
$r=Write-TrainingBlock $ws2 $r "全身C  综合+手臂+额外有氧（每周第3练）" "胸+背+腿(综合)+肩+手臂+额外有氧  第1周=3组/基准" "杠铃硬拉回归——轻重量重温" "42-48" $actsC

# Week 2-3 progression + barbell reintegration
$r++
SetC $ws2 $r 1 "第2周  容量积积累 & 第3周  挑战周 & 杠铃回归节奏" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作框架不变（全身A/B/C）。变化的是组数/重量/杠铃使用频率。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "路线" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "重量/RPE" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "杠铃频率" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "心态" $true 10 $white $headerBlue $false
$r++

$w23=@(
    @("第2周 容量","状态好→进阶版  状态一般→基础版","3组 4组","进阶版RPE7-8","杠铃可以出现在每次训练的进阶版了","动作记忆回来了——杠铃硬拉和划船感觉熟悉了。用第1周同样的杠铃重量 但做更多组=你在变强。"),
    @("第3周 挑战","优先进阶版——杠铃硬拉/划船尝试加重","3-4组","进阶版RPE7-9(留1-2次力竭)","杠铃硬拉/划船尝试比第1-2周重2.5-5kg","加到就加 加不到不勉强。加不到就用慢离心(3秒下放)或增加1次——同样是变强。试了就比不试强。")
)
foreach($w in $w23){
    SetC $ws2 $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws2 $r 2 $w[1] $false 10 0 0 $false
    SetC $ws2 $r 3 $w[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $w[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $w[4] $false 10 0 0 $true
    SetC $ws2 $r 6 $w[5] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=48; SB $ws2 $r 1 8; $r++
}

# Schedule
$r+=2
SetC $ws2 $r 1 "每周排班" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "正常作息" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周一=全身A / 周三=全身B / 周五=全身C / 其他日子=休息+走路8000步" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "夜班调整" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "夜班前(下午)练基础版  夜班后先睡醒再练  换班日休息或只走有氧——具体看Sheet3夜班方案" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8

# Baseline reminder
$r+=2
SetC $ws2 $r 1 "第1周结束——看你的基准数据" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++
SetC $ws2 $r 1 "周末翻一遍记录的重量和次数。你离开私教课时杠铃硬拉大概什么水平——现在什么水平——有差距很正常 不是退步 是停了一段时间需要唤醒。第4周对比 看数字涨了多少。" $false 10 0 0 $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=32

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18
$ws2.Range("C:C").ColumnWidth=18; $ws2.Range("D:D").ColumnWidth=18
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=12
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 夜班方案 + 热身 + 有氧 + 减载 + 拍照
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="夜班与热身"

SetC $ws3 1 1 "夜班方案  热身  有氧  减载周  拍照" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34

# V3: Night shift detail
$r=3
SetC $ws3 $r 1 "夜班训练+饮食方案（火车站上班专用）" $true 14 $white $softBlue $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "不需要按周一三五硬练——按你的生物钟。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws3 $r 1 "班次" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "什么时候练" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "练什么路线" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "强度" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "饮食重点" $true 10 $white $headerBlue $false; $r++

$nightDetail=@(
    @("白班（正常作息）","下班后/晚饭前","正常训练 按计划走","正常 RPE 6-8","正常三餐——碳水每餐一拳"),
    @("夜班（晚上上班）","上班前（下午4-5点）","基础版——固定器械","降一档 RPE 5-7","夜班前正常一餐：碳水+蛋白+蔬菜。夜班中(凌晨)吃轻食：水果+酸奶/蛋白粉和水——不饿就行 不要吃大餐。"),
    @("夜班后（早上下班）","先睡觉——睡到自然醒  醒后再练","基础版或直接休息","如果练就RPE 5-6","睡醒后的第一餐=你的'早餐'：2蛋+牛奶+全麦吐司。夜班后身体脱水——先喝500ml水。"),
    @("换班日（白转夜/夜转白）","休息或只做有氧","散步/拉伸","不训练","生物钟切换时身体压力最大——饮食正常吃 不刻意控碳水。不要空腹换班。")
)
foreach($nd in $nightDetail){
    SetC $ws3 $r 1 $nd[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $nd[1] $false 10 0 0 $false
    SetC $ws3 $r 3 $nd[2] $false 10 0 0 $false
    SetC $ws3 $r 4 $nd[3] $false 10 0 0 $false
    SetC $ws3 $r 5 $nd[4] $false 10 0 0 $true
    $ws3.Range("A$($r):H$($r)").RowHeight=46; SB $ws3 $r 1 8; $r++
}

$r++
SetC $ws3 $r 1 "夜班核心原则" $true 10 $softBlue $lightGray $false
SetC $ws3 $r 2 "夜班那顿(凌晨)不吃碳水大餐——身体在凌晨处理碳水能力差 容易囤积。吃蛋白质+少量水果。夜班后不要空腹睡觉——喝蛋白粉+水或吃一小块鸡胸肉 保护肌肉不分解。全天保持3L水——夜班容易脱水。" $false 10 0 0 $true
MC $ws3 $r 2 8; $ws3.Range("A$($r):H$($r)").RowHeight=36; SB $ws3 $r 1 8; $r++

# Warmup
$r+=2
SetC $ws3 $r 1 "训练前热身（5分钟）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 5 "要点" $true 9 $white $headerBlue $false; MC $ws3 $r 5 8; $r++

$wus=@(
    @("1","原地踏步+摆臂","1min","逐渐加快 微微发热"),
    @("2","肩关节绕圈","前10 后10","画大圈——推胸下拉前必做"),
    @("3","扶墙摆腿(前后+左右)","各10/侧","大腿前后的肌肉拉开"),
    @("4","自重深蹲(扶椅背)","10次","能蹲多深蹲多深"),
    @("5","第一个动作轻重量试做","1组x10次","进阶版用杠铃→先用空杆做1组 唤醒动作记忆")
)
foreach($w in $wus){
    SetC $ws3 $r 1 $w[0] $false 9 0 0 $false
    SetC $ws3 $r 2 $w[1] $true 9 $headerBlue 0 $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $w[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $w[3] $false 9 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Cardio + steps
$r+=2
SetC $ws3 $r 1 "有氧+步数" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "每天步数：8000-10000步。训练日力量后在跑步机爬坡走10-15分钟(坡度8-10% 速度4.5-5km/h)。全身C日走15-20分钟。上次你40天减7.5kg 走路和爬坡功不可没——这次一样。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=30; $r++

# Stretch
$r++
SetC $ws3 $r 1 "训练后拉伸（5分钟）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$strs=@(
    @("1-6","胸+背+大腿前+大腿后+臀+深呼吸","各30s/侧 深呼吸10次","和你上次拉的一样——这些你应该还记得")
)
SetC $ws3 $r 1 "序" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 5 "要点" $true 9 $white $headerBlue $false; MC $ws3 $r 5 8; $r++
foreach($s in $strs){
    SetC $ws3 $r 1 $s[0] $false 9 0 0 $false
    SetC $ws3 $r 2 $s[1] $true 9 $headerBlue 0 $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $s[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $s[3] $false 9 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Deload
$r+=2
SetC $ws3 $r 1 "第4周  恢复周  换动作 不碰杠铃" $true 14 $white $softPurple $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "不是做轻一点——是做不同的动作。让主力关节换受力角度=给关节放假。所有动作2组 轻重量。做完应该感觉'没练够'——这就对了。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $r++

$deloadSubs=@(
    @("杠铃卧推","跪姿俯卧撑(控制)","2组x8-10次"),
    @("杠铃划船","弹力带划船/哑铃单臂划船(轻量)","2组x10次/侧"),
    @("杠铃深蹲","自重深蹲(扶椅背)","2组x15次"),
    @("杠铃硬拉/RDL","早安式体前屈(自重)","2组x12次"),
    @("杠铃臀推","自重臀桥+弹力带","2组x15次"),
    @("坐姿推肩","弹力带/轻哑铃侧平举","2组x各10次"),
    @("有氧","跑步机坡度8% 速度4km/h","10分钟"),
    @("第4周心态","身体在偷偷变强——让它消化前三周的刺激","")
)
SetC $ws3 $r 1 "原来的" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws3 $r 4 8; $r++
foreach($ds in $deloadSubs){
    SetC $ws3 $r 1 $ds[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $ds[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $ds[2] $false 10 0 0 $true; MC $ws3 $r 4 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Photo
$r+=2
SetC $ws3 $r 1 "  拍照——今天就拍" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "穿运动内衣+紧身裤。正面(手机与肚脐同高)+侧面(手机与腰同高)+背面(手机与肩胛骨同高)。不要吸肚子不要找角度。上次你腰围小了7cm——有对比照片才知道自己变了多少。4周后同一位置同一光线同一套衣服。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=32

$ws3.Range("A:A").ColumnWidth=22; $ws3.Range("B:B").ColumnWidth=28
$ws3.Range("C:C").ColumnWidth=14; $ws3.Range("D:H").ColumnWidth=16
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 饮食+PMS+规则
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="饮食与规则"

SetC $ws4 1 1 "饮食优化  PMS情绪管理  判断规则" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

# Diet optimization V3
$r=3
SetC $ws4 $r 1 "你上次怎么吃的——先回顾 再微调" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "你40天瘦了7.5kg 饮食是主力。你饮食已经很自律了——这次不需要'控制' 需要的是'优化窗口'。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

SetC $ws4 $r 1 "  继续做的" $true 12 $accentGreen 0 $true; MC $ws4 $r 1 8; $r++

$keep=@(
    @("早餐蛋+奶","蛋白质吃够 上午不饿——保持"),
    @("午餐米饭一拳","控量不戒碳水——对减脂重要"),
    @("戒含糖饮料","热量缺口最大来源——上次做到 这次继续"),
    @("训练后蛋白粉","上次有的话继续 没有的话训练后半小时内吃含蛋白质的正餐")
)
foreach($kp in $keep){
    SetC $ws4 $r 1 $kp[0] $true 10 $accentGreen 0 $false
    SetC $ws4 $r 2 $kp[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++
}

$r++
SetC $ws4 $r 1 "  这次微调的" $true 12 $accentOrange 0 $true; MC $ws4 $r 1 8; $r++

$tweak=@(
    @("晚餐碳水换粗粮","糙米/红薯/玉米/南瓜——一拳。不饿肚子 饱腹感更强"),
    @("夜班那顿(凌晨)不吃碳水大餐","凌晨身体处理碳水能力差→改蛋白质+水果/酸奶"),
    @("训练前后窗口更精确","训练前2h：碳水30-40g+蛋白15-20g(香蕉+蛋白粉)。训练后1h内：碳水40-50g+蛋白25-30g(蛋白粉+香蕉+牛奶)"),
    @("如果上次太极端坚持不住","这次放宽。80%时间吃对 20%吃想吃的。不禁止 不放纵")
)
foreach($tw in $tweak){
    SetC $ws4 $r 1 $tw[0] $true 10 $accentOrange 0 $false
    SetC $ws4 $r 2 $tw[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++
}

# Cheat strategy
$r+=2
SetC $ws4 $r 1 "如果今天就是想吃炸鸡/火锅/蛋糕" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "可以吃。但吃之前问自己三个问题——不禁止 不放纵。" $false 10 $softPink 0 $true; MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "问题" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "答案" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$cheatQs=@(
    @("1. 我今天训练了吗？","练了→可以吃。没练→少吃一点。","用训练换放纵——公平交易"),
    @("2. 连续吃健康餐几天了？","3天以上→奖励一顿。只吃了1天→再看看。","健康餐是常态 放纵是偶尔"),
    @("3. 我吃完会不会后悔？","会→换温和选择(水果+酸奶代替蛋糕)。不会→享受它。","后悔的吃不如不吃")
)
foreach($cq in $cheatQs){
    SetC $ws4 $r 1 $cq[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $cq[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $cq[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

# Nutrition targets
$r+=2
SetC $ws4 $r 1 "营养方向（不称克数 你上次也不需要）" $true 12 $headerBlue 0 $true; MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "目标" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "怎么吃" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$nutri=@(
    @("总热量","1500-1700","比你维持少400-500——每周掉0.5-1kg"),
    @("蛋白质","100-120g","每餐有蛋白质。夜班那顿也要有(蛋白粉/鸡胸肉/鸡蛋)"),
    @("碳水","140-160g","三餐一拳。晚餐换粗粮。夜班那顿不吃碳水大餐"),
    @("脂肪","35-45g","控油。不吃油炸。坚果一小把"),
    @("饮水","2.5-3L","夜班容易脱水——上班前喝500ml 班中持续补水"),
    @("蔬菜","不限量","每餐占一半。先吃菜再吃肉最后吃饭")
)
foreach($n in $nutri){
    SetC $ws4 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $n[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=32; SB $ws4 $r 1 8; $r++
}

# PMS V2
$r+=2
SetC $ws4 $r 1 "经期前后——身体和情绪都会变" $true 14 $white $softPink $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "身体变化" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "你可能想" $true 10 $white $headerBlue $false
SetC $ws4 $r 5 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 5 8; $r++

$pmsV2=@(
    @("经前一周","体重涨1-3kg(水肿)","我又胖了？  没有 是激素","别称体重。训练降强度但去。想吃甜吃黑巧克力/水果。这周目标=维持。"),
    @("经期1-2天","身体需要休息","回到原点了？  没有 你休息了两天","休息或走路。告诉自己：2天不会让3周白费。"),
    @("经期结束","水肿排掉 体重复原","哇 之前是水肿！不是我胖了","状态回升——这周最适合冲一冲进阶版。"),
    @("全天","情绪波动 想吃甜的","我是不是不适合减脂？  每个女生都经历","不是意志力差——是激素。能做就做 做不到明天重来。")
)
foreach($pm in $pmsV2){
    SetC $ws4 $r 1 $pm[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $pm[1] $false 10 0 0 $false
    SetC $ws4 $r 3 $pm[2] $false 10 0 0 $false
    SetC $ws4 $r 5 $pm[3] $false 10 0 0 $true; MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=40; SB $ws4 $r 1 8; $r++
}

# Weight expectation
$r+=2
SetC $ws4 $r 1 "掉秤预期——和上次一样但更稳" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$exp=@(
    "上次40天平均每周约1.3kg——这次不需要那么快。第一周可能掉1-2kg(水分为主) 之后稳定0.5-1kg/周。",
    "上次体脂从43.1%降到39.5%——体脂率下降比体重慢但更真实。腰围小了7cm——这次也量腰围。",
    "某周体重没变：先看是不是经前(水肿1-3kg正常) 再看是不是吃了咸的(钠储水) 最后看腰围——腰围小了就是进步。"
)
foreach($e in $exp){
    SetC $ws4 $r 1 $e $false 10 0 0 $true; MC $ws4 $r 1 8
    $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++
}

# Rules
$r+=2
SetC $ws4 $r 1 "遇到情况" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "这种情况" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 3 8; $r++

$rules=@(
    @("训练太轻松","第2周加1组 第3周上进阶版杠铃——和上次一样"),
    @("训练后48h酸","正常——停训后重新练前两周会酸。多喝水多睡。关节疼告诉我"),
    @("体重一周没掉","先看是不是经期前后。不是的话检查饮料和夜班那顿的碳水"),
    @("夜班后太累不想练","先睡——睡醒再决定。走5000步也算运动"),
    @("某天吃多了","明天继续 不等周一。一顿不会胖——你上次就知道"),
    @("杠铃硬拉/深蹲膝盖不适","换基础版固定器械。疼超2天告诉我")
)
foreach($r3 in $rules){
    SetC $ws4 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 3 $r3[1] $false 10 0 0 $true; MC $ws4 $r 3 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=28; $ws4.Range("B:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 备用+复盘+三个选择
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="备用与复盘"

SetC $ws5 1 1 "降级方案  4周复盘  三个选择" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

# Fallback
$r=3
SetC $ws5 $r 1 "状态差/加班/夜班/来例假——降级方案" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws5 $r 1 "本来" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "最少完成" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$fallback=@(
    @("3次/周训练","减到2次 保留全身A+全身B","一周1次也比不练强"),
    @("今天该练但很累(夜班后)","热身+基础版做3个动作+拉伸","健身房待20分钟=胜利"),
    @("状态差—进阶版 基础版","降回基础版固定器械","同样完成了训练——只是速度慢一点"),
    @("完全不想出门","走20分钟+10分钟拉伸","也算运动——走路永远不丢人"),
    @("来例假前3天","只走路 不训练","身体需要休息——不是偷懒")
)
foreach($fb in $fallback){
    SetC $ws5 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $fb[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $fb[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

# Review
$r+=2
SetC $ws5 $r 1 "4周后——回头看基准数据" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "翻出第1周记录的基准数据 看数字变了多少。不要只看体重秤——你上次腰围小了7cm 比体重数字真实得多。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "可能的变化" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "上次的成绩" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$progress=@(
    @("体重","1-3kg","40天-7.5kg——4周-1~3kg合理"),
    @("腰围","2-4cm","上次-7cm——这次4周-2~4cm合理"),
    @("杠铃基准","第1周用的杠铃重量 第4周觉得轻了","杠铃硬拉和划船——动作记忆回来了"),
    @("照片","和第1天对比不一样了","照片不会骗人——上次如果没拍 这次一定要拍"),
    @("习惯","训练+饮食——感觉'回来了'","上次坚持了40天 这些习惯没走远")
)
foreach($pg in $progress){
    SetC $ws5 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pg[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $pg[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=28; SB $ws5 $r 1 8; $r++
}

# Three choices
$r+=2
SetC $ws5 $r 1 "4周后——你有三个选择" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "上次停课后没后续——这次有。根据数据一起选。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

$choices=@(
    @("A 继续减脂","还有体重目标","进入下个4周周期 杠铃加重 继续掉秤——你有了一个能一直用下去的框架"),
    @("B 进入维持","接近目标/想稳一稳","热量回到维持 训练不变。重点从'减'转向'稳'——上次减完反弹 这次学会维持"),
    @("C 塑形优化","体重满意 想改善线条","调整训练方向 增加杠铃力量比例 减少有氧——针对你想改善的部位做专项")
)
SetC $ws5 $r 1 "选择" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "什么时候" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "路径" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++
foreach($ch in $choices){
    SetC $ws5 $r 1 $ch[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $ch[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $ch[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=32; SB $ws5 $r 1 8; $r++
}

# Final compare
$r+=2
SetC $ws5 $r 1 "和上次比——这次你有了什么" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$compare=@(
    "上次79.9kg起步→这次72kg。起点低了8kg。",
    "上次从头学器械→这次杠铃回归——你会的动作比上次多。",
    "上次40天7.5kg→这次不需要那么快。稳就行。",
    "上次停了就全停→这次有完整的节奏：唤醒→容量→挑战→恢复+三个选择。",
    "上次你证明了你能瘦——这次你要学会的是：怎么把瘦保持住。",
    "",
    "你不是重新开始——你是继续。起点更低 基础更好 节奏更稳。加油。"
)
foreach($c in $compare){
    SetC $ws5 $r 1 $c $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=20
    if($c -eq ""){$ws5.Range("A$($r):H$($r)").RowHeight=10}
    $r++
}

$ws5.Range("A:A").ColumnWidth=26; $ws5.Range("B:H").ColumnWidth=20
Write-Host "Sheet 5 done"

# Save
$savePath="D:\Codex\蜡笔小杨_减脂方案_V3.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
