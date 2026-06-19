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
    SetC $ws $ref 3 "组数" $true 10 $white $headerBlue $false
    SetC $ws $ref 4 "次数" $true 10 $white $headerBlue $false
    SetC $ws $ref 5 "重量/负载" $true 10 $white $headerBlue $false
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
# SHEET 0:  从这里开始
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "蜡笔小杨  减脂塑形方案 V2" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 72kg  |  继续上次的旅程" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
# V2: 叙事重构——从"重新开始"改为"继续上次的旅程"
SetC $ws0 $row 1 "  你不是重新开始——你是继续" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "半年前你40天瘦了7.5kg。那不是运气——是你做对了一系列事情。体脂从43.1%降到39.5% 腰围小了7cm。现在你72kg 跟当时最低点差不多——说明你没反弹 你的身体记住了那个状态。这次不是重新开始 是继续。上次你学会了怎么瘦下来。这次你要学会的是——怎么把瘦下来的状态保持住。" $false 10 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=48; $row++

# V2: 上次做对了什么 + 这次避开什么
$row+=2
SetC $ws0 $row 1 "  上次你做对了什么  继续" $true 14 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$didRight=@(
    @("  饮食控制","你当时怎么吃的——这次照做。你不需要重新学 只需要重新做。"),
    @("  规律训练","你当时一周练好几次——身体记得那些动作。这次唤醒它。"),
    @("  坚持记录","你当时拍了照 量了腰围——所以才看到7cm的变化。这次也拍。")
)
foreach($dr in $didRight){
    SetC $ws0 $row 1 $dr[0] $true 10 $accentGreen 0 $false
    SetC $ws0 $row 2 $dr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++
}

$row++
SetC $ws0 $row 1 "  上次你踩过的坑  这次避开" $true 14 $white $accentRed $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentRed
$row++

$didWrong=@(
    @("  太激进","上次40天瘦7.5kg速度偏快——这次慢一点 稳一点 每周0.5-1kg就行"),
    @("  恢复不够","上次可能训练频率偏高 恢复跟不上——这次每周3次 留足恢复"),
    @("  停了就全停","上次没有过渡方案——这次有减载周 有降级方案 有4周后的路线图")
)
foreach($dw in $didWrong){
    SetC $ws0 $row 1 $dw[0] $true 10 $accentRed 0 $false
    SetC $ws0 $row 2 $dw[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++
}

# V2: 四周总览——每周围绕"继续旅程"叙事
$row+=2
SetC $ws0 $row 1 "  四周总览——继续的节奏" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "做什么" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周目标" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","唤醒+基准","3次/周 全身","记录现在的起点","每个动作记录重量+次数——找到'现在的位置'。第4周对比 看涨了多少。"),
    @("2","容量积累","3次/周 全身","组数+1 积累代谢","身体开始适应——做第1周同样的动作 但容量更大。你会发现比第1周轻松。"),
    @("3","挑战周","3次/周 全身","试试比上周多一点","尝试比前两周重1-2片。加到就加 加不到就用慢离心或增加次数——不勉强 不评判。"),
    @("4","恢复+复盘","3次/周 减载","让身体消化 看变化","换动作 轻重量。拍对比照 量体重腰围。回头看第1周的基准数据——你会看到进步。")
)
foreach($o4i in $o4){
    SetC $ws0 $row 1 $o4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $o4i[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $o4i[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $o4i[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $o4i[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=30; SB $ws0 $row 1 5; $row++
}

# V2: 这次和上次有什么不同
$row+=2
SetC $ws0 $row 1 "  这次和上次有什么不同" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$diff=@(
    "上次你有教练手把手——这次你靠自己。但这个方案已经帮你铺好了路：每周三天 每次做什么 几组 什么重量——照着走就行。",
    "上次你从79.9kg开始——这次从72kg。起点低了8kg。你不用从那么深的地方开始爬。",
    "上次你40天平均每周掉1.3kg——这次不需要那么快。稳就是快。",
    "这次有减载周——上次可能一路冲到停。这次第4周让你身体消化前三周的刺激 然后带着更强的状态进入下一个周期。",
    "你不需要重新证明什么——你只需要重新开始做。"
)
foreach($d in $diff){
    SetC $ws0 $row 1 $d $false 10 0 0 $true; MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+设计思路"," Sheet 方案概览"),
    @("每次训练练什么（第1-4周）"," Sheet 训练计划"),
    @("热身+有氧+拉伸+减载+拍照"," Sheet 热身有氧减载"),
    @("饮食怎么吃+PMS+掉秤预期"," Sheet 饮食与规则"),
    @("状态差/来例假/加班 怎么办"," Sheet 备用与复盘"),
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

SetC $ws 2 1 "女  25岁以下  160cm / 72kg  BMI 28.1  体脂35%+  40天私教基础(曾减7.5kg/腰围-7cm)  继续旅程" $false 11 $gold $darkBg $true
MC $ws 2 1 8; $ws.Range("A2:H2").Font.Color=$white; $ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 8

$ov=@(
    @("训练频率","每周3练 全身训练（周一/三/五）","每次60分钟以上（含热身有氧拉伸）"),
    @("训练方式","固定器械为主 + 轻哑铃辅助","你用过这些器械——不需要重新学"),
    @("周期结构","唤醒+基准  容量积累  挑战周  恢复+复盘","每周围绕'继续上次的旅程'推进"),
    @("方案周期","4周为起点","4周后根据数据+感受 选择继续减脂/维持/塑形"),
    @("方案特点","回归训练专用  重建信心  上次经验复用  4周后路线图")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 5
    SetC $ws $r 6 $o[2] $false 10 0 0 $true; MC $ws $r 6 8
    SB $ws $r 1 8; $r++
}

# V2: 为什么这样设计——围绕"继续旅程"
$r++
SetC $ws $r 1 "为什么这样设计——继续上次的旅程" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$design=@(
    "你不是纯新手——你练过40天 瘦过7.5kg。你的身体记得每个器械怎么用 你的肌肉记得发力感。第一周不是'学习'——是'唤醒'。",
    "第1周加了一个基准测试：每个动作记录你现在的重量+次数。不是评判你现在多强——是给第4周一个参照。",
    "第2周加容量（组数+1）——和第1周同样的动作 但容量更大。你会发现比第1周轻松——这就是身体在适应。",
    "第3周叫'挑战周'不是'强度周'——可以尝试加重 加不到就用慢离心或更多次数。不勉强 不评判。",
    "第4周减载——让身体消化前三周的刺激。然后回头看第1周的基准 你会看到你比4周前强了多少。",
    "这次和上次最大的不同：上次可能冲到停 这次有完整的节奏——唤醒  累积  挑战  恢复。这个节奏你能做一辈子。"
)
foreach($d in $design){
    SetC $ws $r 1 $d $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws $r 1 "四周节奏一览" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

SetC $ws $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws $r 3 "训练日" $true 10 $white $headerBlue $false
SetC $ws $r 4 "主题" $true 10 $white $headerBlue $false
SetC $ws $r 5 "组数/重量" $true 10 $white $headerBlue $false
SetC $ws $r 7 "本周目标" $true 10 $white $headerBlue $false; $r++

$w4=@(
    @("1","唤醒+基准","3次","全身A/B/C","3组 中等 全记录","找到现在的位置——第4周回来看涨了多少"),
    @("2","容量积累","3次","全身A/B/C","4组 中等 次数上限","身体开始适应——比第1周轻松就是进步"),
    @("3","挑战周","3次","全身A/B/C","尝试比上周重1-2片","加到就加 加不到不勉强——试了就比不试强"),
    @("4","恢复+复盘","3次","全身A/B/C(轻量)","2-3组 换动作 轻重量","拍对比照 量腰围 看基准变化")
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
SetC $ws $r 1 "三个核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$pr=@(
    @("1","你不是从零开始","你练过40天 瘦过7.5kg。身体记得那些动作 肌肉记得发力感。第一周就是唤醒。"),
    @("2","稳就是快","上次40天平均每周1.3kg——这次不需要那么快。0.5-1kg/周就很好。慢一点 稳一点 不反弹。"),
    @("3","上次的经验就是这次的武器","你上次做得好的——继续。上次踩过的坑——这次避开。你不是从头摸索 是带着经验升级。")
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
# SHEET 2: 训练计划（V2: 唤醒+基准 / 容量 / 挑战 / 减载）
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划  第1-4周" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "每周3练（周一/三/五）。全身A=推力+下肢 / 全身B=拉力+下肢后侧 / 全身C=综合+手臂+额外有氧。第4周减载换动作(见Sheet3)。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# Week 1:  唤醒+基准
$r=4
SetC $ws2 $r 1 "第1周  唤醒+基准测试  重建动作记忆" $true 14 $white $accentGreen $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws2 $r 1 "第1周目标：每个动作做3组 12-15次。选一个你'感觉轻松'的重量。重点是——每组记录你用了多重、做了几次、做完什么感觉。周末你会得到一份'现在的你'的基准数据。" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28; $r++

# 全身A W1
$actsA=@(
    @("坐姿推胸","3组","12-15次","选一个做完15次最后3次有点累的重量","60秒","轻松-中等","【呼吸】推出呼气 收回吸气。【要点】后背贴紧靠背 推到前面不锁死手肘。  基准记录：你用了多重？做了几次？"),
    @("坐姿推肩","3组","12-15次","最轻片起步——找到三角肌发力的感觉","60秒","轻松-中等","【呼吸】推起呼气 下放吸气。【要点】下背贴紧靠背 推到上面不锁肘。  基准记录：和上次比 推肩的重量有变化吗？"),
    @("坐姿腿举","3组","12-15次","空机或最轻片——大腿和臀部一起用力","60秒","中等","【呼吸】推出呼气 收回吸气。【要点】脚踩实与肩同宽 膝盖不完全蹬直。  基准记录：腿举是你上次的强项 现在什么重量？"),
    @("坐姿髋外展","3组","15-20次","轻重量——臀外侧有感就行","45秒","轻松","【呼吸】打开呼气 收回吸气。【要点】身体贴紧靠背不晃。感受臀外侧发力。"),
    @("高位下拉","3组","12-15次","最轻片——感受背阔肌发力","60秒","轻松-中等","【呼吸】拉下呼气 还原吸气。【要点】先沉肩膀再用手臂拉。杆拉到锁骨高度。  基准记录：现在是几片？"),
    @("仰卧举腿","2组","10-15次","自重","45秒","中等","【要点】用下腹带动腿上举 不靠惯性摆。")
)
$r=Write-TrainingBlock $ws2 $r "全身A  推力+下肢主导（每周第1练）" "胸+肩+大腿前侧+臀外侧+背阔肌+核心  第1周=3组/基准" "记录每个动作的重量和次数" "38-42" $actsA

# 全身B W1
$actsB=@(
    @("坐姿划船","3组","12-15次","轻-中重量——肩胛骨往中间夹","60秒","轻松-中等","【呼吸】拉向身体呼气 还原吸气。【要点】身体不要后仰——肩胛骨往中间夹 拉到腹部。  基准：现在是几片？"),
    @("高位下拉(宽握)","3组","12-15次","比全身A可以重1片——背已经激活了","60秒","中等","【发力感】宽握更偏向上背阔肌——和全身A窄握角度不同。  记录：宽握vs窄握 哪个感觉更好？"),
    @("坐姿腿弯举","3组","12-15次","轻重量——大腿后侧有感","60秒","中等","【呼吸】勾起呼气 还原吸气。【要点】大腿后侧发力勾下。改善大腿后侧线条。  基准：现在是几片？"),
    @("哑铃臀桥","3组","12-15次","双手持哑铃6-10kg放髋部","60秒","中等","【呼吸】推起呼气 下放吸气。【要点】下巴微收 顶峰夹紧屁股1-2秒。  记录：用了多重的哑铃？"),
    @("坐姿髋内收","3组","15-20次","轻重量","45秒","轻松","【呼吸】内收呼气 打开吸气。【要点】大腿内侧发力——改善大腿内侧线条。"),
    @("平板支撑","2组","20-30秒","自重","30秒","力竭前停","【要点】手肘正下方撑地 腰不塌。能做多久做多久。  基准：这次能撑多少秒？")
)
$r=Write-TrainingBlock $ws2 $r "全身B  拉力+下肢后侧主导（每周第2练）" "上背+背阔肌+大腿后侧+臀部+大腿内侧+核心  第1周=3组/基准" "记录每个动作的重量和次数" "38-42" $actsB

# 全身C W1
$actsC=@(
    @("坐姿推胸","3组","12-15次","比全身A轻1片——综合日不冲重量","60秒","中等","综合日的目的不是突破——是让身体在不同疲劳状态下感受同一个动作。  记录：和全身A比 感觉有什么不同？"),
    @("高位下拉(窄握)","3组","12-15次","轻-中重量","60秒","中等","窄握更偏向下背阔肌——和全身B宽握互补。三种握法 背阔肌全面发展。"),
    @("坐姿腿举","3组","12-15次","和全身A同样重量","60秒","中等","两次下肢训练间隔至少48h——刚好恢复到可以再练。"),
    @("哑铃侧平举","3组","12-15次","最轻哑铃2-4kg","45秒","中等","【要点】不要耸肩——用肩膀外侧发力 手臂微屈。肩部线条的收尾动作。"),
    @("绳索下压(龙门架)","2组","12-15次","轻重量","45秒","中等","【要点】大臂固定身体两侧 压到底微微外旋手腕。手臂后侧紧致。"),
    @("有氧加时","跑步机爬坡走","15-20分钟","坡度8-10% 速度4.5-5km/h","—","中等","全身C有氧比平时多5-10分钟——本周最后一次训练 多消耗一点 周末恢复。")
)
$r=Write-TrainingBlock $ws2 $r "全身C  综合+手臂（每周第3练）" "胸+背+腿(综合)+肩+手臂+额外有氧  第1周=3组/基准" "记录每个动作的感觉" "42-48" $actsC

# Week 2-3 progression notes
$r++
SetC $ws2 $r 1 "第2周  容量积累 & 第3周  挑战周" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作框架不变（同样的全身A/B/C）。变化的是组数、重量、挑战方式。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "重量" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "感觉" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "本周心态" $true 10 $white $headerBlue $false; $r++

$w23=@(
    @("第2周 容量","3组  4组","12-15次(取上限)","和第一周一样——先不加重量","缩短10秒","比第1周轻松——就是进步","身体在适应。同样重量做更多组=你在变强"),
    @("第3周 挑战","3-4组","10-12次","试试比第1-2周重1-2片","75-90秒","有点挑战但不力竭","加到了就加 加不到不勉强。试了就比不试强。 如果加不上去：用慢离心(3秒下放)或增加1次——同样是在变强")
)
foreach($w in $w23){
    SetC $ws $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $w[1] $false 10 0 0 $false
    SetC $ws $r 3 $w[2] $false 10 0 0 $false
    SetC $ws $r 4 $w[3] $false 10 0 0 $false
    SetC $ws $r 5 $w[4] $false 10 0 0 $false
    SetC $ws $r 6 $w[5] $false 10 0 0 $false
    SetC $ws $r 7 $w[6] $false 10 0 0 $true
    $ws.Range("A$($r):H$($r)").RowHeight=42; SB $ws $r 1 8; $r++
}

# Schedule
$r++
SetC $ws2 $r 1 "每周排班" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "推荐" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周一=全身A / 周三=全身B / 周五=全身C / 其他日子=休息+走路8000步" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "备选" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周二=全身A / 周四=全身B / 周六=全身C——如果周一有事的话" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8

$r+=2
SetC $ws2 $r 1 "第1周结束时——看你的基准数据" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

$baseline=@(
    "周末回头翻一遍这周记录的重量和次数。你离开私教课时大概什么水平——现在什么水平——中间有差距很正常。",
    "不是'退步'——是'停了一段时间 身体需要唤醒'。",
    "第4周再做一次同样的测试。你会看到数字的变化——那个变化比任何鼓励都有力量。"
)
foreach($bl in $baseline){
    SetC $ws2 $r 1 $bl $false 10 0 0 $true; MC $ws2 $r 1 8
    $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=20; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=56
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 热身+有氧+拉伸+减载+拍照
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="热身有氧减载"

SetC $ws3 1 1 "热身  有氧+步数  拉伸  减载周  拍照" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34

# Warmup
$r=3
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
    @("4","自重深蹲(扶椅背)","10次","能蹲多深蹲多深 找大腿发力的感觉"),
    @("5","第一个动作空机/最轻重量试做","1组x10次","让身体知道接下来要做什么")
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

SetC $ws3 $r 1 "每天步数" $true 10 $headerBlue $lightGray $false
SetC $ws3 $r 2 "8000-10000步——和上次减脂一样 走路是基础消耗。训练日力量后在跑步机爬坡走10-15分钟(坡度8-10% 速度4.5-5km/h)。全身C日走15-20分钟。" $false 10 0 0 $true
MC $ws3 $r 2 8; $ws3.Range("A$($r):H$($r)").RowHeight=28; SB $ws3 $r 1 8; $r++

SetC $ws3 $r 1 "上次你做到了" $true 10 $headerBlue $lightGray $false
SetC $ws3 $r 2 "上次40天减7.5kg 每天的步数和有氧你是有做到的 这次一样。" $false 10 0 0 $true
MC $ws3 $r 2 8; $ws3.Range("A$($r):H$($r)").RowHeight=22; SB $ws3 $r 1 8; $r++

# Stretch
$r+=2
SetC $ws3 $r 1 "训练后拉伸（5分钟）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "序" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 5 "要点" $true 9 $white $headerBlue $false; MC $ws3 $r 5 8; $r++

$strs=@(
    @("1","胸大肌拉伸(扶墙)","30s/侧","推胸后——手臂张开扶墙 身体反方向转"),
    @("2","背阔肌拉伸(扶墙侧屈)","30s/侧","下拉/划船后——手扶墙 身体侧弯"),
    @("3","大腿前侧拉伸(站姿脚跟贴臀)","30s/侧","扶墙——腿举后"),
    @("4","大腿后侧拉伸(坐姿直腿前屈)","30s","腿弯举后——坐地上 伸直腿 手够脚尖"),
    @("5","臀拉伸(跷二郎腿抱膝)","30s/侧","臀桥/髋外展后"),
    @("6","深呼吸","10次慢吸慢呼","让心跳慢下来——告诉身体训练结束 可以恢复了")
)
foreach($s in $strs){
    SetC $ws3 $r 1 $s[0] $false 9 0 0 $false
    SetC $ws3 $r 2 $s[1] $true 9 $headerBlue 0 $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $s[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $s[3] $false 9 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=28; SB $ws3 $r 1 8; $r++
}

# Deload
$r+=2
SetC $ws3 $r 1 "第4周  恢复周  换动作 让身体消化" $true 14 $white $softPurple $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "核心原则：不是做轻一点——是做不同的动作。让主力关节换一个受力角度=给关节放假。所有动作2组 轻重量 RPE4-5。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $r++

SetC $ws3 $r 1 "原来的" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws3 $r 4 8; $r++

$deloadSubs=@(
    @("坐姿推胸","跪姿俯卧撑(控制节奏)","2组x8-10次 慢节奏"),
    @("坐姿推肩","弹力带/轻哑铃侧平举+前平举","2组x各10次 最轻"),
    @("高位下拉","弹力带高位下拉 单臂","2组x10次/侧"),
    @("坐姿划船","哑铃单臂划船 最轻重量","2组x10次/侧"),
    @("坐姿腿举","自重深蹲 扶椅背","2组x15次"),
    @("坐姿腿弯举","早安式体前屈 自重","2组x12次"),
    @("哑铃臀桥","自重臀桥+弹力带","2组x15次"),
    @("髋外展/内收","站姿弹力带侧向走","2组x15步/侧"),
    @("第4周有氧","跑步机坡度8% 速度4km/h","10分钟——跟着力量一起减载"),
    @("第4周感受","做完应该感觉没练够","——这就是理想状态。身体在偷偷变强。")
)
foreach($ds in $deloadSubs){
    SetC $ws3 $r 1 $ds[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $ds[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $ds[2] $false 10 0 0 $true; MC $ws3 $r 4 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Photo
$r+=2
SetC $ws3 $r 1 "  拍照——今天就拍（和上次一样 但这次更标准）" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "穿运动内衣+紧身裤。正面(手机与肚脐同高)+侧面(手机与腰同高)+背面(手机与肩胛骨同高)。不要吸肚子不要找角度。上次你腰围小了7cm——有了对比照片才知道自己变了多少。4周后同一位置同一光线同一套衣服。" $false 10 $headerBlue $lightGray $true
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

# V2: 饮食——先回顾上次
$r=3
SetC $ws4 $r 1 "你上次怎么吃的？——先回顾 再优化" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "你40天瘦了7.5kg——说明你的饮食执行是有章法的。下面这些你上次应该大部分都做到了。这次不需要大改——只需要微调。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

SetC $ws4 $r 1 "上次做到的  继续" $true 12 $accentGreen 0 $true; MC $ws4 $r 1 8; $r++

$keep=@(
    @("  早餐吃蛋和牛奶","蛋白质吃够 上午不饿——这个习惯保持"),
    @("  午餐米饭一拳","控量但不戒碳水——这个对女性减脂很重要"),
    @("  戒含糖饮料","这是热量缺口最大的来源——上次做到了 这次继续"),
    @("  晚餐控碳水","上次晚餐怎么吃的？差不多保持就行")
)
foreach($kp in $keep){
    SetC $ws4 $r 1 $kp[0] $true 10 $accentGreen 0 $false
    SetC $ws4 $r 2 $kp[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++
}

$r++
SetC $ws4 $r 1 "上次没做到/太极端  这次微调" $true 12 $accentOrange 0 $true; MC $ws4 $r 1 8; $r++

$tweak=@(
    @("  晚餐吃粗粮","如果上次晚餐碳水控太狠会饿 这次晚餐换成粗粮(糙米/红薯/玉米/南瓜)——一拳 不饿肚子"),
    @("  训练后喝蛋白粉吗？","上次有吗？有的话继续。没有的话训练后半小时内吃含蛋白质的正餐"),
    @("  如果上次太极端导致后期坚持不住","这次放宽一点。80%时间吃对 20%时间吃想吃的。不禁止 不放纵。")
)
foreach($tw in $tweak){
    SetC $ws4 $r 1 $tw[0] $true 10 $accentOrange 0 $false
    SetC $ws4 $r 2 $tw[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++
}

# Today cheat strategy V2
$r+=2
SetC $ws4 $r 1 "如果今天就是想吃炸鸡/火锅/蛋糕" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "可以吃。但吃之前问自己三个问题：" $false 11 $softPink 0 $true; MC $ws4 $r 1 8; $r++

$cheatQs=@(
    @("1","我今天训练了吗？","练了——可以吃。没练——少吃一点。"),
    @("2","我连续吃了几天健康餐了？","3天以上——奖励一顿没问题。只吃了1天——再看看。"),
    @("3","我吃完会不会后悔？","如果会后悔——换一个更温和的选择（比如水果+酸奶代替蛋糕）。"),
    @("核心","不禁止 不放纵","80%时间吃对 20%时间吃想吃的。上次可能太严格了——这次换个方式。")
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
    @("总热量","1500-1700","比你维持少400-500——每周掉0.5-1kg。上次你大概也是这个缺口。"),
    @("蛋白质","100-120g","每餐有蛋白质：早2蛋+奶 午晚各一份肉/鱼/豆腐。练后蛋白粉。"),
    @("碳水","140-160g","三餐每餐一拳。晚餐换粗粮。上次你怎么吃碳水的？差不多也这样。"),
    @("脂肪","35-45g","炒菜控油。不吃油炸。坚果一小把。"),
    @("饮水","2.5-3L","和你上次一样——水喝够代谢才运转。戒含糖饮料。"),
    @("蔬菜","不限量","每餐占一半。先吃菜再吃肉最后吃饭——这个顺序你上次应该也用过。")
)
foreach($n in $nutri){
    SetC $ws4 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $n[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=34; SB $ws4 $r 1 8; $r++
}

# V2: PMS emotional management
$r+=2
SetC $ws4 $r 1 "经期前后——不只是体重变 情绪也会变  V2重点" $true 14 $white $softPink $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

SetC $ws4 $r 1 "提前告诉你这些——到时候就不会自我怀疑。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=20; $r++

$pmsV2=@(
    @("经前一周","体重涨1-3kg  正常水肿不是胖了","我是不是又胖了？  没有 是激素","别称体重或少称。训练降强度但尽量去。想吃甜的吃黑巧克力/水果。这周目标不是进步——是维持。"),
    @("经期1-2天","身体在流血 需要休息","我又回到原点了？  没有 你只是休息了两天","休息或只走路。告诉自己：休息2天不会让之前3周的努力白费。"),
    @("经期结束后","体重回落1-3kg 水肿排掉了","哇 原来之前是水肿——不是我胖了","状态回升的一周——最适合'冲一冲'。训练可以稍微加把劲。"),
    @("全天都适用","情绪波动 不想练 想吃甜的——都是激素 不是意志力差","我是不是不适合减脂？  不 每个女生都经历这个","给自己一个抱抱。能做就做 做不到今天过了明天重来。")
)
SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "身体变化" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "你可能会想" $true 10 $white $headerBlue $false
SetC $ws4 $r 5 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 5 8; $r++
foreach($pm in $pmsV2){
    SetC $ws4 $r 1 $pm[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $pm[1] $false 10 0 0 $false
    SetC $ws4 $r 3 $pm[2] $false 10 0 0 $false
    SetC $ws4 $r 5 $pm[3] $false 10 0 0 $true; MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=42; SB $ws4 $r 1 8; $r++
}

# Weight expectation V2
$r+=2
SetC $ws4 $r 1 "掉秤预期——和上次一样但这次更稳" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$exp=@(
    "你上次40天平均每周约1.3kg——速度快。这次不需要那么快。第一周可能掉1-2kg（水分为主）之后稳定在0.5-1kg/周。",
    "上次你体脂从43.1%降到39.5%——体脂率下降比体重慢 但更真实。腰围小了7cm——这次也量腰围 那个数字比体重秤诚实。",
    "如果某周体重没变：先看是不是经前（水肿1-3kg正常）再看是不是吃了咸的（钠储水）最后看腰围——腰围小了就是进步。"
)
foreach($e in $exp){
    SetC $ws4 $r 1 $e $false 10 0 0 $true; MC $ws4 $r 1 8
    $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++
}

# Rules
$r+=2
SetC $ws4 $r 1 "遇到这些情况" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "这种情况" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 3 8; $r++

$rules=@(
    @("训练太轻松","第2周加1组 第3周试试加重——和上次一样"),
    @("训练后48h酸","正常——停了一段时间重新练 前两周就会酸。多喝水多睡。关节疼告诉我。"),
    @("体重一周没掉","先看是不是经期前后。不是的话检查饮料和晚餐。"),
    @("今天完全不练","走5000步+拉伸——做了就算。别连续跳2次——你上次坚持了40天 这次也可以。"),
    @("某天吃多了","明天继续 不等周一。一顿吃多不会胖——你上次就知道了。"),
    @("膝盖/腰/关节疼","立即停那个动作。疼超2天告诉我。")
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
SetC $ws5 $r 1 "状态差/加班/来例假——降级方案" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws5 $r 1 "本来" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "最少完成" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$fallback=@(
    @("3次/周训练","减到2次 保留全身A+全身B","一周1次也比不练强"),
    @("今天该练但很累","热身+最轻重量做3个动作+拉伸","健身房待20分钟=胜利"),
    @("完全不想出门","走20分钟+10分钟拉伸","也算运动——走路永远不丢人"),
    @("来例假前3天","只走路 不训练","身体需要休息——不是偷懒"),
    @("经前想吃甜的","吃！选黑巧克力/水果/酸奶","不是意志力差——是激素")
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

SetC $ws5 $r 1 "翻出第1周记录的基准数据 看看数字变了多少。不要只看体重秤——你上次腰围小了7cm 比体重数字真实得多。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "可能的变化" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "上次的成绩" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$progress=@(
    @("体重","掉了1-3kg","40天-7.5kg——这次4周-1~3kg很合理"),
    @("腰围","小了2-4cm","上次-7cm——这次4周-2~4cm是合理期待"),
    @("力量基准","第1周用的重量 第4周觉得太轻了","上次40天你肯定经历了这个过程——这次一样"),
    @("照片","和第1天对比不一样了","上次没拍对比照？这次一定要拍——照片不会骗人"),
    @("习惯","训练+饮食——感觉'回来了'","上次你坚持了40天 积累了这些习惯 它们没走远")
)
foreach($pg in $progress){
    SetC $ws5 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pg[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $pg[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=30; SB $ws5 $r 1 8; $r++
}

# V2: 4周后三个选择
$r+=2
SetC $ws5 $r 1 "4周后——你有三个选择" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "上次停课后没有后续——这次不一样。4周后我们根据数据一起选。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

$choices=@(
    @("选择A：继续减脂","如果还有体重目标（比如你想从72kg继续往下走）","进入下一个4周周期。重量递增 继续掉秤。你现在有了一个能一直用下去的框架。"),
    @("选择B：进入维持","如果接近目标体重 或者想稳一稳再继续","热量回到维持水平 训练频率不变。重点从'减'转向'稳'。上次减完反弹——这次学会维持。"),
    @("选择C：塑形优化","如果体重满意了 想改善线条和体型","调整训练方向 增加力量比例 减少有氧。针对你想改善的部位做专项调整。")
)
SetC $ws5 $r 1 "选择" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "什么时候选" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "路径" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++
foreach($ch in $choices){
    SetC $ws5 $r 1 $ch[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $ch[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $ch[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=36; SB $ws5 $r 1 8; $r++
}

# V2: 和上次的总结对比
$r+=2
SetC $ws5 $r 1 "和上次比——这次你有了什么" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$compare=@(
    "上次你从79.9kg开始——这次从72kg。起点低了8kg。",
    "上次你从头学器械——这次你都会用了。唤醒就够。",
    "上次你40天瘦了7.5kg——这次不用那么激进。稳就行。",
    "上次停课后没有后续——这次有完整的4周节奏 + 减载周 + 三个选择。你不是在冲刺 你是在建立一个能跑一辈子的系统。",
    "",
    "上次40天瘦了7.5kg——你做到过一次。这次起点更低 基础更好 节奏更稳。你不是重新开始——你是继续。"
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
$savePath="D:\Codex\蜡笔小杨_减脂方案_V2.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
