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
    SetC $ws $ref 5 "重量建议" $true 10 $white $headerBlue $false
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

SetC $ws0 1 1 "蜡笔小杨  减脂塑形方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 72kg  |  回归训练  继续减脂" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你不是从零开始" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "半年前 你40天瘦了7.5kg——你知道自己能做到。体脂从43.1%降到39.5% 腰围小了7cm。现在体重72kg 跟当时最低点差不多 说明你没反弹。你只是停了一段时间 不是退回了原点。" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=42; $row++

$row+=2
SetC $ws0 $row 1 "  四周总览——回归的节奏" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "训练" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周目标" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","唤醒","3次/周 全身训练","重建动作记忆","每个动作重新找到发力感——你练过 会比别人快很多"),
    @("2","进入节奏","3次/周 全身训练","容量提升","组数+1 次数取上限——和第1周比 身体开始适应了"),
    @("3","加速","3次/周 全身训练","重量递增","每个动作比前两周重1-2片——找到'有点挑战'的感觉"),
    @("4","看到变化","3次/周 全身训练","减载周","主动恢复——让身体消化前三周的刺激。拍对比照 量体重腰围")
)
foreach($o4i in $o4){
    SetC $ws0 $row 1 $o4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $o4i[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $o4i[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $o4i[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $o4i[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=30; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  这次和上次有什么不同" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$diff=@(
    "上次你有教练在旁边手把手带——一周可能去了四五次。这次你只有固定器械和自己。",
    "所以这次方案的核心逻辑不一样：每周3次——够刺激但不过度疲劳。固定器械为主——你知道怎么用。",
    "上次40天瘦了7.5kg——你怎么吃的还记得吗？这次饮食部分我把重点帮你捡回来。",
    "核心原则：你不是从零开始。你是重新出发——比别人快得多。"
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
    @("热身+有氧+拉伸+拍照"," Sheet 热身有氧拉伸"),
    @("饮食怎么吃+掉秤预期"," Sheet 饮食与规则"),
    @("状态差/来例假/加班了 怎么办"," Sheet 备用与复盘"),
    @("4周后怎么看效果"," Sheet 备用与复盘")
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

SetC $ws 2 1 "女  25岁以下  160cm / 72kg  BMI 28.1  体脂35%+  有40天私教基础  回归训练  目标：减脂" $false 11 $gold $darkBg $true
MC $ws 2 1 8; $ws.Range("A2:H2").Font.Color=$white; $ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 8

$ov=@(
    @("训练频率","每周3练 全身训练（周一/三/五）","每次60分钟以上（含热身有氧拉伸）"),
    @("训练方式","固定器械为主 + 轻哑铃辅助","你用过这些器械——不需要重新学"),
    @("周期结构","第1周=唤醒(重建动作记忆)","第2周=容量(组数+1) 第3周=强度(重量+1-2片) 第4周=减载(主动恢复)"),
    @("方案周期","4周为起点","4周后根据体重+腰围+体能调整"),
    @("方案特点","回归训练专用  重建信心  固定器械安全感  饮食捡回重点")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 5
    SetC $ws $r 6 $o[2] $false 10 0 0 $true; MC $ws $r 6 8
    SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "为什么这样设计" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$design=@(
    "你和纯新手不一样——你练过40天，知道坐姿推胸怎么用、知道高位下拉什么感觉。第一周不需要'学器械'——需要的是'唤醒身体记忆'。",
    "所以你从第一周就开始训练。每周3次全身训练——和之前私教课的节奏不同（那时可能一周四五次），但3次足够刺激且不过度疲劳。",
    "第1周用中等重量重建动作模式。第2周加组数（容量）。第3周加重量（强度）。第4周减载——让身体恢复。",
    "整套方案用固定器械为主——因为你现在没有教练在旁边保护，固定器械轨道固定、安全。你知道怎么用。",
    "饮食部分：你上次40天瘦7.5kg，说明你的饮食执行过而且有效。这次我把核心原则捡回来——不需要重新学，只需要重新做。"
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
    @("1","唤醒","3次","全身A/B/C","3组 中等重量","重建动作记忆——每个动作找到发力感"),
    @("2","容量","3次","全身A/B/C","4组 中等重量 次数上限","身体开始适应——组数+1 代谢压力增加"),
    @("3","强度","3次","全身A/B/C","3-4组 比前两周重1-2片","找到'有点挑战'的感觉——但不力竭"),
    @("4","减载","3次","全身A/B/C(轻量)","2-3组 轻重量 换动作","主动恢复——拍对比照 量体重腰围")
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
    @("1","你不是从零开始","你练过40天 瘦过7.5kg。身体记得那些动作——第一周就是唤醒记忆 比别人快得多。"),
    @("2","每周3次 够用","和之前私教课一周四五次不同——3次全身训练刚好够刺激但不过度疲劳。你还有工作 还有生活。"),
    @("3","饮食是上次成功的秘密","你上次40天瘦7.5kg 训练当然有功 但饮食执行才是大头。这次把饮食重点捡回来。")
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
# SHEET 2: 训练计划
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划  第1-4周" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "每周3练（如周一/三/五）。每天不同侧重——A推力 B拉力 C综合。第1-3周用固定器械+轻哑铃 第4周减载换动作。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# 全身A
$r=4
SetC $ws2 $r 1 "全身A  推力+下肢主导  第1周" $true 14 $white $accentGreen $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws2 $r 1 "第1周=3组 中等重量 重建动作记忆 | 第2周=4组 | 第3周=3-4组 比前两周重1-2片 | 第4周=减载(见Sheet3)" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

$actsA=@(
    @("坐姿推胸","3组","12-15次","选一个做完15次最后3次有点累的重量","60秒","轻松-中等","【呼吸】推出呼气 收回吸气。【要点】后背贴紧靠背 推到前面不锁死手肘。你练过——找胸肌发力的感觉 应该很快就回来。"),
    @("坐姿推肩","3组","12-15次","最轻片起步 能轻松完成15次","60秒","轻松-中等","【呼吸】推起呼气 下放吸气。【要点】下背贴紧靠背 推到上面不锁肘。肩部线条——你上次减了7.5kg后肩膀应该比现在好看。"),
    @("坐姿腿举","3组","12-15次","空机或最轻片——找到大腿和臀部一起用力的感觉","60秒","中等","【呼吸】推出呼气 收回吸气。【要点】脚踩实与肩同宽 膝盖不要完全蹬直——留一点弯曲。你练过的 这个动作不陌生。"),
    @("坐姿髋外展","3组","15-20次","轻重量——臀外侧有感就行","45秒","轻松","【呼吸】打开呼气 收回吸气。【要点】身体贴紧靠背不晃。感受臀外侧发力。大腿外侧线条——这个动作长期做有效。"),
    @("高位下拉","3组","12-15次","最轻片——感受背阔肌发力","60秒","轻松-中等","【呼吸】拉下呼气 还原吸气。【要点】先沉肩膀再用手臂拉。杆拉到锁骨高度。上半身线条的关键动作。"),
    @("仰卧举腿","2组","10-15次","自重","45秒","中等","【要点】用下腹带动腿上举 不靠惯性摆。核心稳定性——所有动作的地基。")
)
$r=Write-TrainingBlock $ws2 $r "全身A  推力+下肢主导（每周第1练）" "胸+肩+大腿前侧+臀外侧+背阔肌+核心  |  35-40" "" "38-42" $actsA

# 全身B
SetC $ws2 $r 1 "全身B  拉力+下肢后侧主导  第1周" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=3组 中等重量 | 第2周=4组 | 第3周=3-4组 加重1-2片 | 第4周=减载" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

$actsB=@(
    @("坐姿划船","3组","12-15次","轻-中重量——找到肩胛骨往中间夹的感觉","60秒","轻松-中等","【呼吸】拉向身体呼气 还原吸气。【要点】身体不要后仰——想象肩胛骨往中间夹。拉到腹部位置。上半身体态改善的关键。"),
    @("高位下拉(宽握)","3组","12-15次","比全身A可以重1片——背已经激活了","60秒","中等","【发力感】宽握更偏向上背阔肌——和全身A的窄握角度不同 刺激不同区域。"),
    @("坐姿腿弯举","3组","12-15次","轻重量——大腿后侧有感","60秒","中等","【呼吸】勾起呼气 还原吸气。【要点】大腿后侧发力把小腿拉下来。改善大腿后侧线条——你上次应该练过这个。"),
    @("哑铃臀桥","3组","12-15次","双手持一颗哑铃6-10kg放髋部","60秒","中等","【呼吸】推起呼气 下放吸气。【要点】下巴微收。顶峰夹紧屁股1-2秒。臀大肌——改善臀型的核心动作。"),
    @("坐姿髋内收","3组","15-20次","轻重量","45秒","轻松","【呼吸】内收呼气 打开吸气。【要点】大腿内侧发力——改善大腿内侧线条。"),
    @("平板支撑","2组","20-30秒","自重","30秒","力竭前停","【要点】手肘正下方撑地 腰不要塌——收紧腹肌。能做多久做多久。")
)
$r=Write-TrainingBlock $ws2 $r "全身B  拉力+下肢后侧主导（每周第2练）" "上背+背阔肌+大腿后侧+臀部+大腿内侧+核心  |  35-40" "" "38-42" $actsB

# 全身C
SetC $ws2 $r 1 "全身C  综合+手臂+有氧偏向  第1周" $true 14 $white $softBlue $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

SetC $ws2 $r 1 "第1周=3组 | 第2周=4组 | 第3周=3-4组 加重1-2片 | 第4周=减载  |  全身C是综合日——前两天的动作混搭+手臂+有氧加时" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

$actsC=@(
    @("坐姿推胸","3组","12-15次","比全身A轻1片——综合日不冲重量","60秒","中等","要点同全身A。综合日的目的不是突破——是让身体在不同疲劳状态下感受同一个动作。"),
    @("高位下拉(窄握)","3组","12-15次","轻-中重量","60秒","中等","窄握更偏向下背阔肌——和全身B的宽握互补。三种握法 背阔肌全面发展。"),
    @("坐姿腿举","3组","12-15次","和全身A同样重量","60秒","中等","两次下肢训练间隔至少48h——刚好恢复到可以再练。"),
    @("哑铃侧平举","3组","12-15次","最轻哑铃2-4kg","45秒","中等","【要点】不要耸肩——用肩膀外侧发力 手臂微屈。肩部线条的收尾动作。"),
    @("绳索下压(肱三头肌)","2组","12-15次","龙门架轻重量","45秒","中等","【要点】大臂固定身体两侧 压到底时微微外旋手腕。手臂后侧紧致。"),
    @("有氧加时","跑步机爬坡走","15-20分钟","坡度8-10% 速度4.5-5km/h","—","中等","全身C的有氧比平时多5分钟——因为今天是本周最后一次训练 多消耗一点 周末恢复。")
)
$r=Write-TrainingBlock $ws2 $r "全身C  综合+手臂（每周第3练）" "胸+背+腿(综合)+肩+手臂+额外有氧  |  35-40+有氧" "" "42-48" $actsC

# Schedule
$r++
SetC $ws2 $r 1 "每周排班" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "推荐" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周一=全身A / 周三=全身B / 周五=全身C / 其他日子=休息+走路8000步" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "备选" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周二=全身A / 周四=全身B / 周六=全身C——如果周一有事的话" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8

# Progression
$r+=2
SetC $ws2 $r 1 "第1周  第4周 怎么变化" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

$prog=@(
    "第1周：重建动作记忆。每个动作3组 12-15次。选一个做完15次最后3次有点累的重量。记录所有重量。",
    "第2周：每个动作+1组（3组→4组）次数取上限(做15次)。组间休息缩短10秒。身体开始适应。",
    "第3周：组数回到3-4组。每个动作比前两周重1-2片。次数降到10-12次。找到'有点挑战但不变形'的感觉。",
    "第4周：减载周。每个动作2组 轻重量。不冲重量。具体换动作方案见Sheet3。主动恢复——让身体消化前三周。",
    "第4周结束时：翻出第一天的记录——你会看到自己比第1周强了多少。"
)
foreach($p in $prog){
    SetC $ws2 $r 1 $p $false 10 0 0 $true; MC $ws2 $r 1 8
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

SetC $ws3 1 1 "热身  有氧  拉伸  减载周  拍照" $true 16 $headerBlue 0 $true
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

SetC $ws3 $r 1 "每天步数：8000-10000步——和上次减脂一样 走路是基础消耗。训练日力量后在跑步机爬坡走10-15分钟(坡度8-10% 速度4.5-5km/h)。全身C日走15-20分钟。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=28; $r++

# Stretch
$r++
SetC $ws3 $r 1 "训练后拉伸（5分钟）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$strs=@(
    @("1","胸大肌拉伸(扶墙)","30s/侧","推胸后——手臂张开扶墙 身体反方向转"),
    @("2","背阔肌拉伸(扶墙侧屈)","30s/侧","下拉/划船后——手扶墙 身体侧弯"),
    @("3","大腿前侧拉伸(站姿脚跟贴臀)","30s/侧","扶墙——腿举后"),
    @("4","大腿后侧拉伸(坐姿直腿前屈)","30s","腿弯举后——坐地上 伸直腿 手够脚尖"),
    @("5","臀拉伸(跷二郎腿抱膝)","30s/侧","臀桥/髋外展后"),
    @("6","深呼吸","10次慢吸慢呼","让心跳慢下来——告诉身体训练结束 可以恢复了")
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
    $ws3.Range("A$($r):H$($r)").RowHeight=28; SB $ws3 $r 1 8; $r++
}

# Deload Week
$r+=2
SetC $ws3 $r 1 "第4周  减载周  换动作" $true 14 $white $softPurple $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "核心原则：第4周不是做轻一点——是做不同的动作。让主力关节换一个受力角度 = 给关节放假。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $r++

SetC $ws3 $r 1 "原来的动作" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws3 $r 4 8; $r++

$deloadSubs=@(
    @("坐姿推胸","俯卧撑(跪姿/控制节奏)","2组x8-10次 慢节奏"),
    @("坐姿推肩","弹力带/轻哑铃侧平举+前平举","2组x各10次 最轻重量"),
    @("高位下拉","弹力带高位下拉(单臂)","2组x10次/侧"),
    @("坐姿划船","哑铃单臂划船(最轻重量)","2组x10次/侧"),
    @("坐姿腿举","自重深蹲(扶椅背)","2组x15次 控制节奏"),
    @("坐姿腿弯举","早安式体前屈(自重)","2组x12次"),
    @("哑铃臀桥","自重臀桥+弹力带","2组x15次"),
    @("坐姿髋外展/内收","站姿弹力带侧向走","2组x15步/侧"),
    @("所有动作统一","2组 轻重量 RPE4-5","做完应该感觉'没练够'——这就对了")
)
foreach($ds in $deloadSubs){
    SetC $ws3 $r 1 $ds[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $ds[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $ds[2] $false 10 0 0 $true; MC $ws3 $r 4 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Photo
$r+=2
SetC $ws3 $r 1 "  拍照标准——今天就拍" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "今天拍三张：正面(手机与肚脐同高)+侧面(手机与腰同高)+背面(手机与肩胛骨同高)。穿运动内衣+紧身裤。不要吸肚子不要找角度。4周后同一位置同一光线再拍——照片不会骗人。你上次腰围小了7cm——这次再拍一组对比。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=32

$ws3.Range("A:A").ColumnWidth=22; $ws3.Range("B:B").ColumnWidth=26
$ws3.Range("C:C").ColumnWidth=14; $ws3.Range("D:H").ColumnWidth=16
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 饮食+规则
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="饮食与规则"

SetC $ws4 1 1 "饮食方案  经期应对  判断规则" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

# Nutrition
$r=3
SetC $ws4 $r 1 "饮食方案——你上次40天瘦7.5kg 饮食是主力" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "你上次怎么吃的还记得吗？下面这些原则你可能不陌生——不需要重新学，只需要重新做。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws4 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "目标" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "怎么吃（不称克数版本）" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$nutri=@(
    @("总热量","1500-1700 kcal","比维持少400-500kcal——让体重每周掉0.5-1kg。你上次减了7.5kg 说明这个缺口对你有效。"),
    @("蛋白质","100-120g","每餐有蛋白质：早2蛋+奶 午晚各一份肉/鱼/豆腐。练后喝蛋白粉更容易达标。上次瘦下来肌肉没掉太多——蛋白质功不可没。"),
    @("碳水","140-160g","三餐每餐一拳。晚餐换成粗粮(糙米/红薯/玉米/南瓜)——饱腹感更强 睡前不饿。上次你怎么吃碳水的？大概也是这个量。"),
    @("脂肪","35-45g","炒菜控油。不吃油炸。坚果一天一小把够了。"),
    @("饮水","2.5-3L/天","和你上次减脂一样——水喝够代谢才运转。戒含糖饮料——这是热量缺口的最大来源。"),
    @("蔬菜","不限量","每餐占一半盘子。先吃蔬菜再吃肉最后吃饭——这个顺序你上次应该也用过。")
)
foreach($n in $nutri){
    SetC $ws4 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $n[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=38; SB $ws4 $r 1 8; $r++
}

# Meal examples
$r++
SetC $ws4 $r 1 "三餐示例" $true 12 $headerBlue 0 $true; MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "餐次" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "要点" $true 10 $white $headerBlue $false; $r++

$meals=@(
    @("早餐","2个鸡蛋+一杯牛奶+一片全麦吐司或一小碗燕麦","蛋白质吃够上午不饿。和上次一样——早餐是全天饮食的锚点。"),
    @("午餐","米饭一拳+一份肉/鱼(手掌大)+蔬菜占半盘","食堂/外卖：选清蒸/炖/卤 避开油炸和勾芡。先吃菜再吃肉最后吃饭。"),
    @("下午","饿了就吃：小水果/无糖酸奶/坚果10颗以内","不饿可以不吃。别饿到晚餐暴食——你懂的。"),
    @("晚餐","粗粮一拳(糙米/红薯/玉米/南瓜)+一份肉/豆腐+大量蔬菜","晚餐吃粗粮不饿肚子。上次你瘦7.5kg的时候 晚餐怎么吃的？差不多的节奏。"),
    @("训练后","蛋白粉+水或牛奶","训练完30分钟内喝。上次你喝蛋白粉吗？没有的话训练后半小时内吃含蛋白质的正餐也行。")
)
foreach($m in $meals){
    SetC $ws4 $r 1 $m[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $m[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $m[2] $false 10 0 0 $true
    $ws4.Range("A$($r):H$($r)").RowHeight=36; SB $ws4 $r 1 8; $r++
}

# Period
$r+=2
SetC $ws4 $r 1 "经期前后" $true 14 $white $softPink $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

SetC $ws4 $r 1 "经前一周：体重涨1-3kg=正常水肿不是胖了。食欲暴增想吃甜=正常激素波动 选黑巧克力/水果/无糖酸奶。训练降强度但尽量去。经期第1-2天休息或走路 第3天后恢复轻量。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=28; $r++

# Weight expectation
$r++
SetC $ws4 $r 1 "掉秤预期——和上次一样" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$expect=@(
    "你上次40天瘦了7.5kg——平均每周约1.3kg。这次的速度可能略有不同：第一周可能掉1-2kg（水分为主），之后稳定在0.5-1kg/周。",
    "如果某周体重没变：先看是不是经前（水肿1-3kg正常）。不是的话回头检查饮食——是不是饮料没完全戒掉？外食多了？",
    "上次你体脂从43.1%降到39.5%——体脂率下降比体重下降慢 但更真实。腰围小了7cm——这次也量腰围。"
)
foreach($e in $expect){
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
    @("这周训练太轻松","第2周加1组 第3周加1-2片——上次你也经历过这个阶段"),
    @("训练后48h还酸痛","正常——停了一段时间重新练 前两周会酸。多喝水多睡。关节疼告诉我。"),
    @("体重一周没掉","先看是不是经期前后。不是的话检查饮料和晚餐碳水。"),
    @("太累了不想练","走5000步+拉伸——做了就算。但别连续跳2次——上次你坚持了40天 这次也可以。"),
    @("某天吃多了","明天继续 别等周一。一顿吃多不会胖——这个你上次就知道。"),
    @("膝盖/腰/关节疼","立即停那个动作 换不疼的。疼超2天告诉我。")
)
foreach($r3 in $rules){
    SetC $ws4 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 3 $r3[1] $false 10 0 0 $true; MC $ws4 $r 3 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=28; $ws4.Range("B:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 备用+复盘
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="备用与复盘"

SetC $ws5 1 1 "降级方案  4周复盘  下一步" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

# Fallback
$r=3
SetC $ws5 $r 1 "状态差/加班/来例假——降级方案" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws5 $r 1 "宁可降级训练 不要完全不练。降级不丢人——连续放弃才是问题。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "本来计划" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "最少完成" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$fallback=@(
    @("3次/周训练","减到2次 保留全身A+全身B","哪怕一周只练1次 也比不练强"),
    @("今天该练但很累","热身+最轻重量做3个动作+拉伸","走进健身房待20分钟=胜利"),
    @("完全不想出门","在家/楼下走20分钟+10分钟拉伸","也算运动了——走路永远不丢人"),
    @("来例假前3天","只走路 不训练","身体在准备月经 需要休息——不是偷懒")
)
foreach($fb in $fallback){
    SetC $ws5 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $fb[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $fb[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

# Review
$r+=2
SetC $ws5 $r 1 "4周后——怎么知道自己进步了" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "不要只看体重秤——你上次腰围小了7cm 比体重数字更真实。" $false 11 $accentGreen 0 $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "可能的变化" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "怎么看" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$progress=@(
    @("体重","掉了1-3kg","周三早上空腹称。不要每天称——女性体重每天波动。"),
    @("腰围","小了2-4cm","上次你7cm——这次4周2-4cm是合理的。"),
    @("照片","不一样了","翻出第一天的三张照片对比。照片不会骗人。"),
    @("衣服","松了","裤腰松了 上衣不绷了——和上次一样的感觉。"),
    @("力量","变强了","第1周推胸用的重量 第4周觉得太轻了——这就是进步。"),
    @("习惯","回来了","训练+饮食控制——已经不用'坚持'了 和你上次做到的一样。")
)
foreach($pg in $progress){
    SetC $ws5 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pg[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $pg[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=30; SB $ws5 $r 1 8; $r++
}

# Compare to last time
$r+=2
SetC $ws5 $r 1 "和上次比——这次有什么优势" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

$compare=@(
    "上次你从79.9kg开始——这次从72kg开始。起点已经低了8kg。",
    "上次你40天瘦了7.5kg——这次不用那么激进。稳定掉0.5-1kg/周就行。",
    "上次你有教练手把手——这次你靠自己。但这个方案已经给你铺好了路：每周三天 每次做什么 做几组 什么重量——照着走就行。",
    "你知道自己能做到——因为你已经做到过一次了。这次只是'继续' 不是'开始'。"
)
foreach($c in $compare){
    SetC $ws5 $r 1 $c $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++
}

# Next
$r++
SetC $ws5 $r 1 "4周之后" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$next=@(
    "减了2-3kg 体能恢复：继续保持这套方案再走4周。可以加重+加组。",
    "体重没动但腰细了力量涨了：脂肪换肌肉——继续！效果比体重秤显示的好。",
    "想加训练频率了：告诉我——下一周期可以调整到每周4练上下肢分化。",
    "",
    "最后：你上次40天瘦了7.5kg——你做到过一次 就能做到第二次。这次起点更低 基础更好。4周后见分晓。"
)
foreach($ns in $next){
    SetC $ws5 $r 1 $ns $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=20
    if($ns -eq ""){$ws5.Range("A$($r):H$($r)").RowHeight=10}
    $r++
}

$ws5.Range("A:A").ColumnWidth=26; $ws5.Range("B:H").ColumnWidth=20
Write-Host "Sheet 5 done"

# Save
$savePath="D:\Codex\蜡笔小杨_减脂方案_V1.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
