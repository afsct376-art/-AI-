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
    SetC $ws $ref 5 "负荷/RPE" $true 10 $white $headerBlue $false
    SetC $ws $ref 6 "休息" $true 10 $white $headerBlue $false
    SetC $ws $ref 7 "左右策略" $true 10 $white $headerBlue $false
    SetC $ws $ref 8 "动作要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):H$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        SetC $ws $ref 1 $i $false 10 0 $lightGray $false
        SetC $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 6;$j++){SetC $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 6)}
        $ws.Range("A$($ref):H$($ref)").RowHeight=56
        SB $ws $ref 1 8; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "毛丹  力量训练方案 V2" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 56kg  |  骑行+越野跑+铁三  |  术后重建平衡" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  V2 这次升级了什么" $true 16 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  左右平衡专项","每个上肢动作标注了左右策略——左侧(手术侧)额外+2-3次或+1组。目标是让两侧逐渐靠近 不是一步到位。"),
    @("  周期化结构","4周一个循环：第1周=容量积累(组数+1) 第2周=强度突破(加重) 第3周=峰值周(RPE8-9) 第4周=减载(恢复)。"),
    @("  赛前减量协议","环湖赛/长距离骑行前10天停下肢力量训练。赛前3天只做上肢轻量+核心+拉伸。赛后3天只做恢复性训练。"),
    @("  营养窗口","你是有氧大户——训练前后碳水比普通减脂人群高。骑行>2小时需要赛中补给策略。")
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "训练A/C(下肢)" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "训练B(上肢)" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周关键" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","容量积累","组数+1(4组) 次数上限","左侧额外+1组或+3次","建立动作储备 为强度周做准备"),
    @("2","强度突破","重量递增 RPE 8","左侧额外+3次 尝试和右侧同等重量","突破力量 左右差距缩小"),
    @("3","峰值周","冲击本周期上限 RPE 8-9","左侧独立加1组 两侧分别记录","冲——但左侧不勉强"),
    @("4","减载恢复","重量50-60% 换动作","两侧等量 轻重量","让身体消化 为下周期蓄力 赛前减量")
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
SetC $ws0 $row 1 "  每周安排" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$weekly=@(
    @("周一","训练A 臀大肌+下肢力量","力量——为骑行提供臀大肌驱动"),
    @("周二","骑行/有氧","耐力训练——保持有氧基础"),
    @("周三","训练B 上肢力量+肩胛稳定","力量——重点关注左右平衡"),
    @("周四","休息或轻量拉伸/游泳","恢复日——40+需要这天的修复"),
    @("周五","训练C 腘绳肌+爆发力","力量——腘绳肌是骑行拉提踏板的关键"),
    @("周六","长距离骑行/越野跑","耐力专项——你的主项"),
    @("周日","爬楼/跑步或完全休息","交叉训练或休息——看身体感受")
)
foreach($w in $weekly){
    SetC $ws0 $row 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $w[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $w[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  赛前减量协议" $true 14 $white $accentOrange $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentOrange
$row++

$taper=@(
    @("赛前10天","停止下肢力量训练(训练A和C)","保持上肢轻量+核心+骑行"),
    @("赛前3天","只做上肢轻量+核心+拉伸","完全停止下肢和爆发力训练"),
    @("赛后3天","恢复性训练：散步/游泳/拉伸","不碰器械。身体需要修复。"),
    @("赛后第4天","逐步恢复力量——从基础版开始","用赛前重量的60%——第一周不冲")
)
foreach($tp in $taper){
    SetC $ws0 $row 1 $tp[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $tp[1] $false 10 0 0 $false; MC $ws0 $row 2 3
    SetC $ws0 $row 4 $tp[2] $false 10 0 0 $true; MC $ws0 $row 4 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+左右平衡策略"," Sheet 方案概览"),
    @("训练A 臀大肌+下肢力量"," Sheet 训练A-下肢"),
    @("训练B 上肢力量+左右平衡"," Sheet 训练B-上肢"),
    @("训练C 腘绳肌+爆发力"," Sheet 训练C-腘绳"),
    @("热身+有氧+减载+赛前减量"," Sheet 热身减载"),
    @("饮食+骑行营养+判断规则"," Sheet 饮食规则")
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

SetC $ws 1 1 "毛丹  力量训练方案  概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 8; $ws.Range("A1:H1").RowHeight=36

SetC $ws 2 1 "女  41-45岁  160cm/56kg  体脂25-30%  骑行/越野跑/铁三  左乳切除术后(左右力量不均衡)  园林公司白领" $false 11 $gold $darkBg $true
MC $ws 2 1 8; $ws.Range("A2:H2").Font.Color=$white; $ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 8

$ov=@(
    @("训练频率","每周3次力量(周一/三/五)+耐力训练","力量服务于骑行——不是替代"),
    @("训练结构","A=臀大肌+下肢力 / B=上肢力+肩胛 / C=腘绳+爆发","每次含筋膜放松+热身+爆发+激活+力量+拉伸"),
    @("周期结构","容量周(组数+1)  强度周(加重)  峰值周(RPE8-9)  减载周(恢复)","4周一循环 每周末尾减载"),
    @("术后专项","左侧(手术侧)额外+2-3次或+1组——逐步缩小差距","不追求两侧完全相等 追求左侧持续进步"),
    @("方案特点","耐力运动员力量方案+术后左右平衡+赛前减量+骑行营养")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 5
    SetC $ws $r 6 $o[2] $false 10 0 0 $true; MC $ws $r 6 8
    SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "为什么力量训练对骑行重要" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$cycling=@(
    "骑行不只是腿在踩——臀大肌是真正的发动机。每次踩踏的下压阶段 臀大肌提供最大的输出。",
    "腘绳肌负责拉提踏板——很多骑行者腘绳肌偏弱 导致踩踏效率下降和下背代偿。训练C就是为你补这块。",
    "上肢力量——骑行时上半身要稳定控车长达数小时。肩胛稳定和核心力量直接决定长途骑行的舒适度和安全性。",
    "你已经在做铁三和越野跑——力量训练不是为了增肌。是为了让身体在长时间耐力运动中保持结构完整 不受伤。"
)
foreach($cy in $cycling){
    SetC $ws $r 1 $cy $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws $r 1 "术后左右平衡策略——本方案最重要的设计" $true 14 $white $softPink $true
MC $ws $r 1 8; $ws.Range("A$($r):H$($r)").RowHeight=26
$ws.Range("A$($r):H$($r)").Font.Color=$white; $ws.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

$balance=@(
    "你的左乳切除手术影响了左侧胸大肌和周围组织——直接导致左侧上肢的推力和拉力弱于右侧。",
    "这不是你的问题——这是手术的必然结果。但可以通过针对性训练逐步改善。",
    "",
    "本方案的左右策略（每个上肢动作都标注了）：",
    "  左侧额外+2-3次：双侧同样重量做完后 左侧单独多做2-3次。右侧不补。",
    "  左侧额外+1组：双侧做完规定的组数后 左侧独立加一组。右侧休息。",
    "  单侧先做原则：哑铃卧推、划船等单侧动作——左侧先做。右侧用左侧同样的次数和重量(不超)。",
    "  弹力带不对称训练：推胸/拉背时左侧用更粗的弹力带(更大阻力) 右侧正常。",
    "",
    "目标不是追求两侧完全相等——是让左侧持续进步 逐步缩小差距。这是一个以年为单位的过程 急不来。",
    "每4周测试一次：左右两侧分别做哑铃卧推至力竭——记录差距。看到数字在缩小 就是进步。"
)
foreach($ba in $balance){
    SetC $ws $r 1 $ba $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20
    if($ba -eq ""){$ws.Range("A$($r):H$($r)").RowHeight=10}
    $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 8; $r++

$pr=@(
    @("1","力量服务耐力","你不是在练健美——力量训练的目的是让你骑得更快更远更安全。每次训练前问自己：这个动作怎么帮助我的骑行？"),
    @("2","左侧优先 不追求相等","每个上肢动作左侧先做。左侧额外加量。目标不是两侧相等——是左侧持续进步。"),
    @("3","赛前减量 赛后恢复","环湖赛/长距离前10天停下肢力量。赛前3天只上肢轻量。赛后3天不练——身体不是你练的时候变强 是恢复的时候。"),
    @("4","40+的恢复是训练的一部分","你不是20岁——恢复能力不同了。周四休息日是真的休息。减载周是真的减载。睡够7-8小时。")
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
# SHEET 2: 训练A 臀大肌+下肢
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练A-下肢"

SetC $ws2 1 1 "训练A  臀大肌+下肢力量（周一）" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "骑行发动机——臀大肌是踩踏下压阶段的最大输出来源。A日=臀推主项+下肢辅助+核心。第1周=组数+1(容量) 第2周=加重 第3周=峰值 第4周=减载。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

$r=4
# Week 1 actions
$actsA=@(
    @("杠铃臀推(主项)","热身1x15+正式4组","8-10次","RPE 7-8 第1周容量=多1组","90s","—","【呼吸】推起呼气下放吸气。【要点】肩胛骨下缘靠凳 下巴收 顶峰夹臀2秒。【骑行迁移】这是你踩踏下压的力量来源。"),
    @("高脚杯深蹲","4组","10-12次","RPE 7 哑铃/壶铃贴胸口","90s","—","【呼吸】下蹲吸气起身呼气。【要点】膝盖对准脚尖 屁股往后坐 躯干竖直。【骑行迁移】深蹲=爬坡时离开座垫的力量。"),
    @("哑铃反向弓步","4组","10次/侧","RPE 7 双手持哑铃","60s","—","【呼吸】下蹲吸气起身呼气。【要点】后退一步 前脚跟发力站起。上半身前倾5度=骑行姿态。【骑行迁移】单侧发力=踩踏的左右交替模式。"),
    @("坐姿髋外展","3组","15次","RPE 7-8 臀中肌发力","45s","—","【呼吸】打开呼气收回吸气。【要点】身体不晃 顶峰停1秒。【骑行迁移】臀中肌=骑行时膝盖不外翻的稳定器。"),
    @("单腿臀桥","3组","12次/侧","自重或持轻哑铃","45s","左侧多2次","【呼吸】推起呼气下放。【要点】脚跟着地 顶峰夹臀3秒。【骑行迁移】单侧臀大肌独立驱动——弥补骑行中可能的左右发力不均。"),
    @("站姿提踵","3组","15-20次","RPE 7 膝盖微屈","30s","—","【呼吸】提踵呼气下放吸气。【要点】顶峰停1秒 慢放。【骑行迁移】小腿=踩踏的末端推力+足踝稳定。")
)
$r=Write-TrainingBlock $ws2 $r "第1周 容量积累（组数+1 次数上限 重量适中）" "臀推主项+深蹲+弓步+髋外展+臀桥+提踵  |  50-60" "每个动作多1组——为下周加重储备体能" "50-60" $actsA

# Week 2-3 progression
$r++
SetC $ws2 $r 1 "第2周 强度突破 & 第3周 峰值周" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "第2周(强度)" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "第3周(峰值)" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$progA=@(
    @("杠铃臀推","3组 8-10次 +5-10kg RPE8","3组 6-8次 再+5kg RPE8-9","2组 10-12次 轻重量 自重臀桥"),
    @("高脚杯深蹲","3组 10-12次 +2-4kg","3组 8-10次 再+2kg RPE8-9","2组 自重深蹲 15次"),
    @("哑铃反向弓步","3组 10次/侧 +2-4kg","3组 8次/侧 再+2kg","2组 自重弓步 10次/侧"),
    @("坐姿髋外展","3组 15次 +1片","3组 12-15次 再+1片","2组 轻量 20次"),
    @("单腿臀桥","3组 12次/侧 左侧+2次","3组 10次/侧 左侧+3次","2组 自重 15次/侧"),
    @("站姿提踵","3组 15-20次 +5kg","3组 12-15次 再+5kg","2组 自重 20次")
)
foreach($pa in $progA){
    SetC $ws2 $r 1 $pa[0] $true 10 $headerBlue $lightGray $false
    SetC $ws2 $r 2 $pa[1] $false 10 0 0 $true; MC $ws2 $r 2 3
    SetC $ws2 $r 4 $pa[2] $false 10 0 0 $true; MC $ws2 $r 4 5
    SetC $ws2 $r 6 $pa[3] $false 10 0 0 $true; MC $ws2 $r 6 8
    $ws2.Range("A$($r):H$($r)").RowHeight=32; SB $ws2 $r 1 8; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=12; $ws2.Range("H:H").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 训练B 上肢+左右平衡
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="训练B-上肢"

SetC $ws3 1 1 "训练B  上肢力量+左右平衡（周三）" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34
SetC $ws3 2 1 "骑行上半身稳定——肩胛控制+核心力量=长途骑行的舒适度和安全。左右策略是本日核心：左侧(手术侧)额外加量。第1周=容量 第2周=强度 第3周=峰值 第4周=减载。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 8

$r=4
$actsB=@(
    @("高位下拉(主项)","热身1x15+正式4组","10-12次","RPE 7-8 宽握","90s","左侧单臂下拉+3次","【呼吸】下拉呼气还原吸气。【要点】先沉肩胛再拉 杆拉到锁骨。【骑行迁移】背阔肌=控车时上半身的稳定锚点。"),
    @("坐姿划船","4组","10-12次","RPE 7","60s","左侧额外+1组","【呼吸】拉向身体呼气还原。【要点】拉到腹部 肩胛后缩。【骑行迁移】菱形肌=长时间趴低骑行不驼背。"),
    @("哑铃卧推","4组","10-12次","RPE 7 从左侧先做","60s","左侧先做 左侧额外+3次","【呼吸】下放吸气推起呼气。【要点】肘与躯干45-60度。【术后注意】左侧胸大肌被部分切除——不追求和右侧同样重量。感受左侧剩余的胸肌纤维在收缩。"),
    @("哑铃站姿推举","3组","10-12次","RPE 7 核心收紧","60s","左侧先做 左侧额外+2次","【呼吸】推起呼气下放。【要点】从肩上方起 不塌腰。【术后注意】左侧推举感受三角肌和肱三头肌代偿胸大肌的功能——这是好事 它们在帮你。"),
    @("面拉(绳索)","3组","15次","轻重量 RPE 7","45s","左侧额外+3次","【呼吸】拉向面部呼气。【要点】拉向额头 肘外展 肩后束发力。【骑行迁移】肩后束+上背=长途骑行不耸肩 脖子不酸。"),
    @("弹力带不对称推胸","3组","12次","左侧用更粗弹力带 右侧正常","45s","左侧粗带 右侧细带","【要点】站姿双手推弹力带——左侧更大阻力=针对性强化。【术后专项】这是最直接缩小左右差距的训练。")
)
$r=Write-TrainingBlock $ws3 $r "第1周 容量积累（组数+1 左侧额外加量）" "高位下拉+划船+卧推+推举+面拉+弹力带不对称  |  45-55" "每个上肢动作都标了左右策略——左侧是这一天的重点" "45-55" $actsB

# Week 2-3 progression
$r++
SetC $ws3 $r 1 "第2周 强度突破 & 第3周 峰值周（上肢加重幅度比下肢小——上肢关节精细）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "第2周(强度)" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "第3周(峰值)" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$progB=@(
    @("高位下拉","3组 8-10次 +1片 RPE8","3组 6-8次 再+1片 RPE8-9","2组 12-15次 轻量"),
    @("坐姿划船","3组 8-10次 +1片","3组 8-10次 再+1片","2组 12-15次 轻量"),
    @("哑铃卧推","3组 8-10次 左侧先 +1-2kg","3组 6-8次 再+1kg RPE8-9","2组 12-15次 轻量/俯卧撑"),
    @("哑铃推举","3组 8-10次 +1-2kg","3组 6-8次 RPE8-9","2组 12-15次 轻量/弹力带"),
    @("面拉","3组 12-15次 微加重","3组 12次 再微加","2组 15次 轻量"),
    @("弹力带不对称","左侧继续粗带 微增阻力","左侧挑战更粗带","两侧同细带 2组 15次")
)
foreach($pb in $progB){
    SetC $ws3 $r 1 $pb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pb[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pb[2] $false 10 0 0 $true; MC $ws3 $r 4 5
    SetC $ws3 $r 6 $pb[3] $false 10 0 0 $true; MC $ws3 $r 6 8
    $ws3.Range("A$($r):H$($r)").RowHeight=32; SB $ws3 $r 1 8; $r++
}

$ws3.Range("A:A").ColumnWidth=4; $ws3.Range("B:B").ColumnWidth=20
$ws3.Range("C:C").ColumnWidth=7; $ws3.Range("D:D").ColumnWidth=12
$ws3.Range("E:E").ColumnWidth=18; $ws3.Range("F:F").ColumnWidth=8
$ws3.Range("G:G").ColumnWidth=12; $ws3.Range("H:H").ColumnWidth=52
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 训练C 腘绳+爆发
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="训练C-腘绳"

SetC $ws4 1 1 "训练C  腘绳肌+爆发力（周五）" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34
SetC $ws4 2 1 "腘绳肌=骑行拉提踏板的关键。很多骑行者腘绳肌偏弱导致下背代偿。C日=RDL主项+爆发力+腘绳辅助。第1周=容量 第2周=强度 第3周=峰值 第4周=减载。" $false 10 $headerBlue $lightGray $true
MC $ws4 2 1 8

$r=4
$actsC=@(
    @("罗马尼亚硬拉RDL(主项)","热身1x15空杆+正式4组","8-10次","RPE 7-8 杠铃或哑铃","90s","—","【呼吸】下放吸气拉起呼气。【要点】膝盖微屈 杠铃/哑铃贴大腿下滑 全程直背。【骑行迁移】腘绳肌=踩踏拉提踏板的力量。这个动作是你骑行表现的直接增益。"),
    @("跳箱(低箱渐进)","3组","5-6次","40-50cm 从低箱热身","90s","—","【呼吸】起跳呼气落地吸气。【要点】摆臂带动 落地轻柔屈膝缓冲。【骑行迁移】爆发力=冲刺和爬坡加速的能力。 箱旁放垫子 不跳超能力范围。"),
    @("单腿臀推(垫高)","3组","10次/侧","RPE 7 支撑腿臀推","60s","左侧多2次","【呼吸】推起呼气。【要点】单腿架凳另一腿抬起 支撑腿臀推起。【骑行迁移】单侧臀大肌独立发力=纠正骑行中左右发力不均。"),
    @("坐姿腿弯举","3组","12-15次","RPE 7-8 离心3秒","45s","—","【呼吸】勾起呼气还原吸气。【要点】大腿后侧发力 慢放不砸片。【骑行迁移】腘绳肌孤立训练——弥补RDL后的疲劳累积。"),
    @("反向北欧降","3组","6-8次","自重 身边放垫子","60s","—","【呼吸】后倒吸气 控制还原。【要点】跪姿身体缓慢后倒 腘绳肌离心控制。【骑行迁移】腘绳肌离心力量=长时间骑行后段腿不软。!! 手扶地面控制幅度 不勉强。"),
    @("HIIT可选:固定单车冲刺","4-6轮","30s全力/60s恢复","RPE 8-9","60s恢复","—"," 次日有长距离骑行则跳过。30秒全力踩踏 60秒轻柔踩恢复。【骑行迁移】比赛冲刺能力的直接训练。")
)
$r=Write-TrainingBlock $ws4 $r "第1周 容量积累（组数+1 RDL+爆发+腘绳+可选HIIT）" "RDL主项+跳箱+单腿臀推+腿弯举+北欧降+可选冲刺  |  50-60" "次日有长距离骑行→跳HIIT。赛前10天→停跳箱和HIIT" "50-60" $actsC

# Week 2-3 progression
$r++
SetC $ws4 $r 1 "第2周 强度突破 & 第3周 峰值周" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "第2周(强度)" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "第3周(峰值)" $true 10 $white $headerBlue $false
SetC $ws4 $r 6 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$progC=@(
    @("RDL","3组 6-8次 +5-10kg RPE8","3组 4-6次 再+5kg RPE8-9","2组 轻量 哑铃RDL 12次"),
    @("跳箱","3组 5次 尝试高一档","3组 4次 全力起跳","跳过——减载周不做爆发"),
    @("单腿臀推","3组 8-10次/侧 +2kg","3组 6-8次/侧 RPE9","2组 自重 15次/侧"),
    @("腿弯举","3组 10-12次 +1片","3组 8-10次 再+1片","2组 轻量 15次"),
    @("反向北欧降","3组 6-8次 尝试加深幅度","3组 5-6次 更大幅度","2组 手扶地面 浅幅度"),
    @("HIIT冲刺","4-6轮 强度周照做","5-6轮 最后一周冲","跳过——减载")
)
foreach($pc in $progC){
    SetC $ws4 $r 1 $pc[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $pc[1] $false 10 0 0 $true; MC $ws4 $r 2 3
    SetC $ws4 $r 4 $pc[2] $false 10 0 0 $true; MC $ws4 $r 4 5
    SetC $ws4 $r 6 $pc[3] $false 10 0 0 $true; MC $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight=32; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=4; $ws4.Range("B:B").ColumnWidth=20
$ws4.Range("C:C").ColumnWidth=7; $ws4.Range("D:D").ColumnWidth=12
$ws4.Range("E:E").ColumnWidth=18; $ws4.Range("F:F").ColumnWidth=8
$ws4.Range("G:G").ColumnWidth=12; $ws4.Range("H:H").ColumnWidth=52
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 热身+有氧+减载+赛前
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="热身减载"

SetC $ws5 1 1 "热身流程  减载周  赛前减量" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

# Warmup
$r=3
SetC $ws5 $r 1 "训练前完整流程（7步 保留你上一版的结构）" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws5 $r 2 "阶段" $true 9 $white $headerBlue $false
SetC $ws5 $r 3 "内容(下肢日/上肢日)" $true 9 $white $headerBlue $false
SetC $ws5 $r 5 "时间" $true 9 $white $headerBlue $false; MC $ws5 $r 5 8; $r++

$flow=@(
    @("1","筋膜放松","泡沫轴放松相关肌群(下肢日=股四头+腘绳+臀+小腿 上肢日=上背+背阔+胸+斜方)","5min"),
    @("2","动态热身","猫式伸展+青蛙趴+抱膝行进+燕式平衡+最伟大伸展+平板支撑","5min"),
    @("3","爆发力","下肢日=跳箱+药球砸地 上肢日=药球过头砸地+药球旋转抛","3min"),
    @("4","激活","下肢日=蚌式开合+单腿臀桥+死虫式+俯卧IYTWO 上肢日=弹力带拉开+肩胛后缩+死虫式+俯卧YTWL","5min"),
    @("5","力量训练","主项+辅助 按当天Sheet","35-45min"),
    @("6","HIIT可选","固定单车冲刺/战绳 次日有长距离则跳过","5-10min"),
    @("7","拉伸","鸽子式+髋屈肌拉伸+腘绳肌拉伸+蝴蝶式/胸肌拉伸+背阔肌+婴儿式","5min")
)
foreach($fl in $flow){
    SetC $ws5 $r 1 $fl[0] $false 9 0 0 $false
    SetC $ws5 $r 2 $fl[1] $true 9 $headerBlue 0 $false
    SetC $ws5 $r 3 $fl[2] $false 9 0 0 $true; MC $ws5 $r 3 4
    SetC $ws5 $r 5 $fl[3] $false 9 0 0 $false; MC $ws5 $r 5 8
    $ws5.Range("A$($r):H$($r)").RowHeight=30; SB $ws5 $r 1 8; $r++
}

# Deload
$r+=2
SetC $ws5 $r 1 "第4周  减载周  换动作 轻重量" $true 14 $white $softPurple $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

SetC $ws5 $r 1 "所有动作2组 轻重量 RPE 4-5。两侧等量——不做额外左侧加量。做完应该感觉'没练够'——这就是对的。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "减载周替换" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$deload=@(
    @("杠铃臀推","自重臀桥+弹力带","2x15次"),
    @("高脚杯深蹲","自重深蹲(扶墙)","2x15次"),
    @("RDL","早安式体前屈(自重)","2x12次"),
    @("高位下拉","弹力带高位下拉(单臂)","2x10次/侧"),
    @("哑铃卧推","跪姿俯卧撑","2x10次"),
    @("跳箱","跳过——减载周不做爆发",""),
    @("HIIT","跳过",""),
    @("骑行","低强度休闲骑——不冲PR","")
)
foreach($dl in $deload){
    SetC $ws5 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $dl[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

# Race taper
$r+=2
SetC $ws5 $r 1 "赛前减量协议（环湖赛/长距离赛事前必读）" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

$taperDetail=@(
    @("赛前10天","停止下肢力量(训练A和C)","保持上肢轻量+核心+正常骑行"),
    @("赛前7天","骑行减量至平时的60-70%","不再做任何HIIT或冲刺训练"),
    @("赛前3天","只做上肢轻量+核心+拉伸","完全停止下肢力量+爆发力。骑行仅限于休闲骑。"),
    @("赛前1天","完全休息或散步","不训练。碳水加载(多吃一碗饭)。多喝水。早睡。"),
    @("比赛日","热身10-15分钟轻度骑行","赛中补给：每小时30-60g碳水+电解质。你是有经验的车手——这个你应该熟。"),
    @("赛后1-3天","恢复性训练：散步/游泳/拉伸","不碰器械。身体在修复微创伤——训练不会让你变强 恢复才会。"),
    @("赛后第4天起","逐步恢复力量——用赛前60%重量","第一周不冲重量 不做到力竭。第二周回到正常周期。")
)
foreach($td in $taperDetail){
    SetC $ws5 $r 1 $td[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $td[1] $false 10 0 0 $false; MC $ws5 $r 2 3
    SetC $ws5 $r 4 $td[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=28; SB $ws5 $r 1 8; $r++
}

$ws5.Range("A:A").ColumnWidth=18; $ws5.Range("B:B").ColumnWidth=24
$ws5.Range("C:C").ColumnWidth=14; $ws5.Range("D:H").ColumnWidth=16
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 饮食+骑行营养+规则
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="饮食规则"

SetC $ws6 1 1 "骑行营养  日常饮食  判断规则" $true 16 $headerBlue 0 $true
MC $ws6 1 1 8; $ws6.Range("A1:H1").RowHeight=34

# Cycling nutrition
$r=3
SetC $ws6 $r 1 "骑行营养——你是有氧大户 和普通减脂人群吃法不同" $true 14 $white $accentGreen $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws6 $r 1 "你是耐力运动员——碳水是你的朋友 不是敌人。和普通减脂人群不同 你需要足够的碳水支撑训练和恢复。" $false 10 $headerBlue $lightGray $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws6 $r 1 "场景" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么吃" $true 10 $white $headerBlue $false
SetC $ws6 $r 4 "逻辑" $true 10 $white $headerBlue $false; MC $ws6 $r 4 8; $r++

$cyclingNutri=@(
    @("力量训练日","练前2h：碳水30-40g+蛋白15-20g。练后1h内：碳水40-50g+蛋白25-30g。","力量训练消耗肌糖原——训练后窗口补碳水效率最高。"),
    @("骑行<1小时","正常吃 不需要额外补给。骑前正常一餐。","短距离不吃也能骑——身体有足够的糖原储备。"),
    @("骑行1-2小时","骑前碳水30-40g。骑中：每小时一瓶运动饮料或能量胶。","中等距离——开始需要赛中补给了。"),
    @("骑行>2小时","骑前碳水40-60g。骑中：每小时30-60g碳水+电解质。骑后：碳水+蛋白 30分钟内。","长距离——赛中补给是完成的基础。骑后窗口=恢复速度翻倍。"),
    @("休息日","碳水比训练日少20-30%——不需要那么高的碳水。蛋白质不变。","不练的日子不需要那么多能量——自然少一点就好。"),
    @("全天","蛋白质100-120g 碳水训练日200-250g/休息日150-180g 脂肪45-55g 水3L。","你56kg——保持肌肉量需要的蛋白量。碳水化合物是你运动的燃料 不是敌人。")
)
foreach($cn in $cyclingNutri){
    SetC $ws6 $r 1 $cn[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $cn[1] $false 10 0 0 $true; MC $ws6 $r 2 3
    SetC $ws6 $r 4 $cn[2] $false 10 0 0 $true; MC $ws6 $r 4 8
    $ws6.Range("A$($r):H$($r)").RowHeight=36; SB $ws6 $r 1 8; $r++
}

# Rules
$r+=2
SetC $ws6 $r 1 "身体信号——什么时候停 什么时候调" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws6 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws6 $r 3 8; $r++

$rules=@(
    @("左侧训练时出现锐痛(非肌肉酸痛)","立即停止 降重量。术后区域疼痛不是'练到位'——是需要谨慎。"),
    @("训练后48h+仍然极度疲劳","训练量或强度过高——下周降10%。40+的恢复能力不同 听身体的。"),
    @("连续2周左侧力量没有进步","换策略：弹力带不对称训练为主 减少杠铃/哑铃比例。"),
    @("骑行时膝盖前侧疼痛","检查车座高度和踩踏技术。暂停下肢力量训练中的深蹲和弓步。"),
    @("某天极度疲劳不想练","降级：只做热身+激活+拉伸。不强迫——但保持去健身房的习惯。"),
    @("两侧力量差距在缩小","——这是最好的信号。继续当前策略 不急。"),
    @("环湖赛前10天","启动赛前减量协议——见Sheet5。")
)
foreach($r3 in $rules){
    SetC $ws6 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 3 $r3[1] $false 10 0 0 $true; MC $ws6 $r 3 8
    $ws6.Range("A$($r):H$($r)").RowHeight=28; SB $ws6 $r 1 8; $r++
}

$ws6.Range("A:A").ColumnWidth=28; $ws6.Range("B:H").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\毛丹_力量训练方案_V2.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
