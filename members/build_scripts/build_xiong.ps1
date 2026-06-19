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

function Write-TrainingBlock($ws,$ref,$title,$subtitle,$time,$acts){
    SetC $ws $ref 1 $title $true 13 $white $darkBg $true; MC $ws $ref 1 9
    $ws.Range("A$($ref):I$($ref)").RowHeight=26
    $ws.Range("A$($ref):I$($ref)").Font.Color=$white
    $ws.Range("A$($ref):I$($ref)").Interior.Color=$darkBg; $ref++
    SetC $ws $ref 1 "$subtitle  |  $time 分钟" $false 10 $headerBlue $lightGray $true; MC $ws $ref 1 9; $ref++
    SetC $ws $ref 1 "序" $true 10 $white $headerBlue $false
    SetC $ws $ref 2 "动作" $true 10 $white $headerBlue $false
    SetC $ws $ref 3 "角色" $true 10 $white $headerBlue $false
    SetC $ws $ref 4 "组数" $true 10 $white $headerBlue $false
    SetC $ws $ref 5 "次数" $true 10 $white $headerBlue $false
    SetC $ws $ref 6 "负荷/RPE" $true 10 $white $headerBlue $false
    SetC $ws $ref 7 "休息" $true 10 $white $headerBlue $false
    SetC $ws $ref 8 "策略" $true 10 $white $headerBlue $false
    SetC $ws $ref 9 "技术要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):I$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        SetC $ws $ref 1 $i $false 10 0 $lightGray $false
        SetC $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 7;$j++){SetC $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 7)}
        $ws.Range("A$($ref):I$($ref)").RowHeight=52
        SB $ws $ref 1 9; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "熊  减重+越野跑 力量方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 90kg  |  铁三Pro  减重+50km越野备赛  腰痛恢复期" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  这份方案的核心——比你完赛铁三更重要的三件事" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "你能完赛铁三——那不是运气 是你做对了一系列事情。这次和之前不一样：你在腰痛恢复期 要减重 要备赛50km越野。三件事不矛盾——但需要先后顺序。" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=42; $row++

$row+=2
SetC $ws0 $row 1 "  三件事的优先级——不是同时做 是有先后" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$priority=@(
    @("第一优先","腰痛恢复——任何时候腰不舒服 训练降级或跳过。没有健康的腰 越野跑=不可能。"),
    @("第二优先","减重——90kg对160cm的身高来说 膝盖和腰承受的压力太大了。每减1kg 越野跑时膝盖少承受4kg冲击力。"),
    @("第三优先","备赛50km越野跑——你的有氧基础已经很好(铁三完赛) 但越野跑需要的力量类型和铁三不同：上坡推力+下坡控制+核心抗扭转。")
)
foreach($pr in $priority){
    SetC $ws0 $row 1 $pr[0] $true 10 $softBlue 0 $false
    SetC $ws0 $row 2 $pr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=28; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览——从恢复开始 逐步重返" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "力量训练" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "有氧/越野专项" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周关键" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","腰痛恢复+适应","2次/周 轻量 核心+臀激活","游泳+低强度骑行 保持有氧","测试腰：哪些动作疼 哪些不疼。不疼的动作才是本周的训练内容。"),
    @("2","重建力量基础","3次/周 全身力量 逐步加量","跑步+骑行 维持有氧","腰痛如果在减轻→动作范围和重量可以微增。如果不减轻→回到第1周。"),
    @("3","力量进阶","3次/周 上下肢分化 加重量","跑步+越野模拟(爬坡走)","上坡推力+下坡控制 专项力量。腰痛应该明显改善了。"),
    @("4","减载+测试","2次/周 轻量 换动作","保持有氧 不冲PR","测体重+腰围+越野跑5km计时。拍对比照。评估腰痛恢复程度。")
)
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=32; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+腰痛约束+三件事优先级"," Sheet 方案概览"),
    @("训练A(下肢+核心) B(上肢+稳定性) C(越野专项)"," Sheet 训练计划"),
    @("渐进+腰痛决策树+器械突破"," Sheet 渐进决策树"),
    @("热身+主动恢复+减载+腰痛康复动作"," Sheet 恢复策略"),
    @("减重饮食+越野补给+判断规则"," Sheet 饮食与规则"),
    @("精简版+4周复盘+备赛提醒"," Sheet 备用与复盘")
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

SetC $ws 1 1 "熊  减重+越野跑  力量方案概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 9; $ws.Range("A1:I1").RowHeight=36

SetC $ws 2 1 "女  25-30岁  160cm/90kg  BMI 35.2  体脂30-35%  铁三完赛(Pro级)  越野跑备赛  骑车被撞腰痛恢复中  游泳/骑行/跑步/瑜伽" $false 11 $gold $darkBg $true
MC $ws 2 1 9; $ws.Range("A2:I2").Font.Color=$white; $ws.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 9

$ov=@(
    @("训练频率","第1周=2次/周 第2-4周=3次/周","力量服务于越野跑——不是替代"),
    @("训练结构","A=下肢力+核心(腰痛适应) / B=上肢+稳定性 / C=越野专项(爬坡+下坡)","每次含腰痛评估+核心激活+力量+拉伸"),
    @("周期结构","第1周=恢复+测试 第2周=重建 第3周=进阶 第4周=减载+测试","腰痛优先——任何时候疼就降级"),
    @("核心约束","腰痛恢复期——所有动作以无痛为第一原则","下背保护+核心激活是每节课的必修"),
    @("方案特点","减重+越野跑+腰痛保护 三合一。完赛铁三的Pro——给你运动员级别的方案")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 6
    SetC $ws $r 7 $o[2] $false 10 0 0 $true; MC $ws $r 7 9
    SB $ws $r 1 9; $r++
}

# Three priorities V3
$r++
SetC $ws $r 1 "三件事的优先级——为什么这个顺序" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$why=@(
    "你的情况特殊：能完赛铁三=有氧基础极好。但90kg的体重+腰痛=不能用常规运动员方案。",
    "",
    "第一优先：腰痛恢复。骑车被撞后的腰痛——不是肌肉酸痛 是创伤恢复。第1周的核心任务不是训练——是测试：哪些动作疼 哪些不疼。不疼的动作=本周的训练内容。疼的动作=跳过 下周再试。",
    "",
    "第二优先：减重。90kg/160cm——每减1kg 越野跑时膝盖少承受约4kg冲击力。减10kg=膝盖每步少承受40kg。这比任何力量训练对越野跑的帮助都大。你是有氧Pro——饮食控制+保持有氧=减重的主力。力量训练让身体在减重过程中保持肌肉。",
    "",
    "第三优先：越野跑专项。你的有氧基础不是问题——铁三完赛证明了这一点。越野跑需要的和铁三不一样：上坡推力(臀+股四) 下坡控制(股四头离心+核心抗扭转) 长时间单腿稳定(臀中肌)。这些是力量训练要给你的。"
)
foreach($w in $why){
    SetC $ws $r 1 $w $false 10 0 0 $true; MC $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight=20
    if($w -eq ""){$ws.Range("A$($r):I$($r)").RowHeight=8}
    $r++
}

# Back pain constraint
$r++
SetC $ws $r 1 "腰痛训练原则——每节课的必修" $true 14 $white $accentRed $true
MC $ws $r 1 9; $ws.Range("A$($r):I$($r)").RowHeight=26
$ws.Range("A$($r):I$($r)").Font.Color=$white; $ws.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

$back=@(
    " 禁止：传统硬拉 大重量深蹲 负重俯身划船 跳箱 负重跑步——所有对腰椎有冲击或大压力的动作。",
    " 谨慎：深蹲(从自重开始 腰不痛才加重量) 罗马尼亚硬拉(轻哑铃 全程直背 腰有不适立即停)。",
    " 安全：臀推/臀桥(零腰压) 核心稳定性训练(死虫式/鸟狗式/平板支撑) 高位下拉 坐姿划船(固定器械 减少腰的参与)。",
    " 康复：死虫式+鸟狗式(每节课热身必做——核心稳定性是保护腰椎的第一道防线)。猫式伸展(脊柱活动 腰不痛时做)。",
    "  停止信号：训练中腰出现锐痛/刺痛/放射性疼痛→立即停 当天不练。酸胀感/紧绷感=正常 可以继续但减量。",
    "第1周的核心任务：测试。把你平时做的所有动作过一遍——哪些腰不疼？记录下来。疼的动作=本周不做 第2周再试。"
)
foreach($ba in $back){
    SetC $ws $r 1 $ba $false 10 0 0 $true; MC $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight=24; $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$pr=@(
    @("1","腰痛优先于一切","你完赛铁三的意志力是优势——但也可能让你在腰痛时'硬撑'。这次不一样：腰痛是创伤 不是酸痛。任何时候腰不舒服→降级或停。"),
    @("2","减重=最好的越野跑训练","90kg跑50km越野——膝盖承受的压力太大了。每减1kg=膝盖每步少承受4kg。饮食控制+保持有氧=这4周最大的进步来源。"),
    @("3","力量服务越野跑","你不是在练健美——力量训练让你爬坡更有力 下坡更受控 长途不受伤。每次训练前问自己：这个动作怎么帮助我的越野跑？")
)
foreach($p in $pr){
    SetC $ws $r 1 $p[0] $false 10 0 0 $false
    SetC $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws $r 3 $p[2] $false 10 0 0 $true; MC $ws $r 3 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

$ws.Range("A:A").ColumnWidth=30; $ws.Range("B:I").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划——第1-4周 力量+越野专项" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "第1周=测试周(找到腰的安全动作范围)。第2-3周=全身力量 逐步加量。第4周=减载+测试。每周2-3次力量 其余天=游泳/骑行/跑步。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 训练A: 下肢+核心 (腰痛适应版)
$r=4
SetC $ws2 $r 1 "第1周 测试周——找到腰的安全动作范围" $true 14 $white $accentGreen $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws2 $r 1 "第1周只做2次 每次把下面动作过一遍。每个动作做完问自己：腰痛吗？不疼=记录+下周加重。疼=跳过+下周再试。这是测试 不是训练。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=24; $r++

$actsA1=@(
    @("死虫式","核心激活 必做","2组","每侧10次","自重 慢速控制","30s","腰痛康复第一动作——核心稳定=护腰","【要点】下背贴死地面 对侧手脚同步伸出收回。腰贴地=对 腰离地=错。"),
    @("鸟狗式","核心稳定 必做","2组","每侧8次","自重 慢速控制","30s","和死虫式搭配——核心前侧+后侧","【要点】骨盆不旋转 对侧手脚延伸 身体一条线。腰不塌。"),
    @("猫式伸展","脊柱活动(腰不痛时做)","2组","10次","自重 慢节奏","30s","脊柱逐节活动——腰痛缓解","【要点】吸抬头塌腰 呼弓背收腹。如果弓背时腰疼→跳过。"),
    @("自重臀桥","测试：腰不疼就做","2组","15次","自重","45s","零腰压——臀大肌激活","【要点】脚跟着地 顶峰夹臀2秒。腰完全不受力。如果这个都腰痛→当天不做任何下肢训练。"),
    @("自重深蹲(扶墙)","测试：腰不疼就做","2组","10-12次","自重 扶墙保持平衡","60s","测试腰对深蹲的反应","【要点】能蹲多深蹲多深。腰不疼=记录深度。腰疼=跳过。"),
    @("轻哑铃罗马尼亚硬拉","测试：用最轻哑铃4-6kg","2组","10次","最轻哑铃 全程直背","60s","测试腰对髋铰链的反应","【要点】膝盖微屈 哑铃贴大腿下滑 腰不弓。腰有任何不适=立即停 下周再试。"),
    @("高位下拉","腰不参与 放心做","3组","12-15次","轻-中重量 RPE6-7","60s","上半身力量——腰痛期间的背部训练","【要点】先沉肩再拉 杆拉到锁骨。")
)
$r=Write-TrainingBlock $ws2 $r "训练A  下肢+核心测试（第1周 每周2次）" "死虫式+鸟狗式+猫式+臀桥测试+深蹲测试+RDL测试+高位下拉" "45-50" $actsA1

# 训练B: 上肢+稳定性 (第1周)
SetC $ws2 $r 1 "训练B  上肢+稳定性（第1周测试版）" $true 14 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$actsB1=@(
    @("死虫式+鸟狗式","必做核心激活","各2组","各10次/侧","自重 慢速","30s","每节课的核心激活——护腰第一防线","【要点】同训练A。"),
    @("坐姿划船","腰不参与 放心做","3组","12-15次","轻-中重量 RPE6-7","60s","水平拉——中背部+菱形肌","【要点】身体不后仰 肩胛往中间夹。越野跑控臂——长途时上半身稳定。"),
    @("高位下拉(宽握)","腰不参与","3组","12-15次","RPE6-7","60s","宽握=上背阔肌 越野跑摆臂的稳定基础","【要点】先沉肩再拉。"),
    @("哑铃推肩","坐姿 腰有靠背 放心做","3组","10-12次","轻哑铃4-6kg RPE6-7","60s","肩部力量——越野跑长时间摆臂不酸","【要点】下背贴靠背 不反弓。"),
    @("面拉(绳索)","肩后束+肩袖健康","3组","15次","轻重量","45s","肩后束——和推肩互补","【要点】拉向额头 肘外展。"),
    @("平板支撑","核心耐力","3组","20-45秒","自重","30s","核心耐力=越野跑下坡时的身体控制","【要点】手肘正下方撑地 腰不塌。如果腰痛→改做跪姿平板。")
)
$r=Write-TrainingBlock $ws2 $r "训练B  上肢+稳定性测试（第1周）" "核心激活+坐姿划船+高位下拉+推肩+面拉+平板支撑" "40-45" $actsB1

# 第2-3周 进阶
$r++
SetC $ws2 $r 1 "第2-3周——重建+进阶（动作和第1周相同 重量和难度递增）" $true 14 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "第1周(测试)" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "第2周(适应)" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "第3周(进阶)" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("死虫式+鸟狗式","测试 2组","3组 如果腰不痛","3组 可尝试负重(手握轻哑铃)","2组 自重"),
    @("自重臀桥→哑铃臀桥","测试 2组自重","3组 持哑铃6-10kg","3组 哑铃10-14kg RPE7-8","2组 自重"),
    @("自重深蹲→高脚杯深蹲","测试 2组自重","3组 持哑铃6-10kg(腰不痛才加)","3组 哑铃10-14kg","2组 自重"),
    @("RDL(轻哑铃)","测试 最轻哑铃","3组 哑铃8-12kg(腰不疼才做)","3组 哑铃12-16kg","跳过——腰休息"),
    @("高位下拉","测试 轻-中","3-4组 +1片","4组 再+1片","2组 弹力带"),
    @("坐姿划船","测试 轻-中","3-4组 +1片","4组 再+1片","2组 弹力带"),
    @("平板支撑/侧平板","测试 20-45s","30s-60s","45s-90s 加侧平板","30s 轻松做")
)
foreach($pw in $prog){
    SetC $ws2 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false; MC $ws2 $r 1 2
    SetC $ws2 $r 3 $pw[1] $false 9 0 0 $false
    SetC $ws2 $r 4 $pw[2] $false 9 0 0 $false; MC $ws2 $r 4 5
    SetC $ws2 $r 6 $pw[3] $false 9 0 0 $false; MC $ws2 $r 6 7
    SetC $ws2 $r 8 $pw[4] $false 9 0 0 $false; MC $ws2 $r 8 9
    $ws2.Range("A$($r):I$($r)").RowHeight=28; SB $ws2 $r 1 9; $r++
}

# 越野专项 C
$r++
SetC $ws2 $r 1 "训练C  越野跑专项（第2-3周 每周1次 替代一次有氧日）" $true 14 $white $softBlue $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$actsC=@(
    @("上坡模拟:跑步机爬坡走","越野专项 上坡推力","1组","15-20min","坡度10-15% 速度4.5-5.5km/h","—","上坡=臀大肌+股四头 越野跑的核心场景","【要点】身体微前倾 脚步轻盈。不要扶扶手——用臀和腿发力。腰不痛才做高坡度。"),
    @("下坡模拟:倒退走+侧向走","越野专项 下坡控制","各1组","各5min","倒退走=跑步机坡度5% 侧向走=弹力带绑膝","—","下坡=股四头离心+核心抗扭转 越野跑者最需要的专项力量","【要点】倒退走=小步慢速 用大腿前侧控制。侧向走=膝盖微屈 臀中肌发力。"),
    @("单腿臀桥","臀中肌+单腿稳定","3组","12次/侧","自重或持轻哑铃","45s","越野跑=每步都是单腿着地 臀中肌是稳定器","【要点】单腿架凳 另一腿抬起。支撑腿臀推起。腰不受力——放心做。"),
    @("农夫行走","核心抗扭转+握力+斜方","3组","30-40m","大重量哑铃 单边16-24kg(量力而行)","60s","越野跑长途=核心持续收紧 保持身体不晃","【要点】躯干正直不侧倾 核心全程收紧。腰不痛才做。"),
    @("HIIT可选:爬坡冲刺","越野专项 心肺+力量","4-6轮","30s冲/60s恢复","跑步机坡度8-10% 冲刺速度","60s恢复","模拟越野跑的上坡冲刺——但腰不痛才做","  腰有不适或第1周→跳过HIIT。第3周腰不痛时可以加入。")
)
$r=Write-TrainingBlock $ws2 $r "训练C  越野跑专项（第2-3周 每周1次）" "上坡模拟+下坡模拟+单腿臀桥+农夫行走+可选HIIT" "45-55" $actsC

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18; $ws2.Range("C:C").ColumnWidth=18
$ws2.Range("D:D").ColumnWidth=8; $ws2.Range("E:E").ColumnWidth=10; $ws2.Range("F:F").ColumnWidth=16
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=22; $ws2.Range("I:I").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进+腰痛决策树
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进决策树"

SetC $ws3 1 1 "腰痛决策树 + 渐进规则" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34
SetC $ws3 2 1 "腰痛优先——任何时候腰不舒服 降级。这是本方案最重要的规则。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 9

# Back pain decision tree
$r=4
SetC $ws3 $r 1 "腰痛决策树——每次训练前必看" $true 14 $white $accentRed $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

SetC $ws3 $r 1 "腰的感觉" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "今天做什么" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$backTree=@(
    @("完全无不适","  绿灯——安全训练","正常训练 按计划走。深蹲和RDL可以做 但用保守重量。"),
    @("轻微酸胀/紧绷感","  黄灯——可以练 但保守","可以做死虫式+鸟狗式+臀桥+高位下拉+坐姿划船。深蹲和RDL=跳过或用自重。"),
    @("动作过程中出现锐痛/刺痛","  红灯——立即停","停当前动作。换死虫式/臀桥等零腰压动作。如果持续疼→当天训练结束。"),
    @("早上起床腰就疼","  红灯——当日不练下肢","只做上肢(高位下拉+划船+推肩+面拉)+死虫式。不碰深蹲/RDL/农夫行走。"),
    @("连续3天腰不痛","  可以尝试进阶","下次训练可以尝试第2周或第3周的重量。但腰有任何不适→退回第1周。"),
    @("第3周末 腰仍频繁疼痛","  恢复不够——暂不减载","延长第2周的训练到第3-4周。不要急着加重——腰痛完全消失后再谈进阶。")
)
foreach($bt in $backTree){
    SetC $ws3 $r 1 $bt[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $bt[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $bt[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=28; SB $ws3 $r 1 9; $r++
}

# Progression rules
$r+=2
SetC $ws3 $r 1 "渐进规则——腰痛不痛时的正常进阶" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 3 9; $r++

$progressRules=@(
    @("第1周测试 某个动作完全不疼","第2周可以+2-4kg或+1组——但保守。先加组数 再加重。"),
    @("第2周 所有动作腰不疼 且轻松完成","第3周进入进阶。深蹲和RDL可以尝试加重量。"),
    @("第3周 重量加了 腰仍然不疼","  最好的信号。第4周减载后 下周期可以继续进阶。"),
    @("体重每周稳定掉0.5-1kg","  减重是最好的越野跑训练。继续保持——饮食+有氧是主力。"),
    @("越野跑5km计时 比4周前快","  力量和减重在起作用。继续这个方向。")
)
foreach($prr in $progressRules){
    SetC $ws3 $r 1 $prr[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $prr[1] $false 10 0 0 $true; MC $ws3 $r 3 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=30; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "腰痛康复  热身  主动恢复  减载周" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

# Back pain rehab
$r=3
SetC $ws4 $r 1 "腰痛康复——每次训练前必做 每天在家也可以做" $true 14 $white $accentGreen $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "频率" $true 10 $white $headerBlue $false; $r++

$rehab=@(
    @("死虫式","仰卧 手举腿抬90度 对侧手脚同步下放收回 下背贴死地面","每次训练前+每天在家2-3组"),
    @("鸟狗式","四足跪姿 对侧手脚延伸 身体一条线 骨盆不旋转","同上"),
    @("猫式伸展","四足跪姿 吸抬头塌腰 呼弓背收腹 逐节活动脊柱","腰不痛时每天早晚各一组"),
    @("臀桥","脚跟着地 顶峰夹臀2秒 腰完全不受力","每次训练前+每天在家"),
    @("泡沫轴-臀+腘绳","坐姿滚大腿后侧 侧卧滚臀 找到酸痛点停15秒","训练后+每天睡前"),
    @("游泳(康复性)","低强度自由泳/仰泳 水浮力=零腰压有氧","腰痛期间的有氧首选——比跑步安全")
)
foreach($rh in $rehab){
    SetC $ws4 $r 1 $rh[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $rh[1] $false 10 0 0 $true; MC $ws4 $r 2 4
    SetC $ws4 $r 5 $rh[2] $false 10 0 0 $false; MC $ws4 $r 5 9
    $ws4.Range("A$($r):I$($r)").RowHeight=28; SB $ws4 $r 1 9; $r++
}

# Warmup
$r+=2
SetC $ws4 $r 1 "训练前热身（必做 含腰痛评估）" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws4 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "时间/次数" $true 9 $white $headerBlue $false
SetC $ws4 $r 6 "目的" $true 9 $white $headerBlue $false; MC $ws4 $r 6 9; $r++

$wu=@(
    @("1","死虫式(核心激活)","每侧8次","腰痛康复第一动作"),
    @("2","鸟狗式(核心稳定)","每侧8次","保护腰椎的底层防线"),
    @("3","猫式伸展","10次 慢节奏","脊柱活动——腰不痛时做"),
    @("4","自重臀桥","15次","激活臀大肌 零腰压"),
    @("5","摆腿(前后+左右)","各10次/侧","髋关节活动度"),
    @("6","第一个动作轻重量试做","1组x10次","让身体知觉今天的状态")
)
foreach($w in $wu){
    SetC $ws4 $r 1 $w[0] $false 9 0 0 $false
    SetC $ws4 $r 2 $w[1] $true 9 $headerBlue 0 $false; MC $ws4 $r 2 4
    SetC $ws4 $r 5 $w[2] $false 9 0 0 $false
    SetC $ws4 $r 6 $w[3] $false 9 0 0 $true; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=24; SB $ws4 $r 1 9; $r++
}

# Deload
$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 腰充分休息" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 9; $r++

$deload=@(
    @("哑铃臀桥","自重臀桥+弹力带","2组x15次"),
    @("高脚杯深蹲","自重深蹲(扶墙)","2组x12次"),
    @("哑铃RDL","跳过——腰休息",""),
    @("高位下拉","弹力带高位下拉(单臂)","2组x10次/侧"),
    @("坐姿划船","弹力带划船","2组x12次"),
    @("推肩","弹力带推肩","2组x12次"),
    @("越野专项","低强度休闲骑/游泳/散步","20-30min 心率<130"),
    @("第4周心态","腰充分休息。身体在偷偷变强。","")
)
foreach($dl in $deload){
    SetC $ws4 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $dl[2] $false 10 0 0 $true; MC $ws4 $r 4 9
    $ws4.Range("A$($r):I$($r)").RowHeight=26; SB $ws4 $r 1 9; $r++
}

$ws4.Range("A:A").ColumnWidth=22; $ws4.Range("B:I").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食+规则
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与规则"

SetC $ws5 1 1 "减重饮食  越野补给  PMS/经期  判断规则" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

# Diet
$r=3
SetC $ws5 $r 1 "减重饮食——比力量训练更重要的4周任务" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "每减1kg=越野跑时膝盖每步少承受4kg。减10kg=膝盖每步少40kg。你是耐力运动员——饮食控制+保持有氧=这4周最大的进步来源。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

SetC $ws5 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "目标" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "怎么吃" $true 10 $white $headerBlue $false; MC $ws5 $r 4 9; $r++

$nutri=@(
    @("总热量","1700-1900 kcal","比维持少400-600——每周掉0.5-1kg。你有氧消耗大 不要吃太少——低于1500会掉肌肉+影响恢复。"),
    @("蛋白质","120-140g (1.5-1.8g/kg)","每餐有蛋白质：早2蛋+奶 午晚各一份肉/鱼/豆腐。训练后蛋白粉。减重过程中蛋白质不能少——保护肌肉。"),
    @("碳水","180-220g","三餐每餐一拳。训练前后集中吃碳水。有氧日碳水多20-30g——你是耐力运动员 碳水不是敌人。"),
    @("脂肪","40-50g","控油 不吃油炸。坚果一天一小把。"),
    @("饮水","3L/天","水喝够代谢才运转。腰痛期间多喝水——帮助组织修复。"),
    @("蔬菜","不限量","每餐占一半。先吃菜再吃肉最后吃饭——这个顺序掉秤更快。")
)
foreach($n in $nutri){
    SetC $ws5 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $n[2] $false 10 0 0 $true; MC $ws5 $r 4 9
    $ws5.Range("A$($r):I$($r)").RowHeight=34; SB $ws5 $r 1 9; $r++
}

#越野补给
$r++
SetC $ws5 $r 1 "越野跑补给——训练和比赛时" $true 12 $headerBlue 0 $true; MC $ws5 $r 1 9; $r++

SetC $ws5 $r 1 "训练<1h：正常吃 不需额外补给。训练1-2h：每小时运动饮料/能量胶。训练>2h/50km备赛：每小时30-60g碳水+电解质 赛中补给你是Pro 应该有自己的习惯——保持就好。赛后30min内：碳水+蛋白(恢复速度翻倍)。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=28; $r++

# PMS
$r+=2
SetC $ws5 $r 1 "经期前后" $true 14 $white $softPink $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$softPink
$r++

SetC $ws5 $r 1 "经前一周：体重涨1-3kg=正常水肿。训练降强度 腰更要注意——经前韧带松弛 更容易受伤。经期1-2天：休息或游泳(低强度)。第3天恢复。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=28; $r++

# Rules
$r+=2
SetC $ws5 $r 1 "身体信号" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$rules=@(
    @("腰锐痛/刺痛——任何动作触发的","立即停。当天不练。连续2天还疼→看医生。你是运动员 知道硬撑的代价。"),
    @("腰酸胀/紧绷——但不疼","可以练 但降级。死虫式+臀桥+高位下拉=安全。RDL和深蹲=跳过。"),
    @("连续3天腰不痛","可以尝试下一周的进阶重量。但一旦有不适 退回。"),
    @("体重一周没掉","先看是不是经前。不是的话检查饮食——油脂和隐性碳水是罪魁祸首。"),
    @("训练+有氧后极度疲劳","你在减重+恢复伤病+训练——三件事一起做 身体负荷大。这周多睡1小时。"),
    @("腰痛完全消失+力量在涨","  最好的信号。第4周后可以进入更系统的进阶周期。")
)
foreach($r3 in $rules){
    SetC $ws5 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $r3[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=28; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用+复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  4周复盘  备赛提醒" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

# Simplified
$r=3
SetC $ws6 $r 1 "腰痛/加班/来例假——降级方案" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "本来" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws6 $r 4 "底线" $true 10 $white $headerBlue $false; MC $ws6 $r 4 9; $r++

$fallback=@(
    @("3次/周力量","2次/周——保留核心激活+上肢。下肢腰痛跳过。","做了核心激活=今天算完成"),
    @("今天腰不舒服","只做死虫式+鸟狗式+臀桥+高位下拉+拉伸","保护腰——今天不冒险"),
    @("来例假前3天+腰痛","游泳或完全休息","双重压力——身体需要歇"),
    @("这周饮食乱了","明天重新开始 不等周一","一顿吃多不会胖——连续一周才会")
)
foreach($fb in $fallback){
    SetC $ws6 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false; MC $ws6 $r 1 2
    SetC $ws6 $r 3 $fb[1] $false 10 0 0 $false
    SetC $ws6 $r 5 $fb[2] $false 10 0 0 $true; MC $ws6 $r 5 9
    $ws6.Range("A$($r):I$($r)").RowHeight=24; SB $ws6 $r 1 9; $r++
}

# Review
$r+=2
SetC $ws6 $r 1 "4周后复盘——三个核心指标" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 6 9; $r++

$review=@(
    @("腰痛恢复","比4周前减轻了多少？能做深蹲和RDL了吗？","完全消失=下周期正式进阶。仍疼=继续康复 不急。"),
    @("体重+腰围","掉了2-4kg？腰围小了？","这是4周最大的进步——比任何力量数字都重要。"),
    @("越野跑测试","5km比4周前快了吗？爬坡轻松了吗？","快了=力量和减重在起作用。下周期加越野专项频率。")
)
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 5
    SetC $ws6 $r 6 $rv[2] $false 10 0 0 $true; MC $ws6 $r 6 9
    $ws6.Range("A$($r):I$($r)").RowHeight=28; SB $ws6 $r 1 9; $r++
}

# Race prep
$r+=2
SetC $ws6 $r 1 "50km越野跑备赛提醒" $true 14 $white $softBlue $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$race=@(
    "你完赛过铁三——身体的感觉你是知道的。50km越野跑的消耗 ≈ 铁三的全过程 但地形更复杂。",
    "赛前2周：停止加重力量训练。维持轻量 不冲重量。越野跑训练减量至平时60-70%。",
    "赛前1周：力量训练减至1次(只做核心激活+上肢轻量)。越野跑只做短距离(5-10km)轻松跑。",
    "赛前3天：完全停力量 只拉伸+核心激活。碳水加载(比平时多吃半碗饭)。",
    "比赛日：你完赛过铁三——你知道怎么做。注意下坡时核心收紧——腰痛期间下坡是风险点。",
    "赛后：休息3-5天 散步+游泳。不碰器械。你身体需要修复微创伤——训练不会让你变强 恢复才会。"
)
foreach($rc in $race){
    SetC $ws6 $r 1 $rc $false 10 0 0 $true; MC $ws6 $r 1 9
    $ws6.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws6 $r 1 "你完赛铁三的那天哭了——那不是脆弱的眼泪 是Pro的勋章。这次50km越野跑 带着更轻的身体和更强壮的核心去。加油。" $false 10 $softPink 0 $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=28

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\熊\熊_减重越野跑方案_V1.xlsx"
New-Item -ItemType Directory -Force -Path "D:\Codex\members\熊" | Out-Null
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
