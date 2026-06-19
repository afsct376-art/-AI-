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

SetC $ws0 1 1 "熊  减重+越野跑 力量方案 V2" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 90kg  |  铁三Pro  减重+50km越野备赛  腰痛恢复期  V2=FMS筛查+下坡控制+核心阶梯" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  V2 升级了什么——从'安全'到'安全+高效'" $true 16 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  FMS动作筛查","三个标准动作——深蹲筛查+单腿站立+俯卧撑——量化腰痛恢复进度。每次训练前做 分数写在记录表上。"),
    @("  下坡控制专项","离心深蹲+单腿离心下蹲+倒退走+侧向下台阶+单侧农夫行走——越野跑下坡比上坡更容易受伤 这部分力量必须专项练。"),
    @("  核心稳定进阶阶梯","4个级别：基础→进阶→高阶→越野跑专项。每个级别有无痛完成的标准——像升级打怪一样训练核心。"),
    @("  碳水循环","训练日碳水200-220g 非训练日160-180g。你是耐力运动员——碳水不是敌人 但可以聪明地吃。"),
    @("  固定3次/周","第1周也从3次开始——腰痛时降级但不跳过。你是铁三Pro 3次/周是最低有效剂量。"),
    @("  力量+越野跑衔接","先力量后越野跑=脂肪供能更好。方案A/B/C三选一 看当天状态。")
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  三件事的优先级——不是同时做 是有先后" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$priority=@(
    @("第一优先","腰痛恢复——FMS筛查是每周的标尺。三个动作分数在涨=腰在恢复。任何时候不舒服 训练降级或跳过。"),
    @("第二优先","减重——每减1kg 越野跑时膝盖少承受4kg冲击力。碳水循环+保持有氧=减重主力。"),
    @("第三优先","50km越野跑备赛——下坡控制专项+核心抗扭转+上坡推力。铁三有氧基础已够 力量补短板。")
)
foreach($pr in $priority){
    SetC $ws0 $row 1 $pr[0] $true 10 $softBlue 0 $false
    SetC $ws0 $row 2 $pr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=28; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览——固定3次/周" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "周一训练A" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "周三训练B" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "周五训练C" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","FMS筛查+适应","下肢+核心+FMS","上肢+稳定性+FMS","越野专项+下坡控制入门"),
    @("2","重建力量基础","下肢+核心 加重","上肢+稳定性 加重","越野专项+下坡控制 加难度"),
    @("3","力量进阶","下肢+核心 进阶重量","上肢+稳定性 进阶","越野专项+下坡控制 高阶"),
    @("4","减载+测试","轻量 换动作 FMS对比","轻量 换动作","低强度有氧 5km计时")
)
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  力量+越野跑同天怎么排——V2新增" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$combo=@(
    @("方案A(推荐)","力量训练  休息10min  越野跑","力量先消耗糖原 越野跑时身体更依赖脂肪供能——减脂效果更好。注意：疲劳状态下跑步 腰更容易代偿——时刻关注腰的感觉。"),
    @("方案B(备选)","越野跑  休息10min  力量训练","越野跑先激活神经系统 力量时神经驱动更高——适合第3周进阶期。注意：跑完做力量 动作质量是关键——疲劳时容易变形。"),
    @("方案C(不推荐)","上午力量  下午越野跑","身体没有完整恢复时间 两练质量都打折扣。除非时间确实不允许 否则不用。")
)
foreach($cb in $combo){
    SetC $ws0 $row 1 $cb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $cb[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=34; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+腰痛约束+三件事优先级"," Sheet 方案概览"),
    @("训练A(下肢+核心) B(上肢+稳定性) C(越野+下坡)"," Sheet 训练计划"),
    @("FMS动作筛查+腰痛决策树+核心阶梯"," Sheet FMS与核心"),
    @("热身+主动恢复+减载+腰痛康复"," Sheet 恢复策略"),
    @("碳水循环+越野补给+判断规则"," Sheet 饮食与规则"),
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

SetC $ws 1 1 "熊  减重+越野跑 力量方案概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 9; $ws.Range("A1:I1").RowHeight=36

SetC $ws 2 1 "女 25-30岁 160cm/90kg BMI 35.2 体脂30-35% 铁三完赛(Pro级) 越野跑备赛 骑车被撞腰痛恢复中 FMS筛查+下坡控制+核心阶梯" $false 11 $gold $darkBg $true
MC $ws 2 1 9; $ws.Range("A2:I2").Font.Color=$white; $ws.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览 V2" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 9

$ov=@(
    @("训练频率","固定3次/周(周一/三/五) 腰痛时降级不跳过","力量服务越野跑——不是替代"),
    @("核心方法","FMS动作筛查 下坡控制专项 核心进阶阶梯 碳水循环 力量+越野跑衔接","6大升级从'安全'到'安全+高效'"),
    @("周期结构","第1周=FMS筛查+适应 第2周=重建 第3周=进阶 第4周=减载+测试","腰痛优先——任何时候疼就降级"),
    @("核心约束","腰痛恢复期——所有动作以无痛为第一原则","下背保护+核心激活+离心控制"),
    @("方案特点","减重+越野跑+腰痛保护 三合一。铁三Pro的体能方案——不是普通人的入门计划")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 6
    SetC $ws $r 7 $o[2] $false 10 0 0 $true; MC $ws $r 7 9
    SB $ws $r 1 9; $r++
}

# Three priorities
$r++
SetC $ws $r 1 "三件事的优先级——为什么这个顺序" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$why=@(
    "你的情况特殊：能完赛铁三=有氧基础极好。但90kg的体重+腰痛=不能用常规运动员方案。",
    "",
    "第一优先：腰痛恢复。骑车被撞后的腰痛——不是肌肉酸痛 是创伤恢复。V2引入FMS动作筛查——三个标准动作 每次训练前测 分数写在记录表上。分数在涨=腰在恢复。",
    "",
    "第二优先：减重。每减1kg=越野跑时膝盖每步少承受4kg。减10kg=膝盖每步少40kg。这比任何力量训练对越野跑的帮助都大。V2用碳水循环：训练日碳水200-220g 非训练日160-180g。你是有氧Pro 碳水不是敌人——但要聪明地吃。",
    "",
    "第三优先：越野跑专项。下坡控制是V2的重头戏——越野跑下坡比上坡更容易受伤 尤其是腰痛恢复期。离心深蹲+单腿离心下蹲+侧向下台阶——这些是上坡训练给不了你的。"
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
    " 谨慎：深蹲(FMS筛查通过才加重) RDL(最轻哑铃 全程直背 腰有不适立即停)。",
    " 安全：臀推/臀桥(零腰压) 核心稳定性训练(死虫式/鸟狗式/平板支撑) 高位下拉 坐姿划船。",
    " 康复：死虫式+鸟狗式(每节课热身必做)。猫式伸展(腰不痛时做)。游泳(腰痛期有氧首选)。",
    "  停止信号：训练中腰出现锐痛→立即停 当天不练。FMS分数下降→退回上一周的训练内容。"
)
foreach($ba in $back){
    SetC $ws $r 1 $ba $false 10 0 0 $true; MC $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight=24; $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$pr=@(
    @("1","腰痛优先于一切","FMS分数是每周的标尺。分数在涨=可以进阶。分数下降=退回。你是Pro——用数据跟踪恢复 不凭感觉。"),
    @("2","减重=最好的越野跑训练","碳水循环+保持有氧。每减1kg=膝盖每步少4kg。减10kg=每步少40kg。"),
    @("3","下坡控制比上坡推力更重要","越野跑下坡的膝盖冲击力是体重的4-6倍。腰痛恢复期 下坡控制专项训练是你的护身符。")
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
# SHEET 2: 训练计划 V2
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划 V2——第1-4周 力量+越野专项+下坡控制" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "每周固定3次(周一训练A/周三训练B/周五训练C)。腰痛时降级不跳过。FMS筛查在训练A和B的热身前做(见Sheet3)。第4周减载换动作。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 训练A
$r=4
SetC $ws2 $r 1 "第1周 训练A  下肢+核心（FMS筛查+适应）" $true 14 $white $accentGreen $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws2 $r 1 "训练前先做FMS筛查(深蹲+单腿站立+俯卧撑 见Sheet3) 记录分数。第1周=轻量适应 腰不痛的动作记录下来。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=22; $r++

$actsA=@(
    @("FMS筛查(必做!)","腰痛量化标尺","—","3项","见Sheet3 FMS动作筛查","—","第1周建基准 每周对比","【要点】训练前做 记录分数。分数涨=腰在恢复。分数降=退回上周。"),
    @("死虫式+鸟狗式","核心激活 必做","各2组","各10次/侧","自重 慢速控制","30s","核心稳定=护腰第一防线","【要点】死虫式=下背贴地。鸟狗式=骨盆不旋转。"),
    @("自重臀桥","臀大肌激活 零腰压","3组","15次","自重","45s","腰完全不受力——放心做","【要点】脚跟着地 顶峰夹臀2秒。"),
    @("自重深蹲(扶墙)","FMS深蹲筛查 → 决定今天做不做","2组","10-12次","自重","60s","FMS筛查通过=可以做。不通过=跳过。","【要点】能蹲多深蹲多深。腰不疼=记录深度。腰疼=跳过。"),
    @("轻哑铃RDL","FMS筛查通过才做","2组","10次","最轻哑铃4-6kg 全程直背","60s","腰有任何不适→立即停","【要点】膝盖微屈 哑铃贴大腿下滑。腰不弓=最重要的底线。"),
    @("高位下拉","腰不参与 放心做","3组","12-15次","轻-中重量 RPE6-7","60s","上半身力量——腰痛期安全","【要点】先沉肩再拉 杆拉到锁骨。")
)
$r=Write-TrainingBlock $ws2 $r "周一 训练A  下肢+核心" "FMS筛查+核心激活+臀桥+深蹲测试+RDL测试+高位下拉" "40-45" $actsA

# 训练B
SetC $ws2 $r 1 "第1周 训练B  上肢+稳定性" $true 14 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$actsB=@(
    @("FMS俯卧撑筛查(必做)","量化上肢推力+腰痛关系","—","1组力竭","见Sheet3 FMS","—","第1周建基准","【要点】标准俯卧撑做到力竭 记录次数。"),
    @("死虫式+鸟狗式","核心激活 必做","各3组","各10次/侧","自重","30s","每节课的必修——核心稳定=护腰","【要点】同训练A。"),
    @("坐姿划船","水平拉 腰不参与","3组","12-15次","RPE6-7","60s","中背部+菱形肌——越野跑控臂","【要点】身体不后仰 肩胛往中间夹。"),
    @("高位下拉(宽握)","垂直拉 腰不参与","3组","12-15次","RPE6-7","60s","宽握=上背阔肌——长途摆臂稳定","【要点】先沉肩再拉。"),
    @("哑铃推肩","坐姿 腰有靠背","3组","10-12次","轻哑铃4-6kg","60s","肩部力量——长时间摆臂不酸","【要点】下背贴靠背 不反弓。"),
    @("面拉(绳索)+平板支撑","肩后束+核心耐力","面拉3x15 平板3x20-45s","—","轻重量+自重","45s","肩袖健康+核心耐力","【要点】面拉拉向额头。平板腰不塌——如果腰疼改跪姿。")
)
$r=Write-TrainingBlock $ws2 $r "周三 训练B  上肢+稳定性" "FMS俯卧撑+核心激活+坐姿划船+高位下拉+推肩+面拉+平板支撑" "40-45" $actsB

# 训练C V2 (下坡控制专项)
SetC $ws2 $r 1 "第1周 训练C  越野跑专项+下坡控制入门" $true 14 $white $softBlue $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$actsC=@(
    @("离心深蹲","V2新增——下坡控制核心训练","3组","8-10次","自重 3秒下放+1秒起","60s","模拟下坡时股四头离心收缩——越野跑最需要的专项力量","【要点】下放3秒 越慢越好。起1秒。腰不痛才做——从自重开始。"),
    @("跑步机爬坡走","上坡推力","1组","12-15min","坡度10-12% 速度4.5-5.5km/h","—","上坡=臀大肌+股四头 越野跑核心场景","【要点】身体微前倾 不扶扶手。腰不痛才做高坡度。"),
    @("倒退走","下坡模拟——激活后侧链","1组","5min","跑步机 坡度5% 速度3-4km/h","—","激活臀和腘绳肌——下坡时后侧链的刹车作用","【要点】小步慢速 身体微后倾。手扶扶手保持平衡。"),
    @("单腿臀桥","臀中肌+单腿稳定","3组","12次/侧","自重或持轻哑铃","45s","越野跑=每步都是单腿着地 臀中肌是稳定器","【要点】单腿架凳 支撑腿臀推起。腰不受力。"),
    @("农夫行走(双侧)","核心抗侧屈+握力","3组","30-40m","单边16-24kg(量力)","60s","越野跑长途=核心持续收紧","【要点】躯干正直不侧倾。腰不痛才做。")
)
$r=Write-TrainingBlock $ws2 $r "周五 训练C  越野跑专项+下坡控制" "离心深蹲+爬坡走+倒退走+单腿臀桥+农夫行走" "45-50" $actsC

# 第2-3周进阶
$r++
SetC $ws2 $r 1 "第2-3周——重建+进阶（动作框架不变 重量和难度递增）" $true 14 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "第1周(适应)" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "第2周(重建)" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "第3周(进阶)" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("死虫式+鸟狗式","2组 自重","3组 自重","3组 可尝试轻弹力带","2组 自重"),
    @("自重臀桥→哑铃臀桥","2组自重","3组 持哑铃6-10kg","3组 哑铃10-14kg","2组 自重+弹力带"),
    @("自重深蹲→高脚杯深蹲","FMS通过才做","FMS通过才做 哑铃6-10kg","FMS通过 哑铃10-14kg","2组 自重"),
    @("轻哑铃RDL","最轻哑铃 测试","FMS通过 哑铃8-12kg","FMS通过 哑铃12-16kg","跳过——腰休息"),
    @("高位下拉+坐姿划船","轻-中 3组","+1片 4组","再+1片 4组","2组 弹力带"),
    @("离心深蹲(训练C)","自重 3组","+轻哑铃4-6kg 3组","+哑铃6-10kg 3组","2组 自重 更慢离心"),
    @("倒退走 侧向下台阶","倒退走入门","加侧向下台阶 每侧3x10","加单侧农夫行走","跳过——腰休息")
)
foreach($pw in $prog){
    SetC $ws2 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false; MC $ws2 $r 1 2
    SetC $ws2 $r 3 $pw[1] $false 9 0 0 $false
    SetC $ws2 $r 4 $pw[2] $false 9 0 0 $false; MC $ws2 $r 4 5
    SetC $ws2 $r 6 $pw[3] $false 9 0 0 $false; MC $ws2 $r 6 7
    SetC $ws2 $r 8 $pw[4] $false 9 0 0 $false; MC $ws2 $r 8 9
    $ws2.Range("A$($r):I$($r)").RowHeight=28; SB $ws2 $r 1 9; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18; $ws2.Range("C:C").ColumnWidth=18
$ws2.Range("D:D").ColumnWidth=8; $ws2.Range("E:E").ColumnWidth=10; $ws2.Range("F:F").ColumnWidth=14
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=22; $ws2.Range("I:I").ColumnWidth=50
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: FMS筛查+核心阶梯+腰痛决策树 (V2 重头戏)
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="FMS与核心"

SetC $ws3 1 1 "FMS动作筛查  核心进阶阶梯  腰痛决策树" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34
SetC $ws3 2 1 "V2新增核心模块。FMS量化腰痛恢复 核心阶梯让训练有进阶路径 腰痛决策树=每次训练前必看。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 9

# FMS
$r=4
SetC $ws3 $r 1 "FMS动作筛查——每次训练A和B的热身前做 记录分数" $true 14 $white $accentGreen $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "三个标准动作 覆盖下肢/单腿稳定/上肢推力三个维度。每周对比分数——分数涨=腰在恢复。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=22; $r++

SetC $ws3 $r 1 "筛查项目" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "怎么做" $true 10 $white $headerBlue $false; MC $ws3 $r 2 4
SetC $ws3 $r 5 "评分标准" $true 10 $white $headerBlue $false; MC $ws3 $r 5 7
SetC $ws3 $r 8 "训练决策" $true 10 $white $headerBlue $false; $r++

$fms=@(
    @("1.深蹲筛查","双手举过头顶 下蹲至大腿平行地面","  无痛完成=可以加重量。  有痛但完成=只自重不加重量。  无法完成/锐痛=跳过深蹲 改高脚杯深蹲。"),
    @("2.单腿站立筛查","双手叉腰 单腿站立 闭眼保持10秒","  稳定完成=可以加单腿训练(分腿蹲等)。  摇晃完成=只做双腿训练。  无法完成=优先核心稳定。"),
    @("3.俯卧撑筛查(训练B做)","标准俯卧撑 下巴贴地 做到力竭","  无痛完成=上肢正常。  腰痛但完成=改跪姿/上斜俯卧撑。  无法完成/锐痛=只用固定器械。"),
    @("FMS记录示例","第1周：深蹲= 单腿= 俯卧撑=力竭8次","第2周：深蹲= 单腿= 俯卧撑=力竭12次——腰在恢复！")
)
foreach($fm in $fms){
    SetC $ws3 $r 1 $fm[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $fm[1] $false 10 0 0 $true; MC $ws $r 2 4
    SetC $ws $r 5 $fm[2] $false 10 0 0 $true; MC $ws $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=34; SB $ws3 $r 1 9; $r++
}

# Core progression ladder
$r+=2
SetC $ws3 $r 1 "核心稳定进阶阶梯——像升级打怪一样训练核心" $true 14 $white $softPurple $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "级别" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "训练内容" $true 10 $white $headerBlue $false; MC $ws3 $r 2 4
SetC $ws3 $r 5 "目标" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "进阶条件" $true 10 $white $headerBlue $false; MC $ws3 $r 7 9; $r++

$core=@(
    @("第1级 基础(第1周)","死虫式3x10/侧 鸟狗式3x10/侧 臀桥3x15","激活核心 建立神经连接 零腰压","当前级别所有动作无痛完成→下周进阶"),
    @("第2级 进阶(第2周)","平板支撑3x30s 侧平板3x20s/侧 单腿臀桥3x12/侧","提高核心耐力——越野跑下坡的身体控制基础","平板和侧平板无腰痛→下周进阶"),
    @("第3级 高阶(第3周)","农夫行走双侧3x30s 死虫式+弹力带3x10/侧 鸟狗式+弹力带3x10/侧","核心抗旋转+抗侧屈——越野跑长途的稳定基础","农夫行走腰不痛+弹力带动作稳定→下周进阶"),
    @("第4级 越野跑专项(第4周后)","单腿平板3x20s/侧 登山者3x20/侧 药球旋转抛3x10/侧(腰不痛才做)","模拟越野跑时的核心需求——动态稳定","第4周减载周结束后 下周期可用")
)
foreach($cl in $core){
    SetC $ws3 $r 1 $cl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $cl[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $cl[2] $false 10 0 0 $false
    SetC $ws3 $r 7 $cl[3] $false 10 0 0 $true; MC $ws3 $r 7 9
    $ws3.Range("A$($r):I$($r)").RowHeight=42; SB $ws3 $r 1 9; $r++
}

SetC $ws3 $r 1 "  任何时候 当前级别有动作引起腰痛→退回上一级。核心训练不是比谁级别高 是比谁的腰更安全。" $false 10 $softPurple 0 $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=24; $r++

# Back pain decision tree
$r+=2
SetC $ws3 $r 1 "腰痛决策树——每次训练前必看" $true 14 $white $accentRed $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

SetC $ws3 $r 1 "腰的感觉+FMS分数" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "今天做什么" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$backTree=@(
    @("FMS三项全 +腰无不适","  绿灯——安全训练","正常训练 按计划走。深蹲和RDL可以做 但保守重量。核心可以尝试下一级。"),
    @("FMS有 有 腰轻微酸胀","  黄灯——可以练 保守","可以做核心+上肢+臀桥。深蹲和RDL=跳过或用自重。下坡控制=只做倒退走。"),
    @("FMS有 动作中锐痛","  红灯——立即停","停当前动作。只做零腰压动作(死虫式+臀桥+下拉)。持续疼→今天结束。"),
    @("早上起床腰就疼 或 FMS分数下降","  红灯——退回到上周","只做上肢(下拉+划船+推肩+面拉)+核心基础级。不碰下肢和越野跑。"),
    @("FMS分数连续2周在涨","  腰在恢复！","可以尝试下周进阶。但第3周仍要保守——腰比训练计划重要。"),
    @("第3周末 FMS仍有 ","  恢复不够 暂不进阶","延长第2周训练到第3-4周。不要急着加重——腰痛完全消失后再谈进阶。")
)
foreach($bt in $backTree){
    SetC $ws3 $r 1 $bt[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $bt[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $bt[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=28; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=28; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "腰痛康复  热身  主动恢复  减载周" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

$r=3
SetC $ws4 $r 1 "腰痛康复动作——每次训练前必做 每天在家也可以做" $true 14 $white $accentGreen $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "频率" $true 10 $white $headerBlue $false; $r++

$rehab=@(
    @("死虫式","仰卧 手举腿抬90度 对侧手脚同步下放收回 下背贴死地面","每次训练前+每天在家2-3组"),
    @("鸟狗式","四足跪姿 对侧手脚延伸 身体一条线 骨盆不旋转","同上"),
    @("猫式伸展","四足跪姿 吸抬头塌腰 呼弓背收腹 逐节活动","腰不痛时每天早晚各一组"),
    @("臀桥","脚跟着地 顶峰夹臀2秒 腰完全不受力","每次训练前+每天在家"),
    @("泡沫轴-臀+腘绳","坐姿滚大腿后侧 侧卧滚臀 找到酸痛点停15秒","训练后+每天睡前"),
    @("游泳(康复性)","低强度自由泳/仰泳 水浮力=零腰压有氧","腰痛期有氧首选——比跑步安全")
)
foreach($rh in $rehab){
    SetC $ws4 $r 1 $rh[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $rh[1] $false 10 0 0 $true; MC $ws4 $r 2 4
    SetC $ws4 $r 5 $rh[2] $false 10 0 0 $false; MC $ws4 $r 5 9
    $ws4.Range("A$($r):I$($r)").RowHeight=28; SB $ws4 $r 1 9; $r++
}

# Warmup
$r+=2
SetC $ws4 $r 1 "训练前热身（必做 含FMS筛查+核心激活）" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws4 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "时间/次数" $true 9 $white $headerBlue $false
SetC $ws4 $r 6 "目的" $true 9 $white $headerBlue $false; MC $ws4 $r 6 9; $r++

$wu=@(
    @("1","FMS筛查(训练A/B)","1-2min","量化腰痛 记录分数"),
    @("2","死虫式+鸟狗式","各8次/侧","核心激活——护腰第一防线"),
    @("3","猫式伸展","10次慢节奏","脊柱活动 腰不痛时做"),
    @("4","自重臀桥","15次","激活臀大肌 零腰压"),
    @("5","摆腿(前后+左右)","各10次/侧","髋关节活动度"),
    @("6","第一个动作轻重量试做","1组x10次","知觉今天的状态")
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
    @("离心深蹲","自重离心深蹲 更慢(5秒下放)","2组x6-8次"),
    @("高位下拉+划船","弹力带高位下拉+划船","2组x10-12次"),
    @("越野跑","低强度休闲骑/游泳/散步","20-30min 心率<130"),
    @("FMS","第4周再做一次 和第1周对比","看分数变化")
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
# SHEET 5: 饮食+碳水循环+规则
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与规则"

SetC $ws5 1 1 "碳水循环  减重饮食  越野补给  判断规则" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

# Carb cycling V2
$r=3
SetC $ws5 $r 1 "碳水循环——训练日和非训练日碳水不同 V2新增" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你是耐力运动员——碳水不是敌人。碳水循环让你在训练日有足够能量 休息日加速减脂。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=22; $r++

SetC $ws5 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "训练日(力量+有氧)" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "非训练日(休息/只有有氧)" $true 10 $white $headerBlue $false
SetC $ws5 $r 7 "怎么吃" $true 10 $white $headerBlue $false; $r++

$carb=@(
    @("总热量","1800-2000 kcal","1600-1700 kcal","训练日多200-300——支撑训练质量"),
    @("碳水","200-220g","160-180g","三餐各一拳。训练前后集中吃碳水"),
    @("蛋白质","120-140g","120-140g(不变)","每餐有蛋白质 训练后蛋白粉"),
    @("脂肪","40-50g","40-50g(不变)","控油 不吃油炸"),
    @("训练前后窗口","练前2h:碳水30-40g+蛋白15-20g`n练后1h内:碳水40-50g+蛋白25-30g","不需额外加餐","训练前后窗口是碳水循环的关键——训练日吃 非训练日不吃"),
    @("饮水","3L/天","3L/天","腰痛期间多喝水——帮助组织修复")
)
foreach($cb in $carb){
    SetC $ws5 $r 1 $cb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $cb[1] $false 10 0 0 $true; MC $ws5 $r 2 3
    SetC $ws5 $r 4 $cb[2] $false 10 0 0 $true; MC $ws5 $r 4 6
    SetC $ws5 $r 7 $cb[3] $false 10 0 0 $true; MC $ws5 $r 7 9
    $ws5.Range("A$($r):I$($r)").RowHeight=30; SB $ws5 $r 1 9; $r++
}

#越野补给
$r++
SetC $ws5 $r 1 "越野跑补给" $true 12 $headerBlue 0 $true; MC $ws5 $r 1 9; $r++
SetC $ws5 $r 1 "训练<1h:不需额外补给。训练1-2h:每小时运动饮料/能量胶。训练>2h/50km备赛:每小时30-60g碳水+电解质。赛后30min:碳水+蛋白——恢复速度翻倍。你是Pro 应该有自己的习惯——保持。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

# PMS
$r+=2
SetC $ws5 $r 1 "经期前后" $true 14 $white $softPink $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$softPink
$r++

SetC $ws5 $r 1 "经前一周：体重涨1-3kg=正常水肿。腰更要注意——经前韧带松弛更容易受伤。训练降强度。经期1-2天:休息或游泳(低强度)。第3天恢复。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

# Rules
$r+=2
SetC $ws5 $r 1 "身体信号" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$rules=@(
    @("FMS分数下降","退回上周训练内容——腰在告诉你'今天不行'"),
    @("腰锐痛——任何动作触发","立即停。当天不练。连续2天→看医生。"),
    @("FMS分数连续2周在涨","  可以进阶。下周期尝试核心下一级+深蹲加重量。"),
    @("训练+有氧后极度疲劳","减重+恢复伤病+训练——三线作战。这周多睡1小时。"),
    @("体重一周没掉","先看是不是经前。不是的话检查油脂和隐性碳水。"),
    @("腰痛完全消失+FMS三项全 ","  最好的信号。第4周后进入更系统的进阶周期。")
)
foreach($r3 in $rules){
    SetC $ws5 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $r3[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=28; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用+复盘+备赛
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "降级方案  4周复盘  备赛提醒" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

$r=3
SetC $ws6 $r 1 "腰痛/加班/来例假——降级方案" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "本来" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws6 $r 4 "底线" $true 10 $white $headerBlue $false; MC $ws6 $r 4 9; $r++

$fallback=@(
    @("训练A(下肢+核心)","只做核心激活+臀桥+高位下拉——跳过深蹲和RDL","做了核心激活=今天算完成"),
    @("训练B(上肢+稳定性)","只做核心激活+高位下拉+坐姿划船——跳过俯卧撑筛查","同上"),
    @("训练C(越野跑+下坡)","只做倒退走+单腿臀桥——跳过离心深蹲和爬坡走","同上"),
    @("来例假+腰痛","游泳或完全休息——双重压力 身体需要歇","游泳也算训练")
)
foreach($fb in $fallback){
    SetC $ws6 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false; MC $ws6 $r 1 2
    SetC $ws6 $r 3 $fb[1] $false 10 0 0 $false
    SetC $ws6 $r 5 $fb[2] $false 10 0 0 $true; MC $ws6 $r 5 9
    $ws6.Range("A$($r):I$($r)").RowHeight=24; SB $ws6 $r 1 9; $r++
}

# Review V2
$r+=2
SetC $ws6 $r 1 "4周后复盘——四个核心指标" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 6 9; $r++

$review=@(
    @("FMS分数变化","第1周vs第4周——深蹲/单腿/俯卧撑各涨了多少？","三项都 =腰恢复 下周期正式进阶。仍有 →继续康复优先。"),
    @("体重+腰围","掉了2-4kg？腰围小了？","这是4周最大的进步——比任何力量数字都重要。碳水循环在起作用。"),
    @("核心阶梯","从第几级升级到了第几级？","升级=核心在变强=腰的保护层在变厚。"),
    @("越野跑5km计时","比4周前快了吗？下坡控制有没有更稳？","快了=力量+减重+下坡控制在起作用。下周期加越野专项频率。")
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
    "赛前2周：停止加重力量。维持轻量 不冲。越野跑减至平时60-70%。",
    "赛前1周：力量1次(只做核心激活+上肢轻量)。越野跑只短距离轻松跑。",
    "赛前3天：完全停力量 只拉伸+核心激活。碳水加载(比平时多吃半碗饭)。",
    "比赛日：注意下坡时核心收紧——腰痛期间下坡是风险点。下坡控制专项训练就是为这一天准备的。",
    "赛后3-5天：散步+游泳——不碰器械。你身体需要修复微创伤。"
)
foreach($rc in $race){
    SetC $ws6 $r 1 $rc $false 10 0 0 $true; MC $ws6 $r 1 9
    $ws6.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws6 $r 1 "你完赛铁三的那天哭了——那不是脆弱的眼泪 是Pro的勋章。这次50km越野跑 带着FMS三项全  的腰 带着碳水循环减下来的体重 带着下坡控制的专项力量去。加油。" $false 10 $softPink 0 $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=28

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\熊\熊_减重越野跑方案_V2.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
