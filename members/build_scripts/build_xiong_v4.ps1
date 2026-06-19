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
# SHEET 0: 从这里开始（含腰痛决策树——会员打开第一眼）
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "熊  减重+越野跑 力量方案 V4" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 90kg  |  铁三Pro  减重+50km越野备赛  腰痛恢复" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

# V4: 腰痛决策树放在最前面——会员打开第一眼看到
$row=4
SetC $ws0 $row 1 "  今天能不能练？——看这里（每次训练前必做）" $true 16 $white $accentRed $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentRed
$row++

SetC $ws0 $row 1 "Step1: 做FMS三个筛查（5分钟）——详见右侧Sheet「FMS与核心」" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

SetC $ws0 $row 1 "Step2: 把三个分数加起来（满分15分）——找到你的分数：" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

SetC $ws0 $row 1 "总分" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "今天做什么" $true 10 $white $headerBlue $false
MC $ws0 $row 2 5; $row++

$decision=@(
    @("13-15分","正常训练 按计划走 深蹲和RDL可以做"),
    @("10-12分","降级训练 深蹲/RDL跳过 改臀桥+高脚杯 核心维持当前级"),
    @("7-9分","只做核心+上肢 下肢全跳过 核心退回上一级"),
    @("4-6分","只做核心基础级(死虫式+鸟狗式+臀桥) 力量取消 改散步/游泳"),
    @("3分以下","今天休息。连续2天都是这个分数→看医生。")
)
foreach($dc in $decision){
    SetC $ws0 $row 1 $dc[0] $true 12 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $dc[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row++
SetC $ws0 $row 1 "Step3: 对应打开Sheet「训练计划」找到今天的训练日——开始训练" $false 11 $0 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

$row+=2
SetC $ws0 $row 1 "  V4 升级对比——从'教练能看懂'到'会员能执行'" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  FMS打分 三档→五档","5分=流畅 4分=有点紧 3分=轻微酸胀 2分=明显疼但能完成 1分=锐痛/做不到。对号入座 不再靠感觉。"),
    @("  决策 主观→分数驱动","总分13-15=正常 10-12=降级 7-9=只做上肢 4-6=只做核心 <3=休息。不用判断'酸还是疼'——看分数。" ),
    @("  核心训练 动作列表→每周目标","第1周=1级基础 第2周=2级进阶 第3周=3级高阶 第4周=回1级对比。每周练什么不用想。" ),
    @("  碳水循环 数字→具体食物","训练日：吐司2片+香蕉2根+米饭2碗+红薯1个。非训练日：吐司1片+米饭1碗+粗粮1拳。照着吃。" ),
    @("  训练计划 动作表→每周执行要点","每周开头加了'本周执行要点'——本周目标+核心原则+腰安全提醒。" )
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  三件事优先级" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$priority=@(
    @("第一","腰痛恢复——FMS分数是标尺。分数涨=恢复。分数降=退回。"),
    @("第二","减重——每减1kg=膝盖每步少4kg。碳水循环+有氧。"),
    @("第三","50km越野备赛——下坡控制+核心抗扭转+上坡推力。")
)
foreach($pr in $priority){
    SetC $ws0 $row 1 $pr[0] $true 10 $softBlue 0 $false
    SetC $ws0 $row 2 $pr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览——固定3次/周（腰痛时看上面决策树降级）" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "周一A" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "周三B" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "周五C" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","适应+筛查","下肢+核心 FMS建基准","上肢+稳定 FMS建基准","越野+下坡控制入门"),
    @("2","重建力量","下肢+核心 加重 核心升2级","上肢+稳定 加重 核心升2级","越野+下坡 加难度"),
    @("3","力量进阶","下肢+核心 进阶重量 核心升3级","上肢+稳定 进阶 核心升3级","越野+下坡 高阶"),
    @("4","减载+测试","轻量换动作 FMS对比","轻量换动作","低强有氧 5km计时")
)
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

# Quick nav
$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("完整方案逻辑+腰痛约束"," Sheet 方案概览"),
    @("训练A/B/C 每周执行要点"," Sheet 训练计划"),
    @("FMS五档评分+核心阶梯+食物换算"," Sheet FMS与饮食"),
    @("热身+康复+减载+赛前减量"," Sheet 恢复与备赛"),
    @("精简版+复盘+备赛提醒"," Sheet 备用与复盘")
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
$ws1=$wb.Worksheets.Add();$ws1.Name="方案概览"

SetC $ws1 1 1 "熊  减重+越野跑  力量方案概览 V4" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "女 25-30岁 160cm/90kg BMI 35.2 体脂30-35% 铁三完赛(Pro) 越野跑备赛 骑车被撞腰痛恢复中" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览 V4" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ov=@(
    @("训练频率","固定3次/周（周一A/周三B/周五C）","腰痛时看封面决策树降级——不跳过"),
    @("核心方法","FMS五档评分+分数驱动决策+核心每周目标+碳水具体食物+训练执行要点","从'教练能看懂'到'会员能执行'"),
    @("周期结构","第1周=适应+FMS基准 第2周=重建 第3周=进阶 第4周=减载+FMS对比","FMS分数决定能不能进阶"),
    @("核心约束","腰痛恢复期——所有动作以无痛为第一原则","下背保护+核心激活+离心控制=每节课必修"),
    @("方案特点","减重+越野跑+腰痛保护 三合一。铁三Pro的体能方案——执行层面傻瓜化")
)
$r=5
foreach($o in $ov){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "为什么这样设计" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$why=@(
    "你完赛过铁三——有氧基础不是问题。但90kg体重+腰痛=不能用常规运动员方案。",
    "V4的核心创新：用FMS五档评分驱动决策。你不需要判断'腰酸还是腰疼'——把三个分数加起来 总分落在哪个区间 就知道今天练什么。",
    "每个动作都有执行要点——不是'做3组x12次' 而是'这周的目标是什么 怎么判断自己做对了'。",
    "碳水循环不用算克数——直接告诉你训练日吃什么 非训练日吃什么。照着吃。"
)
foreach($w in $why){
    SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "腰痛训练原则——每节课必修" $true 14 $white $accentRed $true
MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=26
$ws1.Range("A$($r):I$($r)").Font.Color=$white; $ws1.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

$back=@(
    " 禁止：传统硬拉 大重量深蹲 负重俯身划船 跳箱 负重跑步",
    " 谨慎：深蹲(FMS 4分以上才加重) RDL(轻哑铃 全程直背)",
    " 安全：臀推/臀桥(零腰压) 死虫式/鸟狗式 高位下拉 坐姿划船",
    " 康复：死虫式+鸟狗式(热身必做) 猫式伸展(腰不痛时) 游泳(腰痛期有氧首选)",
    "  停止：腰锐痛(FMS得1分)→当天不练。FMS总分下降→退回上周。"
)
foreach($ba in $back){
    SetC $ws1 $r 1 $ba $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$pr=@(
    @("1","腰痛优先于一切","FMS总分是标尺。分数涨=进阶 分数降=退回。数据比感觉诚实。"),
    @("2","减重=最好的越野训练","碳水循环+有氧。每减1kg=膝盖每步少4kg。"),
    @("3","下坡控制比上坡推力重要","越野跑下坡膝盖冲击=体重4-6倍。腰痛期 下坡控制=你的护身符。")
)
foreach($p in $pr){
    SetC $ws1 $r 1 $p[0] $false 10 0 0 $false
    SetC $ws1 $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws1 $r 3 $p[2] $false 10 0 0 $true; MC $ws1 $r 3 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$ws1.Range("A:A").ColumnWidth=30; $ws1.Range("B:I").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划 V4（含每周执行要点）
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划 V4——第1-4周 力量+越野专项+下坡控制" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "固定3次/周(周一A/周三B/周五C)。FMS筛查在训练前做(详见封面决策树+Sheet3)。每周围绕执行要点推进。第4周减载。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# --- 训练A 第1周 ---
$r=4
SetC $ws2 $r 1 "周一  训练A  下肢+核心（第1周  FMS筛查+适应）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "  本周执行要点：目标=测试腰的反应 不是练到累。FMS深蹲4分以上=做深蹲 3分以下=跳过改高脚杯。每个动作做完问：腰疼吗？不疼=记录 疼=跳过。核心=死虫式+鸟狗式+臀桥(零腰压安全底线)。" $false 10 $accentGreen 0 $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=36; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰安全" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsA1=@(
    @("1","FMS深蹲+单腿站立筛查(必做!)","—","见Sheet3五档评分","—","—","训练前测 填分数","双手举过头顶下蹲至大腿平行/闭眼单腿站立10秒。填分。" ),
    @("2","死虫式+鸟狗式（核心1级）","各2组","各10次/侧","自重","30s","  零腰压","下背贴地 骨盆不旋转。护腰第一防线——这周只练这级。" ),
    @("3","自重臀桥","3组","15次","自重","45s","  零腰压","脚跟着地 顶峰夹臀2秒。腰完全不受力。" ),
    @("4","自重深蹲(扶墙)","2组","10-12次","自重","60s","FMS深蹲4分+才做","能蹲多深蹲多深。腰不疼=记录深度。腰疼=跳过 改高脚杯。" ),
    @("5","轻哑铃RDL","2组","10次","4-6kg哑铃","60s","FMS总分10+才做","膝盖微屈 哑铃贴腿下滑 全程直背。腰不适=立即停。" ),
    @("6","高位下拉","3组","12-15次","轻-中 RPE6-7","60s","  无腰参与","先沉肩再拉 杆拉到锁骨。" )
)
$i=1
foreach($a in $actsA1){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=56; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# --- 训练B 第1周 ---
SetC $ws2 $r 1 "周三  训练B  上肢+稳定性（第1周  FMS筛查+适应）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "  本周执行要点：目标=测试上肢推力+腰痛关系。FMS俯卧撑5分=正常做 4分=正常 3分以下=改跪姿。所有上肢动作腰有靠背或无腰参与——安全。" $false 10 $accentGreen 0 $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=30; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰安全" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsB1=@(
    @("1","FMS俯卧撑筛查(必做!)","—","1组力竭 记录次数","自重","—","训练前测","标准俯卧撑到力竭。5分=15次+ 4分=10-14次 3分=5-9次。" ),
    @("2","死虫式+鸟狗式（核心1级）","各3组","各10次/侧","自重","30s","  零腰压","每节课必修——核心稳定=护腰。" ),
    @("3","坐姿划船","3组","12-15次","RPE6-7","60s","  无腰参与","身体不后仰 肩胛夹紧。越野跑控臂——长途稳定。" ),
    @("4","高位下拉(宽握)","3组","12-15次","RPE6-7","60s","  无腰参与","宽握=上背阔肌。长途摆臂不酸。" ),
    @("5","哑铃推肩(坐姿)","3组","10-12次","4-6kg哑铃","60s","  腰有靠背","下背贴靠背 不反弓。核心全程收紧。" ),
    @("6","面拉+平板支撑","面拉3x15 平板3x20-45s","—","轻重量/自重","45s","平板腰痛改跪姿","面拉=肩后束 平板=核心耐力 腰塌=改跪姿。" )
)
$i=1
foreach($a in $actsB1){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=56; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# --- 训练C 第1周 ---
SetC $ws2 $r 1 "周五  训练C  越野跑专项+下坡控制入门（第1周）" $true 13 $white $softBlue $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

SetC $ws2 $r 1 "  本周执行要点：下坡控制入门。离心深蹲=FMS深蹲4分+才做 从自重开始 3秒下放越慢越好。腰不痛才做爬坡走。倒退走激活臀+腘绳——下坡刹车。" $false 10 $accentGreen 0 $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=30; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰安全" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsC1=@(
    @("1","离心深蹲(下坡控制核心)","3组","8-10次","自重 3秒下+1秒起","60s","FMS深蹲4分+才做","模拟下坡股四头离心收缩。下放越慢越好。腰不痛才做。" ),
    @("2","跑步机爬坡走(上坡推力)","1组","12-15min","坡度10-12% 速度4.5-5.5","—","腰不痛才做","身体微前倾 不扶扶手。臀+股四头发力爬坡。" ),
    @("3","倒退走(下坡模拟)","1组","5min","坡度5% 速度3-4km/h","—","腰不痛才做","小步慢速 激活臀+腘绳肌——下坡时的刹车。" ),
    @("4","单腿臀桥","3组","12次/侧","自重或持轻哑铃","45s","  零腰压","越野跑每步单腿着地。腰不受力。" ),
    @("5","农夫行走(双侧)","3组","30-40m","哑铃16-24kg/侧","60s","腰不痛才做","躯干正直不侧倾 核心全程收紧。" )
)
$i=1
foreach($a in $actsC1){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=56; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# 第2-4周渐进表
SetC $ws2 $r 1 "第2-4周渐进——动作框架不变 按执行要点推进" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "第1周(适应)" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "核心等级" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "第2周(重建)" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "第3周(进阶)" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("死虫式+鸟狗式","2组 自重","1级","3组 自重→2级","3组 可加弹力带→3级","2组 自重→回1级"),
    @("臀桥","3组 自重","—","3组 哑铃6-10kg","3组 哑铃10-14kg","2组 自重+弹力带"),
    @("深蹲(高脚杯)","自重测试","—","FMS4+ =哑铃6-10kg","FMS4+ =哑铃10-14kg","2组 自重"),
    @("哑铃RDL","测试4-6kg","—","FMS10分+ =8-12kg","FMS10分+ =12-16kg","跳过"),
    @("下拉/划船","轻-中3组","—","+1片4组","再+1片4组","2组弹力带"),
    @("离心深蹲(训练C)","自重3组","—","+哑铃4-6kg","+哑铃6-10kg","2组 更慢"),
    @("倒退走 侧向下台阶","倒退走入","—","+侧向下台阶3x10","+单侧农夫走","跳过")
)
foreach($pw in $prog){
    SetC $ws2 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws2 $r 2 $pw[1] $false 10 0 0 $false
    SetC $ws2 $r 3 $pw[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $pw[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $pw[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $pw[5] $false 10 0 0 $false
    $ws2.Range("A$($r):H$($r)").RowHeight=28; SB $ws2 $r 1 8; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=14; $ws2.Range("H:H").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: FMS五档评分 + 核心阶梯 + 碳水食物换算 V4
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="FMS与饮食"

SetC $ws3 1 1 "FMS五档评分  核心阶梯每周目标  碳水具体食物" $true 16 $headerBlue 0 $true
MC $ws3 1 1 6; $ws3.Range("A1:F1").RowHeight=34
SetC $ws3 2 1 "V4核心工具。FMS对号入座打分 总分决策。核心按周训练 不用想。碳水照着吃 不用算。训练前后必看。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 6

# FMS 五档 V4
$r=4
SetC $ws3 $r 1 "FMS动作筛查——五档评分 对号入座" $true 14 $white $accentGreen $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$accentGreen
$r++

# Survey 1: 深蹲
SetC $ws3 $r 1 "筛查1：深蹲筛查（训练A做）——双手举过头顶 下蹲至大腿平行地面" $true 11 $headerBlue $lightBlue $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

SetC $ws3 $r 1 "分数" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "你的表现" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "训练决策" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "你几分" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "记录" $true 10 $white $headerBlue $false; $r++

$squatScores=@(
    @("5分","无痛完成 动作流畅 深度达标(大腿平行地面)","深蹲可以加重","  分","第1周基准=___分"),
    @("4分","无痛完成 但有点紧(脚跟抬起/上身前倾)","深蹲可以做 但先不加重","  分","第2周=___分"),
    @("3分","有轻微酸胀感 但能完成深度","深蹲只做自重 不加重","  分","第3周=___分"),
    @("2分","有明显疼痛 但咬牙能蹲到平行","跳过深蹲 改高脚杯深蹲(轻哑铃)","  分","第4周=___分"),
    @("1分","无法蹲到平行 或动作中出现锐痛/刺痛","今天不做任何下肢训练","  分","总分15分里这部分贡献")
)
foreach($sc in $squatScores){
    SetC $ws3 $r 1 $sc[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $sc[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $sc[2] $false 10 0 0 $false
    SetC $ws3 $r 5 $sc[3] $false 10 0 0 $false
    SetC $ws3 $r 6 $sc[4] $false 9 0 0 $false
    $ws3.Range("A$($r):F$($r)").RowHeight=30; SB $ws3 $r 1 6; $r++
}

$r++
SetC $ws3 $r 1 "筛查2：单腿站立（训练A做）——双手叉腰 闭眼 单腿站立10秒" $true 11 $headerBlue $lightBlue $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

SetC $ws3 $r 1 "分数" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "你的表现" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "训练决策" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "你几分" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "记录" $true 10 $white $headerBlue $false; $r++

$standScores=@(
    @("5分","稳定站立10秒 几乎不晃","可以加入单腿训练(分腿蹲等)","  分","第1周基准=___分"),
    @("4分","稳定站立10秒 有轻微摇晃但没落地","可以做单腿训练 但先保守","  分","第2周=___分"),
    @("3分","站立5-9秒后落地","只做双腿训练 不加入单腿","  分","第3周=___分"),
    @("2分","站立2-4秒后落地","优先核心稳定训练","  分","第4周=___分"),
    @("1分","几乎站不住 1秒就落地","核心训练是你的第一优先级","  分","总分15分里这部分贡献")
)
foreach($st in $standScores){
    SetC $ws3 $r 1 $st[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $st[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $st[2] $false 10 0 0 $false
    SetC $ws3 $r 5 $st[3] $false 10 0 0 $false
    SetC $ws3 $r 6 $st[4] $false 9 0 0 $false
    $ws3.Range("A$($r):F$($r)").RowHeight=30; SB $ws3 $r 1 6; $r++
}

$r++
SetC $ws3 $r 1 "筛查3：俯卧撑筛查（训练B做）——标准俯卧撑 做到力竭" $true 11 $headerBlue $lightBlue $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

SetC $ws3 $r 1 "分数" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "你的表现" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "训练决策" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "你几分" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "记录" $true 10 $white $headerBlue $false; $r++

$puScores=@(
    @("5分","无痛完成15次以上","上肢训练正常 可以渐进加重","  分","第1周基准=___次"),
    @("4分","无痛完成10-14次","上肢训练正常","  分","第2周=___次"),
    @("3分","无痛完成5-9次","上肢正常 但重量保守","  分","第3周=___次"),
    @("2分","有腰痛但能完成5次以上","改跪姿俯卧撑或上斜俯卧撑","  分","第4周=___次"),
    @("1分","腰痛无法完成 或做1-2次就疼","只用固定器械(推胸机等) 不做自由俯卧撑","  分","总分15分里这部分贡献")
)
foreach($pu in $puScores){
    SetC $ws3 $r 1 $pu[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pu[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pu[2] $false 10 0 0 $false
    SetC $ws3 $r 5 $pu[3] $false 10 0 0 $false
    SetC $ws3 $r 6 $pu[4] $false 9 0 0 $false
    $ws3.Range("A$($r):F$($r)").RowHeight=30; SB $ws3 $r 1 6; $r++
}

SetC $ws3 $r 1 "  把三个分数加起来(满分15分)。13-15=正常 10-12=降级 7-9=只做上肢 4-6=只做核心 <3=休息。" $false 11 $accentGreen 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=26; $r++

# Core Ladder by week V4
$r+=2
SetC $ws3 $r 1 "核心训练——每周明确目标 不用想'今天练哪一级'" $true 14 $white $softPurple $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "级别" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "训练内容（每天热身时做）" $true 10 $white $headerBlue $false; MC $ws3 $r 3 4
SetC $ws3 $r 5 "目标" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "进阶条件" $true 10 $white $headerBlue $false; $r++

$coreByWeek=@(
    @("第1周","1级 基础","死虫式3x10/侧  鸟狗式3x10/侧  臀桥3x15","激活核心 建立神经连接 零腰压","所有动作无痛完成→下周进阶2级"),
    @("第2周","2级 进阶","平板支撑3x30s  侧平板3x20s/侧  单腿臀桥3x12/侧","提高核心耐力——下坡控制基础","平板和侧平板无腰痛→下周进阶3级"),
    @("第3周","3级 高阶","农夫行走3x30s  死虫式+弹力带3x10/侧  鸟狗式+弹力带3x10/侧","核心抗旋转+抗侧屈——越野长途稳定","农夫行走无腰痛→下周进阶4级"),
    @("第4周","回1级 对比","死虫式+鸟狗式+臀桥——和第1周同样动作","对比第1周：动作更稳了吗？腰还疼吗？","和第1周FMS分数对比——看进步")
)
foreach($cw in $coreByWeek){
    SetC $ws3 $r 1 $cw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $cw[1] $false 10 0 0 $false
    SetC $ws3 $r 3 $cw[2] $false 10 0 0 $true; MC $ws3 $r 3 4
    SetC $ws3 $r 5 $cw[3] $false 10 0 0 $false
    SetC $ws3 $r 6 $cw[4] $false 10 0 0 $true
    $ws3.Range("A$($r):F$($r)").RowHeight=46; SB $ws3 $r 1 6; $r++
}

SetC $ws3 $r 1 "  任何时候动作引起腰痛→退回上一级。训练核心不是比谁级别高——是比谁腰更安全。" $false 10 $softPurple 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

# 碳水具体食物 V4
$r+=2
SetC $ws3 $r 1 "碳水循环——直接照着吃 不用算克数" $true 14 $white $softBlue $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "餐次" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "训练日（力量+有氧）" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "非训练日（休息/只有氧）" $true 10 $white $headerBlue $false; MC $ws3 $r 4 6; $r++

$mealPlan=@(
    @("早餐","全麦吐司2片(30g)+香蕉1根(25g)+鸡蛋2个+牛奶","全麦吐司1片(15g)+鸡蛋2个+牛奶(没有香蕉)"),
    @("午餐","米饭1碗(45g)+手掌大肉/鱼+蔬菜不限量","米饭1碗(45g)+手掌大肉/鱼+蔬菜不限量"),
    @("练前2h","香蕉1根(25g)——不吃这个 训练质量打折","不需额外加餐"),
    @("练后1h","米饭1碗(45g)+红薯1个(30g)+蛋白粉","不需额外加餐"),
    @("晚餐","粗粮1拳(30g)糙米/红薯/玉米+手掌大肉/豆腐+蔬菜不限","粗粮1拳(30g)+手掌大肉/豆腐+蔬菜不限(和训练日一样)"),
    @("全天碳水","约210g  达标","约170g  达标")
)
foreach($mp in $mealPlan){
    SetC $ws3 $r 1 $mp[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $mp[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $mp[2] $false 10 0 0 $true; MC $ws3 $r 4 6
    $ws3.Range("A$($r):F$($r)").RowHeight=38; SB $ws3 $r 1 6; $r++
}

$r++
SetC $ws3 $r 1 "全天蛋白质120-140g(每餐有蛋白质+练后蛋白粉)。脂肪40-50g(控油 不吃油炸)。水3L。腰痛期多喝水——助组织修复。" $false 10 0 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

SetC $ws3 $r 1 "越野跑补给：<1h正常吃。1-2h=每小时运动饮料/能量胶。>2h=每小时30-60g碳水+电解质。赛后30min:碳水+蛋白。你是Pro——保持你的习惯。" $false 10 0 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24

$ws3.Range("A:A").ColumnWidth=16; $ws3.Range("B:B").ColumnWidth=24
$ws3.Range("C:D").ColumnWidth=22; $ws3.Range("E:F").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复与备赛
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复与备赛"

SetC $ws4 1 1 "腰痛康复  热身  主动恢复  减载  赛前减量" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

$r=3
SetC $ws4 $r 1 "腰痛康复动作——训练前必做 每天在家可做" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "频率" $true 10 $white $headerBlue $false; MC $ws4 $r 6 8; $r++

$rehab=@(
    @("死虫式","仰卧 对侧手脚同步下放收回 下背贴死地面","训练前+每天在家2-3组"),
    @("鸟狗式","四足跪姿 对侧手脚延伸 骨盆不旋转","同上"),
    @("猫式伸展","四足跪姿 吸抬头塌腰 呼弓背收腹","腰不痛时每天早晚各一组"),
    @("臀桥","脚跟着地 顶峰夹臀2秒 腰完全不受力","训练前+每天在家"),
    @("泡沫轴-臀+腘绳","坐姿滚大腿后侧 侧卧滚臀","训练后+每天睡前"),
    @("游泳(康复性)","低强度自由泳/仰泳 水浮力=零腰压","腰痛期有氧首选")
)
foreach($rh in $rehab){
    SetC $ws4 $r 1 $rh[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $rh[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $rh[2] $false 10 0 0 $false; MC $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight=28; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "训练前热身（含FMS+核心激活 必做）" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws4 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "时间" $true 9 $white $headerBlue $false
SetC $ws4 $r 6 "目的" $true 9 $white $headerBlue $false; MC $ws4 $r 6 8; $r++

$wuList=@(
    @("1","FMS筛查(训练A/B)","2min","量化腰痛 填分数"),
    @("2","死虫式+鸟狗式","各8次/侧","核心激活"),
    @("3","猫式伸展","10次慢节奏","脊柱活动"),
    @("4","自重臀桥","15次","零腰压激活臀"),
    @("5","摆腿(前后+左右)","各10次/侧","髋活动度"),
    @("6","第一个动作轻量试做","1组x10次","知觉今天状态")
)
foreach($wu in $wuList){
    SetC $ws4 $r 1 $wu[0] $false 9 0 0 $false
    SetC $ws4 $r 2 $wu[1] $true 9 $headerBlue 0 $false; MC $ws4 $r 2 4
    SetC $ws4 $r 5 $wu[2] $false 9 0 0 $false
    SetC $ws4 $r 6 $wu[3] $false 9 0 0 $true; MC $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

# Deload
$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 腰休息" $true 14 $white $softPurple $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$deloadList=@(
    @("哑铃臀桥","自重臀桥+弹力带","2组x15次"),
    @("高脚杯深蹲","自重深蹲(扶墙)","2组x12次"),
    @("哑铃RDL","跳过——腰休息",""),
    @("离心深蹲","自重 更慢(5秒)","2组x6-8次"),
    @("下拉/划船","弹力带下拉+划船","2组x10-12次"),
    @("越野跑","低强度游泳/散步","30min <130bpm")
)
foreach($dl in $deloadList){
    SetC $ws4 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $dl[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

# Race taper
$r+=2
SetC $ws4 $r 1 "50km越野跑赛前减量协议" $true 14 $white $accentOrange $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

$taperList=@(
    "赛前2周：停止加重力量。越野跑减至平时60-70%。",
    "赛前1周：力量1次(核心+上肢轻量)。越野跑只短距离。",
    "赛前3天：完全停力量 只拉伸+核心激活。碳水加载(比平时多吃半碗饭)。",
    "比赛日：注意下坡时核心收紧——离心深蹲+侧向下台阶就是为这天。",
    "赛后3-5天：散步+游泳——不碰器械。身体在恢复时变强。"
)
foreach($tp in $taperList){
    SetC $ws4 $r 1 $tp $false 10 0 0 $true; MC $ws4 $r 1 8
    $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++
}

$ws4.Range("A:A").ColumnWidth=14; $ws4.Range("B:B").ColumnWidth=28
$ws4.Range("C:C").ColumnWidth=14; $ws4.Range("D:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 备用与复盘
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="备用与复盘"

SetC $ws5 1 1 "降级方案  4周复盘  备赛提醒" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

$r=3
SetC $ws5 $r 1 "腰痛/经期/加班——降级方案" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws5 $r 1 "本来" $true 10 $white $headerBlue $false; MC $ws5 $r 1 2
SetC $ws5 $r 3 "降级为" $true 10 $white $headerBlue $false
SetC $ws5 $r 5 "底线" $true 10 $white $headerBlue $false; MC $ws5 $r 5 8; $r++

$fbList=@(
    @("训练A(下肢+核心)","只做核心+臀桥+下拉 跳过深蹲RDL","做了核心=完成"),
    @("训练B(上肢+稳定)","只做核心+下拉+划船 跳过俯卧撑","同上"),
    @("训练C(越野+下坡)","只做倒退走+单腿臀桥","同上"),
    @("来例假+腰痛","游泳或完全休息","游泳也算训练")
)
foreach($fb in $fbList){
    SetC $ws5 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $fb[1] $false 10 0 0 $false
    SetC $ws5 $r 5 $fb[2] $false 10 0 0 $true; MC $ws5 $r 5 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

$r+=2
SetC $ws5 $r 1 "4周后复盘——四个核心指标" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "第1周vs第4周" $true 10 $white $headerBlue $false; MC $ws5 $r 2 5
SetC $ws5 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws5 $r 6 8; $r++

$reviewList=@(
    @("FMS总分","第1周总分___ 第4周总分___ 涨了多少？","涨=腰恢复。降=继续康复优先。"),
    @("体重+腰围","减了2-4kg？腰围小了？","4周最大进步——比力量数字重要。"),
    @("核心级别","第1周1级→第4周回到1级 感觉稳了多少？","核心变稳=腰保护层变厚。"),
    @("越野跑5km","比4周前快了吗？下坡更稳了吗？","快了=力量+减重+下坡起作用。")
)
foreach($rv in $reviewList){
    SetC $ws5 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $rv[1] $false 10 0 0 $true; MC $ws5 $r 2 5
    SetC $ws5 $r 6 $rv[2] $false 10 0 0 $true; MC $ws5 $r 6 8
    $ws5.Range("A$($r):H$($r)").RowHeight=28; SB $ws5 $r 1 8; $r++
}

$r+=2
SetC $ws5 $r 1 "50km越野跑备赛提醒" $true 14 $white $softBlue $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

$raceList=@(
    "赛前2周：停止加重。越野跑减60-70%。",
    "赛前1周：力量1次(核心+上肢轻量)。越野跑短距离。",
    "赛前3天：完全停力量。只拉伸+核心。碳水加载。",
    "比赛日：下坡核心收紧——离心深蹲就是为这天。",
    "赛后3-5天：散步+游泳。不碰器械。"
)
foreach($rc in $raceList){
    SetC $ws5 $r 1 $rc $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++
}

$r++
SetC $ws5 $r 1 "你完赛铁三那天哭了——那是Pro的勋章。这次50km 带着FMS 15分的腰 减下来的体重 下坡控制的专项力量去。加油。" $false 10 $softPink 0 $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=28

$ws5.Range("A:A").ColumnWidth=24; $ws5.Range("B:H").ColumnWidth=20
Write-Host "Sheet 5 done"

# Save
$savePath="D:\Codex\members\熊\熊_减重越野跑方案_V4.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
