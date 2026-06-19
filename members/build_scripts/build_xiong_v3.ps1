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

SetC $ws0 1 1 "熊  减重+越野跑 力量方案 V3" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  160cm / 90kg  |  铁三Pro  减重+50km越野备赛  腰痛恢复期  V3=FMS筛查+下坡控制+核心阶梯+碳水循环" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  V3 升级了什么——从'安全'到'安全+高效'" $true 16 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  FMS动作筛查","深蹲筛查+单腿站立+俯卧撑——量化腰痛恢复 每次训练前测 分数写在Sheet3"),
    @("  下坡控制专项","离心深蹲+单腿离心下蹲+倒退走+侧向下台阶——越野跑下坡比上坡更伤"),
    @("  核心稳定阶梯","4级：基础→进阶→高阶→越野跑专项——像升级打怪一样训练核心"),
    @("  碳水循环","训练日碳水200-220g 非训练日160-180g——你是耐力运动员 碳水不是敌人"),
    @("  固定3次/周","腰痛时降级不跳过——3次/周是铁三Pro的最低有效剂量"),
    @("  力量+越野跑同天排法","方案A(力量→越野跑)/B(越野跑→力量)/C(不推荐分开)——按当天状态选")
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  三件事的优先级" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$priority=@(
    @("第一","腰痛恢复——FMS分数是标尺。分数涨=腰在恢复。任何时候不舒服 降级或跳过。"),
    @("第二","减重——每减1kg 越野跑膝盖每步少4kg。碳水循环+有氧=减重主力。"),
    @("第三","50km越野备赛——下坡控制专项+核心抗扭转+上坡推力。铁三有氧已够 力量补短板。")
)
foreach($pr in $priority){
    SetC $ws0 $row 1 $pr[0] $true 10 $softBlue 0 $false
    SetC $ws0 $row 2 $pr[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览——固定3次/周" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "周一A" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "周三B" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "周五C" $true 10 $white $headerBlue $false; $row++

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
SetC $ws0 $row 1 "  力量+越野跑同天怎么排" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$combo=@(
    @("A(推荐)","力量  越野跑","先力量消耗糖原 越野跑时脂肪供能更好——减脂优选。注意疲劳时腰易代偿。"),
    @("B(备选)","越野跑  力量","先跑激活神经 力量时驱动更高——适合第3周。跑后力量注意动作质量。"),
    @("C(不推荐)","分开上下午","没有完整恢复时间 两练质量打折。除非时间确实不允许。")
)
foreach($cb in $combo){
    SetC $ws0 $row 1 $cb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $cb[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+腰痛约束+三件事优先级"," Sheet 方案概览"),
    @("训练A(下肢+核心) B(上肢) C(越野+下坡)"," Sheet 训练计划"),
    @("FMS筛查+核心阶梯+腰痛决策树"," Sheet FMS与核心"),
    @("热身+康复+减载+赛前减量"," Sheet 恢复策略"),
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
$ws1=$wb.Worksheets.Add();$ws1.Name="方案概览"

SetC $ws1 1 1 "熊  减重+越野跑  力量方案概览 V3" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "女 25-30岁 160cm/90kg BMI 35.2 体脂30-35% 铁三完赛(Pro) 越野跑备赛 骑车被撞腰痛恢复中 FMS+下坡控制+核心阶梯+碳水循环" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览 V3" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ov=@(
    @("训练频率","固定3次/周(周一A/周三B/周五C)","腰痛时降级不跳过——3次/周=Pro最低有效剂量"),
    @("核心方法","FMS动作筛查+下坡控制专项+核心进阶阶梯+碳水循环+力量越野跑衔接","6大升级从'安全'到'安全+高效'"),
    @("周期结构","第1周=FMS筛查+适应 第2周=重建 第3周=进阶 第4周=减载+FMS对比","腰痛优先——FMS分数决定能不能进阶"),
    @("核心约束","腰痛恢复期——所有动作以无痛为第一原则","下背保护+核心激活+离心控制=每节课必修"),
    @("方案特点","减重+越野跑+腰痛保护 三合一。铁三Pro的体能方案")
)
$r=5
foreach($o in $ov){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "三件事的优先级——为什么这个顺序" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$why=@(
    "你的情况特殊：能完赛铁三=有氧基础极好。但90kg的体重+腰痛=不能用常规运动员方案。",
    "",
    "第一优先：腰痛恢复。骑车被撞后的腰痛不是肌肉酸痛——是创伤恢复。V3用FMS动作筛查：三个标准动作 每次训练前测 分数记录。分数涨=腰在恢复 分数降=退回上周。数据比感觉诚实。",
    "",
    "第二优先：减重。90kg/160cm——每减1kg=越野跑膝盖每步少4kg。减10kg=每步少40kg。碳水循环：训练日碳水200-220g 非训练日160-180g。你是有氧Pro 碳水不是敌人——但要聪明地吃。",
    "",
    "第三优先：越野跑专项。下坡控制是V3的重头戏——越野跑下坡比上坡更伤 尤其腰痛期。离心深蹲+倒退走+侧向下台阶——这些是爬坡训练给不了你的。"
)
foreach($w in $why){
    SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=20
    if($w -eq ""){$ws1.Range("A$($r):I$($r)").RowHeight=8}
    $r++
}

$r++
SetC $ws1 $r 1 "腰痛训练原则" $true 14 $white $accentRed $true
MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=26
$ws1.Range("A$($r):I$($r)").Font.Color=$white; $ws1.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

$back=@(
    " 禁止：传统硬拉 大重量深蹲 负重俯身划船 跳箱 负重跑步",
    " 谨慎：深蹲(FMS通过才加重) RDL(轻哑铃 全程直背 腰不适停)",
    " 安全：臀推/臀桥(零腰压) 死虫式/鸟狗式 高位下拉 坐姿划船",
    " 康复：死虫式+鸟狗式(热身必做) 猫式伸展(腰不痛时) 游泳(腰痛期有氧首选)",
    "  停止：腰锐痛→当天不练。FMS分数下降→退回上周。"
)
foreach($ba in $back){
    SetC $ws1 $r 1 $ba $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$pr=@(
    @("1","腰痛优先于一切","FMS分数是每周标尺。分数涨=进阶 分数降=退回。用数据跟踪 不凭感觉。"),
    @("2","减重=最好的越野跑训练","碳水循环+保持有氧。每减1kg=膝盖每步少4kg。"),
    @("3","下坡控制比上坡推力更重要","越野跑下坡膝盖冲击力=体重4-6倍。腰痛期 下坡控制=你的护身符。")
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
# SHEET 2: 训练计划 V3
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划 V3——第1-4周 力量+越野专项+下坡控制" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "每周固定3次(周一A/周三B/周五C)。FMS筛查在训练A/B热身时做(见Sheet3)。第4周减载换动作。腰痛时降级不跳过。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# 第1周 训练A
$r=4
SetC $ws2 $r 1 "第1周  周一 训练A  下肢+核心（FMS筛查+适应）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "训练前先做FMS筛查(深蹲+单腿站立——见Sheet3) 记录分数。第1周=轻量适应 腰不痛的动作记录下来。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; $r++

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
    @("1","FMS筛查(必做!)","—","见Sheet3","—","—","训练前测","深蹲筛查+单腿站立筛查 记录分数。" ),
    @("2","死虫式+鸟狗式","各2组","各10次/侧","自重","30s","  零腰压","下背贴地 骨盆不旋转。护腰第一防线。" ),
    @("3","自重臀桥","3组","15次","自重","45s","  零腰压","脚跟着地 顶峰夹臀2秒。腰完全不受力。" ),
    @("4","自重深蹲(扶墙)","2组","10-12次","自重","60s","FMS通过才做","能蹲多深蹲多深。腰不疼=记录深度。腰疼=跳过。" ),
    @("5","轻哑铃RDL","2组","10次","4-6kg哑铃","60s","FMS通过才做","膝盖微屈 哑铃贴大腿下滑 全程直背。腰有不适=立即停。" ),
    @("6","高位下拉","3组","12-15次","轻-中 RPE6-7","60s","  腰不参与","先沉肩再拉 杆拉到锁骨。" )
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
    $ws2.Range("A$($r):H$($r)").RowHeight=54; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# 第1周 训练B
SetC $ws2 $r 1 "第1周  周三 训练B  上肢+稳定性" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

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
    @("1","FMS俯卧撑筛查(必做)","—","1组力竭","自重","—","训练前测","标准俯卧撑做到力竭 记录次数。" ),
    @("2","死虫式+鸟狗式","各3组","各10次/侧","自重","30s","  零腰压","每节课必修——核心稳定=护腰。" ),
    @("3","坐姿划船","3组","12-15次","RPE6-7","60s","  腰不参与","身体不后仰 肩胛夹紧。越野跑控臂。" ),
    @("4","高位下拉(宽握)","3组","12-15次","RPE6-7","60s","  腰不参与","宽握=上背阔肌。长途摆臂稳定。" ),
    @("5","哑铃推肩(坐姿)","3组","10-12次","4-6kg哑铃","60s","  腰有靠背","下背贴靠背 不反弓。" ),
    @("6","面拉+平板支撑","面拉3x15 平板3x20-45s","—","轻重量/自重","45s","平板腰痛改跪姿","面拉=肩后束。平板=核心耐力。" )
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
    $ws2.Range("A$($r):H$($r)").RowHeight=54; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# 第1周 训练C
SetC $ws2 $r 1 "第1周  周五 训练C  越野跑专项+下坡控制入门" $true 13 $white $softBlue $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

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
    @("1","离心深蹲(下坡控制核心)","3组","8-10次","自重 3秒下+1秒起","60s","FMS深蹲通过才做","模拟下坡时股四头离心收缩。下放3秒越慢越好。腰不痛才做。" ),
    @("2","跑步机爬坡走(上坡)","1组","12-15min","坡度10-12% 速度4.5-5.5","—","腰不痛才做","身体微前倾 不扶扶手。臀+股四头发力。" ),
    @("3","倒退走(下坡模拟)","1组","5min","坡度5% 速度3-4km/h","—","腰不痛才做","小步慢速 激活臀+腘绳肌。手扶扶手保平衡。" ),
    @("4","单腿臀桥","3组","12次/侧","自重或持轻哑铃","45s","  零腰压","越野跑每步都是单腿着地。腰不受力。" ),
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
    $ws2.Range("A$($r):H$($r)").RowHeight=54; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# 第2-4周渐进
SetC $ws2 $r 1 "第2-4周渐进——动作框架不变 重量和难度递增" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "第1周(适应)" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "第2周(重建)" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "第3周(进阶)" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "第4周(减载)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("死虫式+鸟狗式","2组 自重","3组 自重","3组 可加轻弹力带","2组 自重"),
    @("臀桥","3组 自重","3组 哑铃6-10kg","3组 哑铃10-14kg","2组 自重+弹力带"),
    @("深蹲(高脚杯)","FMS通过=自重测试","FMS通过=哑铃6-10kg","FMS通过=哑铃10-14kg","2组 自重"),
    @("哑铃RDL","测试 4-6kg","FMS通过=8-12kg","FMS通过=12-16kg","跳过——腰休息"),
    @("高位下拉+划船","轻-中 3组","+1片 4组","再+1片 4组","2组 弹力带"),
    @("离心深蹲(训练C)","自重 3组","+哑铃4-6kg","+哑铃6-10kg","2组 自重 更慢"),
    @("倒退走 侧向下台阶","倒退走入","+侧向下台阶3x10/侧","+单侧农夫行走","跳过——腰休息")
)
foreach($pw in $prog){
    SetC $ws2 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false; MC $ws2 $r 1 2
    SetC $ws2 $r 3 $pw[1] $false 9 0 0 $false
    SetC $ws2 $r 4 $pw[2] $false 9 0 0 $false
    SetC $ws2 $r 5 $pw[3] $false 9 0 0 $false
    SetC $ws2 $r 6 $pw[4] $false 9 0 0 $false
    $ws2.Range("A$($r):H$($r)").RowHeight=28; SB $ws2 $r 1 8; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=8; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=14; $ws2.Range("H:H").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: FMS与核心 V3 (完全重写)
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="FMS与核心"

SetC $ws3 1 1 "FMS动作筛查  核心进阶阶梯  腰痛决策树" $true 16 $headerBlue 0 $true
MC $ws3 1 1 6; $ws3.Range("A1:F1").RowHeight=34
SetC $ws3 2 1 "V3核心模块。每次训练前必看——FMS量化腰痛恢复 核心阶梯让训练有进阶路径 腰痛决策树决定今天做什么。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 6

# FMS Section
$r=4
SetC $ws3 $r 1 "FMS动作筛查——训练A和B的热身前做 记录分数" $true 14 $white $accentGreen $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "筛查项目" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "怎么做" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "评分标准" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "你的分数" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "训练决策" $true 10 $white $headerBlue $false; $r++

$fmsItems=@(
    @("1.深蹲筛查(训练A做)","双手举过头顶 下蹲至大腿平行地面","  无痛完成 /   有痛但完成 /   无法完成或锐痛","___周=___分","  可加重 /   只自重 /   跳过深蹲"),
    @("2.单腿站立(训练A做)","双手叉腰 闭眼单腿站立10秒","  稳定完成 /   摇晃但完成 /   无法完成","___周=___分","  可单腿 /   只双腿 /   核心优先"),
    @("3.俯卧撑筛查(训练B做)","标准俯卧撑 下巴贴地 做到力竭","  无痛完成 /   腰痛但完成 /   无法完成或锐痛","___周=力竭___次","  正常 /   跪姿 /   固定器械")
)
foreach($fm in $fmsItems){
    SetC $ws3 $r 1 $fm[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $fm[1] $false 10 0 0 $true
    SetC $ws3 $r 4 $fm[2] $false 10 0 0 $false
    SetC $ws3 $r 5 $fm[3] $false 10 0 0 $false
    SetC $ws3 $r 6 $fm[4] $false 10 0 0 $true
    $ws3.Range("A$($r):F$($r)").RowHeight=48; SB $ws3 $r 1 6; $r++
}

SetC $ws3 $r 1 "  第1周建基准 第4周再做一次对比。分数涨=腰在恢复。数据比感觉诚实。" $false 10 $accentGreen 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

# Core Ladder
$r+=2
SetC $ws3 $r 1 "核心稳定进阶阶梯——像升级打怪一样训练核心" $true 14 $white $softPurple $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$softPurple
$r++

SetC $ws3 $r 1 "级别" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "训练内容" $true 10 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "目标" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "进阶条件" $true 10 $white $headerBlue $false; $r++

$coreLadder=@(
    @("1级 基础(第1周)","死虫式3x10/侧  鸟狗式3x10/侧  臀桥3x15","激活核心 神经连接 零腰压","当前级所有动作无痛→进阶"),
    @("2级 进阶(第2周)","+平板支撑3x30s  +侧平板3x20s/侧  +单腿臀桥3x12/侧","提高核心耐力——下坡控制基础","平板和侧平板无腰痛→进阶"),
    @("3级 高阶(第3周)","+农夫行走3x30s  +死虫式弹力带3x10/侧  +鸟狗式弹力带3x10/侧","核心抗旋转+抗侧屈——越野长途稳定","农夫行走无腰痛+弹力带动稳定→进阶"),
    @("4级 专项(第4周后)","+单腿平板3x20s/侧  +登山者3x20/侧  +药球旋转抛3x10/侧","模拟越野跑动态核心需求","第4周减载结束后 下周期可启用")
)
foreach($cl in $coreLadder){
    SetC $ws3 $r 1 $cl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $cl[1] $false 10 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $cl[2] $false 10 0 0 $false
    SetC $ws3 $r 5 $cl[3] $false 10 0 0 $true
    $ws3.Range("A$($r):F$($r)").RowHeight=44; SB $ws3 $r 1 6; $r++
}

SetC $ws3 $r 1 "  任何时候当前级别有动作引起腰痛→退回上一级。核心训练不是比谁级别高——是比谁的腰更安全。" $false 10 $softPurple 0 $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=24; $r++

# Back Pain Decision Tree
$r+=2
SetC $ws3 $r 1 "腰痛决策树——每次训练前必看" $true 14 $white $accentRed $true
MC $ws3 $r 1 6; $ws3.Range("A$($r):F$($r)").RowHeight=28
$ws3.Range("A$($r):F$($r)").Font.Color=$white; $ws3.Range("A$($r):F$($r)").Interior.Color=$accentRed
$r++

SetC $ws3 $r 1 "今天的信号" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "今天做什么" $true 10 $white $headerBlue $false; MC $ws3 $r 4 6; $r++

$tree=@(
    @("FMS三项全 +腰无不适","  绿灯——安全训练","正常按计划走。深蹲和RDL可以加重量。核心尝试下一级。" ),
    @("FMS有 腰轻微酸胀","  黄灯——保守训练","做核心+上肢+臀桥。深蹲和RDL跳过或用自重。下坡只做倒退走。" ),
    @("FMS有 动作中锐痛","  红灯——立即停","只做零腰压动作(死虫式+臀桥+下拉)。持续疼→训练结束。" ),
    @("早上起床腰就疼/FMS分数下降","  红灯——退回上周","只做上肢(下拉+划船+推肩+面拉)+核心基础级。不碰下肢和越野。" ),
    @("FMS分数连续2周在涨","  腰在恢复！","可以尝试下周进阶。但第3周仍保守——腰比训练计划重要。" ),
    @("第3周末FMS仍有 ","  恢复不够 不急","延长第2周训练到第3-4周。不要急着加重——等腰痛完全消失。" )
)
foreach($tr in $tree){
    SetC $ws3 $r 1 $tr[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $tr[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $tr[2] $false 10 0 0 $true; MC $ws3 $r 4 6
    $ws3.Range("A$($r):F$($r)").RowHeight=28; SB $ws3 $r 1 6; $r++
}

$ws3.Range("A:A").ColumnWidth=24; $ws3.Range("B:B").ColumnWidth=20
$ws3.Range("C:C").ColumnWidth=16; $ws3.Range("D:F").ColumnWidth=22
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "腰痛康复  热身  主动恢复  减载周" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

$r=3
SetC $ws4 $r 1 "腰痛康复动作——每次训练前必做 每天在家也可以做" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "频率" $true 10 $white $headerBlue $false; MC $ws4 $r 6 8; $r++

$rehab=@(
    @("死虫式","仰卧 手举腿抬90度 对侧手脚同步下放收回 下背贴死地面","训练前+每天在家2-3组"),
    @("鸟狗式","四足跪姿 对侧手脚延伸 身体一条线 骨盆不旋转","同上"),
    @("猫式伸展","四足跪姿 吸抬头塌腰 呼弓背收腹 逐节活动","腰不痛时每天早晚各一组"),
    @("臀桥","脚跟着地 顶峰夹臀2秒 腰完全不受力","训练前+每天在家"),
    @("泡沫轴-臀+腘绳","坐姿滚大腿后侧 侧卧滚臀 酸痛点停15秒","训练后+每天睡前"),
    @("游泳(康复性)","低强度自由泳/仰泳 水浮力=零腰压有氧","腰痛期有氧首选")
)
foreach($rh in $rehab){
    SetC $ws4 $r 1 $rh[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $rh[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $rh[2] $false 10 0 0 $false; MC $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight=28; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "训练前热身（含FMS+核心激活）" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws4 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws4 $r 2 4
SetC $ws4 $r 5 "时间" $true 9 $white $headerBlue $false
SetC $ws4 $r 6 "目的" $true 9 $white $headerBlue $false; MC $ws4 $r 6 8; $r++

$wuList=@(
    @("1","FMS筛查(训练A/B)","1-2min","量化腰痛 记录分数"),
    @("2","死虫式+鸟狗式","各8次/侧","核心激活——护腰防线"),
    @("3","猫式伸展","10次慢节奏","脊柱活动 腰不痛做"),
    @("4","自重臀桥","15次","零腰压激活臀大肌"),
    @("5","摆腿(前后+左右)","各10次/侧","髋关节活动度"),
    @("6","第一个动作轻重量试做","1组x10次","知觉今天状态")
)
foreach($wu in $wuList){
    SetC $ws4 $r 1 $wu[0] $false 9 0 0 $false
    SetC $ws4 $r 2 $wu[1] $true 9 $headerBlue 0 $false; MC $ws4 $r 2 4
    SetC $ws4 $r 5 $wu[2] $false 9 0 0 $false
    SetC $ws4 $r 6 $wu[3] $false 9 0 0 $true; MC $ws4 $r 6 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 腰充分休息" $true 14 $white $softPurple $true
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
    @("离心深蹲","自重 更慢(5秒下放)","2组x6-8次"),
    @("高位下拉+划船","弹力带下拉+划船","2组x10-12次"),
    @("越野跑","低强度游泳/散步/骑","20-30min 心率<130")
)
foreach($dl in $deloadList){
    SetC $ws4 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $dl[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=14; $ws4.Range("B:B").ColumnWidth=28
$ws4.Range("C:C").ColumnWidth=14; $ws4.Range("D:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食与规则
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与规则"

SetC $ws5 1 1 "碳水循环  越野补给  经期应对  判断规则" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

$r=3
SetC $ws5 $r 1 "碳水循环——训练日和非训练日碳水不同 V3新增" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你是耐力运动员——碳水不是敌人。碳水循环让你训练日有能量 休息日加速减脂。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws5 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "训练日(力量+有氧)" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "非训练日" $true 10 $white $headerBlue $false
SetC $ws5 $r 6 "怎么吃" $true 10 $white $headerBlue $false; MC $ws5 $r 6 8; $r++

$carbData=@(
    @("总热量","1800-2000 kcal","1600-1700 kcal","训练日多200-300——支撑训练质量"),
    @("碳水","200-220g","160-180g","三餐一拳 训练前后集中吃碳水"),
    @("蛋白质","120-140g","120-140g(不变)","每餐有蛋白质 练后蛋白粉"),
    @("脂肪","40-50g","40-50g(不变)","控油 不吃油炸"),
    @("训练窗口","练前2h:碳水30-40g 练后1h:碳水40-50g","不需额外加餐","训练前后窗口=碳水循环核心"),
    @("饮水","3L/天","3L/天","腰痛期多喝水——助组织修复")
)
foreach($cd in $carbData){
    SetC $ws5 $r 1 $cd[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $cd[1] $false 10 0 0 $true; MC $ws5 $r 2 3
    SetC $ws5 $r 4 $cd[2] $false 10 0 0 $true; MC $ws5 $r 4 5
    SetC $ws5 $r 6 $cd[3] $false 10 0 0 $true; MC $ws5 $r 6 8
    $ws5.Range("A$($r):H$($r)").RowHeight=30; SB $ws5 $r 1 8; $r++
}

$r++
SetC $ws5 $r 1 "越野跑补给：训练<1h正常吃  1-2h=每小时运动饮料/能量胶  >2h=每小时30-60g碳水+电解质。赛后30min:碳水+蛋白。你是Pro 应该有自己的习惯——保持。" $false 10 0 0 $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=24; $r++

$r+=2
SetC $ws5 $r 1 "经期前后" $true 14 $white $softPink $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

SetC $ws5 $r 1 "经前一周：体重涨1-3kg=正常水肿。腰更要注意——经前韧带松弛 更容易受伤。训降强度。经期1-2天：休息或游泳(低强度)。第3天恢复。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=24; $r++

$r+=2
SetC $ws5 $r 1 "身体信号与判断规则" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "信号" $true 10 $white $headerBlue $false; MC $ws5 $r 1 2
SetC $ws5 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws5 $r 3 8; $r++

$rulesList=@(
    @("FMS分数下降","退回上周——腰在说今天不行"),
    @("腰锐痛 任何动作触发","立即停。当天不练。连续2天→看医生"),
    @("FMS分数连续2周在涨","  可以进阶。下周期尝试核心下一级"),
    @("训练+有氧后极度疲劳","减重+伤病+训练 三线作战。这周多睡1h"),
    @("体重一周没掉","先看是否经前。不是→查油脂和隐性碳水"),
    @("腰痛消失+FMS三项全 ","  最好信号。下周期正式进阶")
)
foreach($rl in $rulesList){
    SetC $ws5 $r 1 $rl[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $rl[1] $false 10 0 0 $true; MC $ws5 $r 3 8
    $ws5.Range("A$($r):H$($r)").RowHeight=28; SB $ws5 $r 1 8; $r++
}

$ws5.Range("A:A").ColumnWidth=24; $ws5.Range("B:H").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用与复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "降级方案  4周复盘  备赛提醒" $true 16 $headerBlue 0 $true
MC $ws6 1 1 8; $ws6.Range("A1:H1").RowHeight=34

$r=3
SetC $ws6 $r 1 "腰痛/经期/加班——降级方案" $true 14 $white $accentOrange $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "本来" $true 10 $white $headerBlue $false; MC $ws6 $r 1 2
SetC $ws6 $r 3 "降级为" $true 10 $white $headerBlue $false
SetC $ws6 $r 5 "底线" $true 10 $white $headerBlue $false; MC $ws6 $r 5 8; $r++

$fbList=@(
    @("训练A(下肢+核心)","只做核心激活+臀桥+下拉——跳过深蹲RDL","做了核心=今天算完成"),
    @("训练B(上肢+稳定)","只做核心+下拉+划船——跳过俯卧撑筛查","同上"),
    @("训练C(越野+下坡)","只做倒退走+单腿臀桥——跳过离心深蹲爬坡","同上"),
    @("来例假+腰痛","游泳或完全休息","游泳也算训练")
)
foreach($fb in $fbList){
    SetC $ws6 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false; MC $ws6 $r 1 2
    SetC $ws6 $r 3 $fb[1] $false 10 0 0 $false
    SetC $ws6 $r 5 $fb[2] $false 10 0 0 $true; MC $ws6 $r 5 8
    $ws6.Range("A$($r):H$($r)").RowHeight=26; SB $ws6 $r 1 8; $r++
}

$r+=2
SetC $ws6 $r 1 "4周后复盘——四个核心指标" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "第1周vs第4周" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期决策" $true 10 $white $headerBlue $false; MC $ws6 $r 6 8; $r++

$reviewList=@(
    @("FMS分数","深蹲/单腿/俯卧撑各涨多少？","三项全 =正式进阶。仍有 →继续康复"),
    @("体重+腰围","减了2-4kg？腰围小了？","4周最大进步——比力量数字重要"),
    @("核心阶梯","从第几级升到第几级？","升级=腰保护层在变厚"),
    @("越野跑5km","快了吗？下坡更稳了吗？","快了=力量+减重+下坡起作用")
)
foreach($rv in $reviewList){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 5
    SetC $ws6 $r 6 $rv[2] $false 10 0 0 $true; MC $ws6 $r 6 8
    $ws6.Range("A$($r):H$($r)").RowHeight=28; SB $ws6 $r 1 8; $r++
}

$r+=2
SetC $ws6 $r 1 "50km越野跑备赛提醒" $true 14 $white $softBlue $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

$raceList=@(
    "赛前2周：停止加重力量 越野跑减至平时60-70%",
    "赛前1周：力量1次(核心+上肢轻量) 越野跑只短距离",
    "赛前3天：完全停力量 只拉伸+核心 碳水加载",
    "比赛日：注意下坡时核心收紧——离心深蹲+侧向下台阶就是为这天练的",
    "赛后3-5天：散步+游泳——不碰器械。身体在恢复时变强"
)
foreach($rc in $raceList){
    SetC $ws6 $r 1 $rc $false 10 0 0 $true; MC $ws6 $r 1 8
    $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++
}

$r++
SetC $ws6 $r 1 "你完赛铁三那天哭了——那不是脆弱的眼泪 是Pro的勋章。这次50km越野跑 带着FMS三项全  的腰 带着碳水循环减下的体重 带着下坡控制的专项力量去。加油。" $false 10 $softPink 0 $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=28

$ws6.Range("A:A").ColumnWidth=24; $ws6.Range("B:H").ColumnWidth=20
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\熊\熊_减重越野跑方案_V3.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
