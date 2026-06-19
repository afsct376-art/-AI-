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

SetC $ws0 1 1 "小冯  减脂训练方案 V2" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-19  |  107kg  减脂  |  把之前会的动作捡起来" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你的第一个月——数据不会骗人" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "5月7日→6月7日：体重107kg→105.2kg（-1.8kg）腰围116.5cm→113cm（-3.5cm）胸围123→121 手臂40.5→39.8 大腿70.5→69.7。这一个月我们大概上课练了10次——你有工作 有应酬 有痛风 睡眠不好。但数据就在这——你做到了。" $false 10 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=36; $row++

$row+=2
SetC $ws0 $row 1 "  现在这个阶段——三个方案 什么时候都有一条路" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$plans=@()
$plans+=,@("方案A 在家15分钟","完全不想出门的日子——自重深蹲+跪姿俯卧撑+仰卧举腿+原地踏步。15分钟搞定。")
$plans+=,@("方案B 铁馆精简版","去铁馆但不想练太久——选2-3个核心动作+有氧。30-40分钟。")
$plans+=,@("方案C 铁馆完整版","精力好的日子——4-5个动作+有氧20min+拉伸。50-60分钟。你之前上课的标准。")
foreach($pl in $plans){
    SetC $ws0 $row 1 $pl[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $pl[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; $row++
}

$row+=2
SetC $ws0 $row 1 "  训练结构——上下肢两分化 全部是你做过的动作" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "动作来源" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "强度" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$split=@()
$split+=,@("训练A","上肢日（胸+背+肩+手臂）","坐姿推胸/上斜推胸/高位下拉/坐姿划船/推肩/下压/弯举——都做过","中等","胸和背都有泵感——推拉平衡")
$split+=,@("训练B","下肢+核心日","倒蹬机/哈克深蹲/腿屈伸/腿弯举/髋外展/臀桥——都做过","中等","大腿和臀有劲——腰不参与")
foreach($sp in $split){
    SetC $ws0 $row 1 $sp[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $sp[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $sp[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $sp[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $sp[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  这周怎么练——根据你的状态自己选" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$freq=@()
$freq+=,@("  精力还行","3-4次/周：周一A 周三B 周五A 周日B（轮换）","方案C——完整版 4-5个动作+20min有氧")
$freq+=,@("  有点累 但想去","2-3次/周：周一A 周四B（或周三A 周六B）","方案B——精简版 选2-3个核心动作+有氧")
$freq+=,@("  累到不行 完全不想出门","——","方案A——在家15分钟 自重训练。做了就完成")
foreach($fr in $freq){
    SetC $ws0 $row 1 $fr[0] $true 10 $headerBlue $lightGray $false; MC $ws0 $row 1 2
    SetC $ws0 $row 3 $fr[1] $false 10 0 0 $false
    SetC $ws0 $row 4 $fr[2] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

SetC $ws0 $row 1 " 完成比完美重要。方案A 15分钟也算训练。不要因为只能在家练就觉得不算——做了就是做了。" $false 10 $accentGreen 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++

$row+=2
SetC $ws0 $row 1 "  四周总览——稳就行 不追求快" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "动作数" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "重量" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周重点" $true 10 $white $headerBlue $false; $row++

$o4=@()
$o4+=,@("1","适应——重建节奏","方案C=4-5个 方案B=2-3个 方案A=自重","中等——选做完15次有点累的","把节奏找回来——不求重量 不求频率")
$o4+=,@("2","稳定——把频率稳下来","同上","比第1周重1片（如果轻松）","上周去几次这周就几次——稳下来就是进步")
$o4+=,@("3","加量——能做更多","方案C做到5-6个动作 方案B做到3-4个","不加重——先加动作数","身体开始适应了——同样的时间做得更多")
$o4+=,@("4","减载——恢复+测量","减半——方案C变B B变A","轻重量","测体重腰围 和第1天对比。这个月变了多少？")
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  腰和脚踝——训练前检查" $true 14 $white $accentRed $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentRed
$row++

$protect=@()
$protect+="  腰保护：不杠铃硬拉 不杠铃划船。下拉和划船用固定器械——腰不参与。硬拉器硬拉=腰安全——可以做。"
$protect+="  脚踝保护：不跳跃 不站姿提踵。倒蹬机和哈克深蹲=脚踝无冲击——可以做。有氧选椭圆机或快走。"
$protect+="  热身必须做：上次你练上肢小臂酸——是没热身上重量了。每次5分钟（踏步+肩绕+摆腿+第一个动作轻量试做）。"
foreach($prt in $protect){SetC $ws0 $row 1 $prt $false 10 0 0 $true; MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@()
$idx+=,@("完整方案逻辑+为什么这样设计"," Sheet 方案概览")
$idx+=,@("训练A上肢日 B下肢日 完整版和精简版"," Sheet 训练计划")
$idx+=,@("方案A在家15分钟 自重训练"," Sheet 在家训练")
$idx+=,@("每周怎么进阶+怎么判断"," Sheet 渐进与规则")
$idx+=,@("热身流程+有氧+腰脚踝保护"," Sheet 恢复策略")
$idx+=,@("饮食怎么吃+睡眠+应酬"," Sheet 饮食与生活")
$idx+=,@("4周后复盘+数据对比+最后的话"," Sheet 备用与复盘")
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

SetC $ws1 1 1 "小冯  减脂训练方案  概览 V2" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "男 仓储物流管理 107kg 减脂 铁馆器械齐全 腰+脚踝保护 ABC三方案 完成比完美重要" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览 V2" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ovl=@()
$ovl+=,@("训练频率","每周2-4练 两分化轮换（上肢日/下肢日交替）","方案C=完整版 B=精简版 A=在家版")
$ovl+=,@("训练结构","上肢日（胸背肩手臂推拉平衡）下肢日（股四腘绳臀核心）","全部是你跟我做过的动作——不需要学新的")
$ovl+=,@("设备来源","铁馆器械（坐姿推胸 倒蹬 哈克深蹲 高位下拉 坐姿划船等）+ 在家自重","杠铃硬拉和划船不用——腰受不了 用固定器械替代")
$ovl+=,@("周期结构","第1周=适应 第2周=稳定 第3周=加量 第4周=减载+测量","不追求大重量——先加动作数 再加重")
$ovl+=,@("核心约束","腰痛=不杠铃硬拉/划船。脚踝=不跳跃/站姿提踵","你之前上课用的动作 腰和脚踝都是安全的")
$ovl+=,@("方案特点","ABC三方案 不怕没时间 不怕不想出门 不怕累——总有一条路")
$r=5
foreach($o in $ovl){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "为什么V2这样设计——给现在的你" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$why=@(
    "你第一个月上课练了大概10次——体重-1.8kg 腰围-3.5cm。这个数据不是'不够'——是'你已经开始了'。",
    "最近工作累 睡眠不好 痛风脚踝不舒服——训练断了一段时间。V2不催你'赶紧回去练'——给你三条路：精力好去铁馆完整练 有点累去铁馆精简练 累到不行在家15分钟。",
    "全部用你跟我做过的动作——坐姿推胸 倒蹬 哈克深蹲 高位下拉 坐姿划船 髋外展……不需要学新的。杠铃硬拉和划船腰疼就不做。脚踝不舒服就椭圆机和坐姿动作。",
    "完成比完美重要。方案A 15分钟在家也算训练。不要因为'只能在家练'就觉得不算——做了就是做了。"
)
foreach($w in $why){
    SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "你的动作库——全部是跟我做过的" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

SetC $ws1 $r 1 "上肢动作（胸+背+肩+手臂）" $true 10 $headerBlue $lightGray $false
SetC $ws1 $r 2 "坐姿水平推胸 坐姿上斜推胸 仰卧器械推胸 上斜俯卧撑 坐姿水平夹胸 坐姿高位下拉 坐姿高位下拉(窄距) 坐姿水平划船 哑铃单手划船 站姿辅助引体 站姿屈臂下压 哑铃臂屈伸 哑铃站姿弯举 悍马机大剪刀 支撑式悍马俯身划船" $false 10 0 0 $true
MC $ws1 $r 2 9; $ws1.Range("A$($r):I$($r)").RowHeight=36; SB $ws1 $r 1 9; $r++

SetC $ws1 $r 1 "下肢动作（股四+腘绳+臀）" $true 10 $headerBlue $lightGray $false
SetC $ws1 $r 2 "倒蹬机 哈克深蹲 坐姿腿屈伸 俯卧腿弯举 坐姿髋外展 坐姿髋内收 硬拉器硬拉 哑铃高脚杯深蹲 自重臀桥 分腿蹲 坐姿正蹬" $false 10 0 0 $true
MC $ws1 $r 2 9; $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++

SetC $ws1 $r 1 "在家自重动作" $true 10 $headerBlue $lightGray $false
SetC $ws1 $r 2 "自重深蹲(扶椅背) 跪姿/上斜俯卧撑 仰卧举腿 原地踏步/开合跳 拉伸" $false 10 0 0 $true
MC $ws1 $r 2 9; $ws1.Range("A$($r):I$($r)").RowHeight=22; SB $ws1 $r 1 9; $r++

$r++
SetC $ws1 $r 1 "腰和脚踝——不是你的问题 是动作的问题" $true 14 $white $accentRed $true
MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=26
$ws1.Range("A$($r):I$($r)").Font.Color=$white; $ws1.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

$back=@()
$back+="  上次练上肢小臂酸——不是你不会练。是没热身+重量大了。这次每次训练前先热身5分钟 第一个动作轻重量试做。"
$back+="  杠铃硬拉和划船腰不舒服——不是你的问题。是动作本身对腰的压力大。用坐姿划船和高位下拉代替——腰完全不受力。"
$back+="  脚踝不舒服——不是借口。是痛风引起的真问题。不跳跃 不站姿提踵 有氧选椭圆机。脚踝好之前 腿日只做坐姿动作。"
foreach($ba in $back){SetC $ws1 $r 1 $ba $false 10 0 0 $true; MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$prl=@()
$prl+=,@("1","完成比完美重要","方案A 15分钟在家=训练。方案B 30分钟精简版=训练。方案C 60分钟完整版=训练。做了就是完成了。")
$prl+=,@("2","用你做过的动作","全部是我带你做过的——坐姿推胸 倒蹬 哈克深蹲 高位下拉 划船。不需要自己摸索 不需要学新的。")
$prl+=,@("3","腰和脚踝优先","杠铃硬拉划船腰痛就不做——用固定器械。脚踝不舒服就不跳跃——有方案A在家。身体信号优先于训练计划。")
foreach($p in $prl){
    SetC $ws1 $r 1 $p[0] $false 10 0 0 $false
    SetC $ws1 $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws1 $r 3 $p[2] $false 10 0 0 $true; MC $ws1 $r 3 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$ws1.Range("A:A").ColumnWidth=30; $ws1.Range("B:I").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划（两分化 方案B和C）
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划——上肢日/下肢日 方案B(精简)+方案C(完整)" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "所有动作都是你跟我做过的。方案C=完整版4-5个动作(精力好)。方案B=精简版选2-3个核心动作(有点累)。方案A=在家15分钟(见Sheet3)。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# 训练A: 上肢日
$r=4
SetC $ws2 $r 1 "训练A  上肢日（胸+背+肩+手臂）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=每组12-15次 中等重量 | 第2周=比上周重1片 | 第3周=组数+1或次数+2-3 | 第4周=轻重量 每组10-12次 减半" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "你做过" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "方案C完整版" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "方案B精简版" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰安全" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsA=@()
$actsA+=,@("1","坐姿水平推胸","","3组x12-15次","  精简版必选","60s","  腰有靠背","后背贴紧靠背 推到前面不锁肘。你上课做过——找胸肌发力的感觉。")
$actsA+=,@("2","高位下拉（宽握）","","3组x12-15次","  精简版必选","60s","  腰不参与","先沉肩再下拉——不要用手臂硬拽。杆拉到锁骨。你练过很多次了。")
$actsA+=,@("3","坐姿上斜推胸 或 仰卧器械推胸","","3组x12-15次","完整版选1个","60s","  腰有靠背","上斜=上胸。仰卧器械=和平板推胸不同角度。和动作1互补。")
$actsA+=,@("4","坐姿水平划船","","3组x12-15次","完整版选做","60s","  腰有靠背","身体不后仰 拉到腹部 肩胛骨往中间夹。和下拉互补=垂直拉+水平拉。")
$actsA+=,@("5","站姿屈臂下压 或 哑铃弯举","","各3组x12-15次","完整版选1个","45s","  站姿但腰安全","下压=手臂后侧 弯举=手臂前侧。你上课都做过——轻重量先找感觉。上次小臂酸就是没热身+太重了。")
$actsA+=,@("6","有氧","","快走/椭圆机 20min","完整版=20min 精简版=能做就做","—","脚踝=椭圆机","心率130-140。能断续说话不能流畅聊天。时间不够就砍有氧——动作优先。")

$i=1
foreach($a in $actsA){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=52; SB $ws2 $r 1 8; $r++; $i++
}
$r+=2

# 训练B: 下肢日
SetC $ws2 $r 1 "训练B  下肢+核心日（股四头+腘绳肌+臀）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=每组12-15次 中等重量 | 第2周=比上周重1片 | 第3周=组数+1或次数+2-3 | 第4周=轻重量 减半  |  脚踝不舒服=只做坐姿动作" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "你做过" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "方案C完整版" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "方案B精简版" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰/脚踝" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsB=@()
$actsB+=,@("1","倒蹬机 或 哈克深蹲","","3组x12-15次","  精简版必选（二选一）","90s","  腰有靠背 脚踝无冲击","你上课都做过。倒蹬机=脚踩实 膝盖不完全蹬直。哈克深蹲=躯干贴靠背。")
$actsB+=,@("2","坐姿腿屈伸","","3组x12-15次","  精简版必选","60s","  腰有靠背 脚踝无冲击","脚尖回勾 大腿前侧发力。膝盖不舒服就减重量——你练过的。")
$actsB+=,@("3","俯卧腿弯举 或 硬拉器硬拉","","3组x12-15次","完整版选1个","60s","  腰安全 硬拉器比杠铃安全","腿弯举=大腿后侧发力勾下。硬拉器硬拉=比杠铃硬拉腰安全得多——可以做。")
$actsB+=,@("4","坐姿髋外展","","3组x15-20次","完整版选做","45s","  腰有靠背","身体贴紧靠背不晃 臀外侧发力。改善胯和大腿外侧线条。")
$actsB+=,@("5","自重臀桥","","3组x15次","完整版选做","45s","  零腰压 零脚踝","脚跟着地 顶峰夹屁股2秒。腰完全不受力——放心做。")
$actsB+=,@("6","有氧","","快走/椭圆机 20min","完整版=20min 精简版=能做就做","—","脚踝=椭圆机","心率130-140。腿日+有氧=今天消耗最大化。")

$i=1
foreach($a in $actsB){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 10 0 0 $true
    $ws2.Range("A$($r):H$($r)").RowHeight=52; SB $ws2 $r 1 8; $r++; $i++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18
$ws2.Range("C:C").ColumnWidth=8; $ws2.Range("D:D").ColumnWidth=18
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=16; $ws2.Range("H:H").ColumnWidth=48
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 在家训练（方案A）
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="在家训练"

SetC $ws3 1 1 "方案A——在家15分钟 完全不想出门的日子" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34
SetC $ws3 2 1 "这是'退路方案'——不是'偷懒方案'。15分钟做完就完成。你之前在家练过开合跳和深蹲——这个你也熟。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 8

$r=4
SetC $ws3 $r 1 "在家15分钟——照着做" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "怎么做" $true 9 $white $headerBlue $false; MC $ws3 $r 4 6
SetC $ws3 $r 7 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 8 "感觉" $true 9 $white $headerBlue $false; $r++

$homeEx=@()
$homeEx+=,@("1","原地踏步/开合跳","原地踏步1min→开合跳30s→重复","3min","微微出汗 心率上来")
$homeEx+=,@("2","自重深蹲(扶椅背)","脚与肩同宽 扶着椅子/墙 能蹲多深蹲多深","3组x15次","大腿有感觉 膝盖不疼")
$homeEx+=,@("3","跪姿/上斜俯卧撑","手扶沙发/床沿做上斜俯卧撑 或跪姿做","3组x力竭","胸肌和手臂有感觉")
$homeEx+=,@("4","仰卧举腿","躺地上 腿伸直上举 下腹发力","3组x15次","下腹有收紧感")
$homeEx+=,@("5","拉伸","胸+背+大腿前后 各30s","3min","放松——告诉身体今天练完了")
foreach($hm in $homeEx){
    SetC $ws3 $r 1 $hm[0] $true 9 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $hm[1] $true 9 $white $darkBg $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $hm[2] $false 9 0 0 $true; MC $ws3 $r 4 6
    SetC $ws3 $r 7 $hm[3] $false 9 0 0 $false
    SetC $ws3 $r 8 $hm[4] $false 9 0 0 $false
    $ws3.Range("A$($r):H$($r)").RowHeight=32; SB $ws3 $r 1 8; $r++
}

$r++
SetC $ws3 $r 1 "  15分钟做完=今天训练完成。不要因为'只能在家练'就觉得不算——做了就是做了。完成比完美重要。" $false 10 $accentGreen 0 $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=24; $r++

$r+=2
SetC $ws3 $r 1 "在家训练——什么时候用 用多久" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$homeUse=@()
$homeUse+="什么时候用：今天完全不想出门/下大雨/加班到很晚/脚踝不舒服/腰不舒服——任何你觉得'今天去不了铁馆'的日子。"
$homeUse+="用多久：15分钟。不要超过20分钟——在家训练的目的不是'练到位' 是'保持习惯不中断'。"
$homeUse+="如果连续好几天都在家用方案A——正常。连续一周都是方案A 也比完全不练强。"
$homeUse+="如果这周方案A用了3次以上——下周试着去一次铁馆 哪怕方案B精简版30分钟。去了就完成。"
foreach($hu in $homeUse){SetC $ws3 $r 1 $hu $false 10 0 0 $true; MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws3.Range("A:A").ColumnWidth=10; $ws3.Range("B:B").ColumnWidth=26
$ws3.Range("C:H").ColumnWidth=16
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 渐进与规则
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="渐进与规则"

SetC $ws4 1 1 "每周怎么进阶  怎么判断今天练不练" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34
SetC $ws4 2 1 "先加动作数 再加次数 最后加重。判断规则帮你决定今天选ABC哪个方案。" $false 10 $headerBlue $lightGray $true
MC $ws4 2 1 8

$r=4
SetC $ws4 $r 1 "4周进阶——稳就行 不追求快" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "方案C完整版" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "方案B精简版" $true 10 $white $headerBlue $false
SetC $ws4 $r 5 "有氧" $true 10 $white $headerBlue $false
SetC $ws4 $r 6 "完成标准" $true 10 $white $headerBlue $true; $r++

$prog=@()
$prog+=,@("1","适应","4-5个动作 中等重量 每组12-15次","2-3个核心动作 中等重量","20min 快走/椭圆机","去了就完成")
$prog+=,@("2","稳定","重量+1片（轻松的话）","重量+1片（轻松的话）","20min 保持","去了就完成")
$prog+=,@("3","加量","每个动作+1组（3变4）或次数+2-3次","从2-3个动作加到3-4个","20min 保持","多做一组/一个动作=完成")
$prog+=,@("4","减载","每个动作2组 轻重量","每个动作2组 轻重量","有氧减半或不强求","测体重腰围 和第1天对比")
foreach($pw in $prog){
    SetC $ws4 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $pw[1] $true 10 $white $darkBg $false
    SetC $ws4 $r 3 $pw[2] $false 10 0 0 $false
    SetC $ws4 $r 4 $pw[3] $false 10 0 0 $false
    SetC $ws4 $r 5 $pw[4] $false 10 0 0 $false
    SetC $ws4 $r 6 $pw[5] $false 10 0 0 $true
    $ws4.Range("A$($r):H$($r)").RowHeight=28; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "今天选哪个方案——根据状态判断" $true 14 $white $softBlue $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

SetC $ws4 $r 1 "你的状态" $true 10 $white $headerBlue $false; MC $ws4 $r 1 2
SetC $ws4 $r 3 "选哪个" $true 10 $white $headerBlue $false
SetC $ws4 $r 5 "练什么" $true 10 $white $headerBlue $false; MC $ws4 $r 5 8; $r++

$colors=@()
$colors+=,@("精力还行 想去铁馆 有50分钟+","方案C 完整版","4-5个动作+20min有氧+拉伸")
$colors+=,@("有点累 但还想去 只有30-40分钟","方案B 精简版","选2-3个核心动作 有氧能做就做")
$colors+=,@("累到不行 完全不想出门","方案A 在家","自重15分钟——做了就完成")
$colors+=,@("脚踝不舒服","不管哪个方案——腿日只用坐姿动作 有氧椭圆机","腿日可跳过 只做上肢日")
$colors+=,@("腰不舒服","不管哪个方案——只用坐姿动作（有靠背）","不站姿 不负重")
foreach($cl in $colors){
    SetC $ws4 $r 1 $cl[0] $true 10 $headerBlue $lightGray $false; MC $ws4 $r 1 2
    SetC $ws4 $r 3 $cl[1] $false 10 0 0 $false
    SetC $ws4 $r 5 $cl[2] $false 10 0 0 $true; MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "训练中遇到这些情况" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "情况" $true 10 $white $headerBlue $false; MC $ws4 $r 1 2
SetC $ws4 $r 3 "怎么做" $true 10 $white $headerBlue $false; MC $ws4 $r 3 8; $r++

$rules=@()
$rules+=,@("训练太轻松 15次做完还能再做5-6个","下次+1片——先加动作数 再加重")
$rules+=,@("训练后48小时还酸痛","正常——刚开始恢复时就会酸。继续练 不是受伤")
$rules+=,@("今天完全不想动","方案A——15分钟在家。做了就完成。")
$rules+=,@("连续好几天都只能做方案A","正常——工作压力是真实的。方案A也比不练强。下周试着去一次铁馆")
$rules+=,@("腰/脚踝疼","疼≠酸。停那个动作。腰疼=全部坐姿。脚踝疼=不上腿日。")
$rules+=,@("今晚有应酬要喝酒","先练再去。应酬前练=消耗一些热量。第二天休息+多喝水")
foreach($ru in $rules){
    SetC $ws4 $r 1 $ru[0] $true 10 $headerBlue $lightGray $false; MC $ws4 $r 1 2
    SetC $ws4 $r 3 $ru[1] $false 10 0 0 $true; MC $ws4 $r 3 8
    $ws4.Range("A$($r):H$($r)").RowHeight=28; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=28; $ws4.Range("B:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 恢复策略
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="恢复策略"

SetC $ws5 1 1 "热身流程  有氧方案  腰脚踝保护" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

$r=3
SetC $ws5 $r 1 "训练前热身——5分钟 必须做" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "上次你练上肢小臂酸——就是没热身+重量大了。这次每次训练前5分钟 按下面做。" $false 10 0 0 $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws5 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws5 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws5 $r 2 3
SetC $ws5 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws5 $r 5 "目的" $true 9 $white $headerBlue $false; MC $ws5 $r 5 8; $r++

$wu=@()
$wu+=,@("1","原地踏步+摆臂","1min","全身升温")
$wu+=,@("2","肩关节绕圈","前10次 后10次","肩关节润滑——推胸下拉前必做")
$wu+=,@("3","摆腿(前后+左右)","各10次/侧","髋关节活动——腿日前必做")
$wu+=,@("4","猫式伸展","8次","脊柱活动——腰舒服就做 不舒服跳过")
$wu+=,@("5","第一个动作轻重量试做","1组x10-12次","最重要！让身体知道接下来做什么——避免小臂酸")
foreach($w in $wu){
    SetC $ws5 $r 1 $w[0] $false 9 0 0 $false
    SetC $ws5 $r 2 $w[1] $true 9 $headerBlue 0 $false; MC $ws5 $r 2 3
    SetC $ws5 $r 4 $w[2] $false 9 0 0 $false
    SetC $ws5 $r 5 $w[3] $false 9 0 0 $true; MC $ws5 $r 5 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

$r+=2
SetC $ws5 $r 1 "有氧方案" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$cardio=@()
$cardio+="  方式：快走（跑步机坡度5-8% 速度4.5-5.5km/h）或椭圆机。脚踝不舒服=椭圆机——脚不离踏板。"
$cardio+="  时间：方案C=力量后20min。方案B=能做就做 不做也行。方案A=原地踏步也算有氧。"
$cardio+="  强度：心率130-140bpm——能断续说话不能流畅聊天。不需要跑——快走就够了。"
$cardio+="  在家有氧：开合跳/原地踏步 10-15min。你之前在家做过的。"
foreach($ca in $cardio){SetC $ws5 $r 1 $ca $false 10 0 0 $true; MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++}

$r+=2
SetC $ws5 $r 1 "腰和脚踝保护——每次训练前检查" $true 14 $white $accentRed $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentRed
$r++

$protect=@()
$protect+="  腰保护：不杠铃硬拉 不杠铃划船——用固定器械（坐姿划船 高位下拉 硬拉器硬拉）代替。所有上身动作首选坐姿（有靠背）。"
$protect+="  脚踝保护：不跳跃（不开合跳 不深蹲跳）不站姿提踵。腿日=倒蹬 哈克深蹲 坐姿腿屈伸——都是坐姿 脚踝无冲击。有氧=椭圆机。"
$protect+="  痛风：多喝水 每天3L。啤酒海鲜火锅——知道但一次戒不掉 先减频次。喝完第二天多喝水 不练。"
$protect+="  训练后拉伸（5min）：胸拉伸+背拉伸+大腿前后拉伸+深呼吸。帮你第二天不酸痛。"
foreach($prt in $protect){SetC $ws5 $r 1 $prt $false 10 0 0 $true; MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws5.Range("A:A").ColumnWidth=14; $ws5.Range("B:H").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 饮食与生活
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="饮食与生活"

SetC $ws6 1 1 "饮食方向  睡眠调整  应酬应对" $true 16 $headerBlue 0 $true
MC $ws6 1 1 8; $ws6.Range("A1:H1").RowHeight=34

$r=3
SetC $ws6 $r 1 "饮食——你第一个月做到过的 不需要重新学" $true 14 $white $accentGreen $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws6 $r 1 "第一个月体重-1.8kg 腰围-3.5cm——饮食坚持得好。最近工作压力大 饮食有点乱——正常。不需要一步到位：先做一顿 再做一天 再做一周。" $false 10 $headerBlue $lightGray $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=28; $r++

SetC $ws6 $r 1 "餐次" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws6 $r 2 4
SetC $ws6 $r 5 "要点" $true 10 $white $headerBlue $false; MC $ws6 $r 5 8; $r++

$meals=@()
$meals+=,@("早餐","2个蛋+牛奶+全麦吐司1片/燕麦","蛋白吃够上午不饿——你之前这么吃的")
$meals+=,@("午餐","米饭一拳+肉/鱼一份+蔬菜多夹","食堂或外卖：选清蒸炖卤 避开油炸。先菜后肉后饭")
$meals+=,@("晚餐","粗粮一拳(糙米/红薯/玉米)+肉/豆腐+蔬菜","粗粮饱腹感更强。没有粗粮=白米饭一拳也可以")
$meals+=,@("应酬时","当天午餐做减法——少油少碳水","喝酒前吃蛋白质垫肚子。啤酒=热量炸弹。第二天多喝水")
$meals+=,@("全天","戒含糖饮料——你第一个月做到了。水3L","痛风要多喝水——帮助尿酸代谢")
foreach($ml in $meals){
    SetC $ws6 $r 1 $ml[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $ml[1] $false 10 0 0 $true; MC $ws6 $r 2 4
    SetC $ws6 $r 5 $ml[2] $false 10 0 0 $true; MC $ws6 $r 5 8
    $ws6.Range("A$($r):H$($r)").RowHeight=36; SB $ws6 $r 1 8; $r++
}

$r+=2
SetC $ws6 $r 1 "睡眠——这个月比多做两次训练更重要" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$sleep=@()
$sleep+="你连续上班快一个月 精神压力大 睡眠不好——这些我都知道。睡眠不好=代谢受影响 减脂效率打折。"
$sleep+="这个月试一件事：睡前一小时不看手机——比多做两次训练更有效。"
$sleep+="睡不着：起来坐一会儿 喝温水 深呼吸。不强迫自己睡着。"
$sleep+="睡眠好的那天=去铁馆方案C。睡不好=方案A在家或只做有氧。睡眠优先于训练。"
foreach($sl in $sleep){SetC $ws6 $r 1 $sl $false 10 0 0 $true; MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++}

$r+=2
SetC $ws6 $r 1 "应酬——80%时间做好就是成功" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$social=@()
$social+="你之前有几次和朋友出去喝酒——正常。你不是职业运动员 你有社交 有朋友——这些不需要放弃。"
$social+="应酬当天午餐做减法。应酬前先去铁馆练——练完再去比直接去强。第二天休息+多喝水。"
$social+="一周21顿饭 2-3顿吃得不完美=只占10%。不影响整体。80%时间做好=成功。"
$social+="痛风注意：啤酒海鲜火锅——知道但一下戒不掉 先减频次。喝一次 后面三天不喝。"
foreach($so in $social){SetC $ws6 $r 1 $so $false 10 0 0 $true; MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws6.Range("A:A").ColumnWidth=22; $ws6.Range("B:H").ColumnWidth=18
Write-Host "Sheet 6 done"

# ============================================
# SHEET 7: 备用与复盘
# ============================================
$ws7=$wb.Worksheets.Add();$ws7.Name="备用与复盘"

SetC $ws7 1 1 "4周复盘  和第一个月对比  最后的话" $true 16 $headerBlue 0 $true
MC $ws7 1 1 8; $ws7.Range("A1:H1").RowHeight=34

$r=3
SetC $ws7 $r 1 "你的第一个月——数据就在这里" $true 14 $white $accentGreen $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws7 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws7 $r 2 "5月7日" $true 10 $white $headerBlue $false
SetC $ws7 $r 3 "6月7日" $true 10 $white $headerBlue $false
SetC $ws7 $r 4 "变化" $true 10 $white $headerBlue $false
SetC $ws7 $r 5 "这个月合理目标" $true 10 $white $headerBlue $false; MC $ws7 $r 5 8; $r++

$data=@()
$data+=,@("体重","107kg","105.2kg","-1.8kg","再减1-2kg——不需要多")
$data+=,@("腰围","116.5cm","113cm","-3.5cm","再减2-3cm")
$data+=,@("胸围","123cm","121cm","-2cm","保持或微减")
$data+=,@("手臂","40.5cm","39.8cm","-0.7cm","保持——不要掉太多")
$data+=,@("大腿","70.5cm(左) 70cm(右)","69.7cm(左) 68.2cm(右)","-0.8cm","保持")
$data+=,@("小腿","44cm","43.2cm","-0.8cm","保持")
foreach($dt in $data){
    SetC $ws7 $r 1 $dt[0] $true 10 $headerBlue $lightGray $false
    SetC $ws7 $r 2 $dt[1] $false 10 0 0 $false
    SetC $ws7 $r 3 $dt[2] $false 10 0 0 $false
    SetC $ws7 $r 4 $dt[3] $false 10 0 0 $false
    SetC $ws7 $r 5 $dt[4] $false 10 0 0 $true; MC $ws7 $r 5 8
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

$r++
SetC $ws7 $r 1 "第一个月体重-1.8kg 腰围-3.5cm——你已经证明了自己能做到。这个月不需要那么快 稳就行。" $false 10 $accentGreen 0 $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=24; $r++

$r+=2
SetC $ws7 $r 1 "4周后复盘——看你变了多少" $true 14 $white $darkBg $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws7 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws7 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws7 $r 2 4
SetC $ws7 $r 5 "下周期" $true 10 $white $headerBlue $false; MC $ws7 $r 5 8; $r++

$review=@()
$review+=,@("体重","减了1-2kg？","减了=继续。没减=查饮食和睡眠——先修睡眠")
$review+=,@("腰围","小了2-3cm？","小了=继续。没小=查有氧频率")
$review+=,@("训练频率","方案A/B/C各用了多少次？","C多=很好。A多=工作压力大——正常。下周期保持")
$review+=,@("睡眠","睡前一小时不看手机做到没？","做到了=最大进步。没做到=下周期继续试")
$review+=,@("照片","第1天vs第28天","镜子比秤真实——这个月拍一组")
foreach($rv in $review){
    SetC $ws7 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws7 $r 2 $rv[1] $false 10 0 0 $true; MC $ws7 $r 2 4
    SetC $ws7 $r 5 $rv[2] $false 10 0 0 $true; MC $ws7 $r 5 8
    $ws7.Range("A$($r):H$($r)").RowHeight=26; SB $ws7 $r 1 8; $r++
}

$r+=2
SetC $ws7 $r 1 "最后的话——给你自己" $true 14 $white $softPink $true
MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=26
$ws7.Range("A$($r):H$($r)").Font.Color=$white; $ws7.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

$final=@()
$final+="你第一个月证明了你能改变——数据不会骗人。最近工作太累了 连续上班快一个月 睡眠不好——这些不是你的错。"
$final+="不要因为'这周只做了方案A'就觉得自己不行。这个方案的核心就是：做减法。能做多少做多少——做了就完成。"
$final+="你问过我'我是不是坚持不下去了'——你还在问这个问题 就说明你还在坚持。"
$final+="不管续不续课 我都希望你练得好。如果后面有什么问题——随时找我。"
foreach($fn in $final){SetC $ws7 $r 1 $fn $false 10 0 0 $true; MC $ws7 $r 1 8; $ws7.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws7.Range("A:A").ColumnWidth=24; $ws7.Range("B:H").ColumnWidth=18
Write-Host "Sheet 7 done"

# Save
$savePath="D:\Codex\members\小冯\小冯_减脂方案_V2.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
