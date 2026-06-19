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

SetC $ws0 1 1 "小冯  减脂训练方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-19  |  107kg  减脂  |  固定器械为主  腰痛保护  从私教到自己练" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你已经有了一个不错的开始" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "一个月前你107kg 现在105.2kg。腰围从116.5cm减到113cm。胸围从123到121。手臂40.5到39.8。这一个月我们大概练了10次 饮食你也做到过——你知道自己能做到。最近工作太累了 连续上班快一个月 睡眠不好 痛风脚踝不舒服——这些我都知道。这份方案就是给你这个阶段准备的：压得上来就多练 压不上来就少练——不管怎样 有一条路。" $false 10 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=48; $row++

$row+=2
SetC $ws0 $row 1 "  现在这个阶段——别逼自己 做减法" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$situation=@(
    @("你不需要","每天逼自己下班练到力竭——你工作已经够累了。一周2-3次就够了 多做一次算赚的 少做一次不亏。"),
    @("你只需要","固定器械——你之前都跟我用过 安全 一个人敢做。杠铃哑铃先不碰——等动作稳了再说。"),
    @("你只需要","保护好腰——杠铃硬拉和划船暂时不做 用固定器械替代。脚踝不舒服就跳过站姿动作。"),
    @("你只需要","饮食做减法——不要求每天带饭。能做一顿是一顿 做不到不批评自己。"),
    @("你只需要","睡前一小时不看手机——你睡眠不好 这个比训练重要。")
)
foreach($st in $situation){
    SetC $ws0 $row 1 $st[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $st[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; $row++
}

$row+=2
SetC $ws0 $row 1 "  训练结构——三分化 + 每周2-4练 按你状态来" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "内容" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "强度" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$split=@()
$split+=,@("训练A","推日（胸+肩+三头）","坐姿推胸+推肩+臂屈伸+飞鸟+下压+有氧20min","中等","胸肌泵感——推起来有劲")
$split+=,@("训练B","拉日（背+二头）","高位下拉+坐姿划船+面拉+弯举+有氧20min","中等","背部收紧——肩胛骨夹紧的感觉")
$split+=,@("训练C","腿日（股四+臀+小腿）","腿举+腿屈伸+臀推+髋外展+提踵+有氧20min","中等","大腿有劲——注意腰和脚踝不疼")

foreach($sp in $split){
    SetC $ws0 $row 1 $sp[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $sp[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $sp[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $sp[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $sp[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  每周训练频率——根据你的状态选" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$freq=@()
$freq+=,@("  绿色周（精力好 睡得还行）","4练：周一A 周二B 周四C 周五A 下周周一B 周二C 周四A 周五B——轮换","每次练完都有氧20min")
$freq+=,@("  黄色周（工作累 但还想去）","3练：周一A 周三B 周五C","有氧能做就做 不做也行")
$freq+=,@("  红色周（连续加班 累到不行）","2练：选两个最想练的 比如A和C","有氧=散步 不算训练")
foreach($fr in $freq){
    SetC $ws0 $row 1 $fr[0] $true 10 $headerBlue $lightGray $false; MC $ws0 $row 1 2
    SetC $ws0 $row 3 $fr[1] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周总览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "训练安排" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "有氧" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周重点" $true 10 $white $headerBlue $false; $row++

$o4=@()
$o4+=,@("1","适应——重建节奏","3-4次 三分化轮换 轻-中重量","每次20min快走/椭圆机 心率130-140","找到节奏——不追求重量 追求'今天去练了'")
$o4+=,@("2","稳定——把频率稳住","3-4次 重量比第1周重1片","有氧保持","如果上周做了4次=这周继续。做了3次=试试能不能做4次")
$o4+=,@("3","加量——容量提升","3-4次 每个动作+1组 或次数+2-3次","有氧保持","重量不加重——只加组数和次数。身体开始适应了")
$o4+=,@("4","减载——恢复+测试","2-3次 轻重量 换动作","有氧减半或不强求","测体重+腰围。和第1天对比——看一个月变了多少")
foreach($oi in $o4){
    SetC $ws0 $row 1 $oi[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $oi[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $oi[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $oi[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $oi[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  腰痛和脚踝保护——每节课的必修" $true 14 $white $accentRed $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentRed
$row++

$protect=@()
$protect+=,@("  腰保护","不杠铃硬拉 不杠铃划船——用坐姿划船+高位下拉+面拉代替。所有坐姿动作腰都有靠背。")
$protect+=,@("  脚踝保护","不站姿提踵——改坐姿。不跳跃——改快走和椭圆机。脚踝不舒服就跳过腿日。")
$protect+=,@("  热身必须做","上次你练上肢小臂酸——那是没热身直接上重量了。每次训练前5分钟热身（见Sheet恢复策略）。")
foreach($prt in $protect){
    SetC $ws0 $row 1 $prt[0] $true 10 $accentRed 0 $false
    SetC $ws0 $row 2 $prt[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@()
$idx+=,@("完整方案逻辑+为什么这样设计"," Sheet 方案概览")
$idx+=,@("训练A推日 B拉日 C腿日 三个训练计划"," Sheet 训练计划")
$idx+=,@("每月变化+怎么判断该加重还是该休息"," Sheet 渐进与规则")
$idx+=,@("热身流程+有氧+恢复+腰痛脚踝保护"," Sheet 恢复策略")
$idx+=,@("饮食怎么吃+睡眠+应酬应对"," Sheet 饮食与生活")
$idx+=,@("精简版+4周后怎么复盘"," Sheet 备用与复盘")
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

SetC $ws1 1 1 "小冯  减脂训练方案  概览" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "男 仓储物流管理 107kg 减脂 固定器械为主 腰痛+脚踝保护 从私教到自己练" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ovl=@()
$ovl+=,@("训练频率","每周2-4练 按你的状态选（见封面绿色/黄色/红色周）","不需要每周4练——2练也有效")
$ovl+=,@("训练结构","三分化（推日/拉日/腿日） 全部固定器械","你跟我用过这些器械——不需要重新学")
$ovl+=,@("每次时长","50-60分钟 力量35min+有氧20min+热身5min","时间不够就砍有氧——力量优先")
$ovl+=,@("周期结构","第1周=适应 第2周=稳定 第3周=加量 第4周=减载","第4周测体重腰围——和第一个月对比")
$ovl+=,@("核心约束","腰痛=不杠铃硬拉/划船。脚踝=不跳跃/站姿提踵","固定器械=安全 一个人也敢做")
$ovl+=,@("方案特点","绿色/黄色/红色三档灵活选择 固定器械安全 腰痛脚踝保护 睡眠饮食调整")
$r=5
foreach($o in $ovl){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "为什么这样设计——这个阶段 做减法比做加法重要" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$why=@(
    "你现在的状态不是第一个月那种'充满干劲'——连续上班快一个月 精神压力大 睡眠不好 偶尔应酬。这个时候最怕的不是练得少——是逼自己练太多然后彻底放弃。",
    "所以这个方案的最大特点：绿色周（精力好）=4练 黄色周（有点累）=3练 红色周（累到不行）=2练。每周你自己选颜色——不管选哪个 都是'完成'。",
    "全部用固定器械——你之前跟我用过 轨道固定 不会做错 一个人也安全。杠铃硬拉和划船腰痛就不做——用高位下拉和坐姿划船代替。脚踝不舒服就改坐姿提踵和椭圆机。",
    "你的第一个月已经证明了你能改变——体重-1.8kg 腰围-3.5cm。这个月不需要那么快 稳就行。"
)
foreach($w in $why){
    SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "腰痛和脚踝保护——每节课必修" $true 14 $white $accentRed $true
MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=26
$ws1.Range("A$($r):I$($r)").Font.Color=$white; $ws1.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

$back=@()
$back+="  腰保护：不杠铃硬拉 不杠铃划船——用坐姿划船+高位下拉+面拉代替。所有坐姿动作腰都有靠背。"
$back+="  脚踝保护：不站姿提踵 不跳跃——改坐姿提踵和快走/椭圆机。训练前检查脚踝——不舒服就改练推日和拉日。"
$back+="  热身必须做：上次你练上肢小臂酸——是没热身上重量了。每次开练前5分钟（关节活动+轻重量试做）。"
$back+="  通风/痛风：多喝水 每天3L。啤酒海鲜火锅——知道但做不到的话 先减频次 不用一步到位。"
foreach($ba in $back){SetC $ws1 $r 1 $ba $false 10 0 0 $true; MC $ws1 $r 1 9; $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$prl=@()
$prl+=,@("1","做减法比做加法重要","这个月不需要'冲'——需要'稳'。绿色4练 黄色3练 红色2练——做到哪个都是完成。不追求完美。")
$prl+=,@("2","固定器械 = 安全","你之前跟我用过——不需要学。自己一个人也敢做。杠铃哑铃暂时不碰——等动作模式稳了再加。")
$prl+=,@("3","睡眠比训练重要","你工作压力大 睡眠不好——这个月如果能改善睡眠 效果比多做两次训练还好。睡前一小时不看手机。")
foreach($p in $prl){
    SetC $ws1 $r 1 $p[0] $false 10 0 0 $false
    SetC $ws1 $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws1 $r 3 $p[2] $false 10 0 0 $true; MC $ws1 $r 3 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$ws1.Range("A:A").ColumnWidth=30; $ws1.Range("B:I").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 训练计划
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划——三分化 推日/拉日/腿日" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "绿色周=4练(A/B/C/A轮换) 黄色周=3练(A/B/C各一次) 红色周=2练(选两个做)。全部固定器械 腰不参与 脚踝不勉强。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# 训练A: 推日
$r=4
SetC $ws2 $r 1 "训练A  推日（胸+肩+三头）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=3组 中等重量 | 第2周=3组 +1片 | 第3周=3组 +1组 | 第4周=2组 轻量" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量参考" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰/脚踝" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsA=@()
$actsA+=,@("1","坐姿推胸","3组","12-15次","15-20kg 选一个做完15次有点累的","60s","  腰有靠背","后背贴紧靠背 推到前面不锁肘。上次你试过——找胸肌发力的感觉。")
$actsA+=,@("2","坐姿推肩","3组","12-15次","最轻片起步","60s","  腰有靠背","下背贴紧靠背 不反弓。推到上面不锁肘。")
$actsA+=,@("3","坐姿臂屈伸","3组","12-15次","轻-中重量","60s","  腰有靠背","大臂固定身体两侧 压到底微微外旋。手臂后侧收紧。")
$actsA+=,@("4","绳索飞鸟","3组","15次","轻重量","45s","  站姿但腰安全","手臂微屈 胸肌驱动 不是手臂拉。想象抱一棵大树。")
$actsA+=,@("5","绳索下压","3组","12-15次","轻-中重量","45s","  站姿但腰安全","大臂固定不外摆 压到底。上次你练手臂小臂酸——减轻重量 先找感觉。")
$actsA+=,@("6","有氧","1组","快走/椭圆机 20min","心率130-140","—","脚踝不舒服=椭圆机","快走=跑步机坡度5-8% 速度4.5-5.5km/h。能断续说话 不能流畅聊天。")

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

# 训练B: 拉日
SetC $ws2 $r 1 "训练B  拉日（背+二头）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=3组 中等重量 | 第2周=3组 +1片 | 第3周=3组 +1组 | 第4周=2组 轻量" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量参考" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰/脚踝" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsB=@()
$actsB+=,@("1","高位下拉","3组","12-15次","20-25kg 选12-15RM","60s","  腰不参与","先沉肩再下拉——不要用手臂硬拽。杆拉到锁骨。你上次练过的。")
$actsB+=,@("2","坐姿划船","3组","12-15次","26-33kg 轻-中","60s","  腰有靠背","身体不后仰 拉到腹部 肩胛骨往中间夹。不杠铃划船——腰受不了。")
$actsB+=,@("3","面拉(绳索)","3组","15次","轻重量","45s","  腰不参与","拉向额头 肘外展。肩后束——长时间坐着上班的人最需要练这个。")
$actsB+=,@("4","哑铃弯举","3组","12-15次","4-8kg 轻-中","45s","  腰不参与","大臂固定 顶峰外旋。你练过的——不要借力甩。")
$actsB+=,@("5","有氧","1组","快走/椭圆机 20min","心率130-140","—","脚踝不舒服=椭圆机","和推日一样——力量后直接做。")

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
$r+=2

# 训练C: 腿日
SetC $ws2 $r 1 "训练C  腿日（股四头+臀+小腿）" $true 13 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=28
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=3组 中等重量 | 第2周=3组 +1片 | 第3周=3组 +1组 | 第4周=2组 轻量  |  注意腰和脚踝" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "重量参考" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "腰/脚踝" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "动作要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):H$($r)").RowHeight=26; $r++

$actsC=@()
$actsC+=,@("1","坐姿腿举","3组","12-15次","空机或最轻片起步","60s","  腰有靠背 脚踝无冲击","脚踩实与肩同宽 膝盖不完全蹬直。你练过的——找大腿发力的感觉。")
$actsC+=,@("2","坐姿腿屈伸","3组","12-15次","2-3片 轻-中","60s","  腰有靠背 脚踝无冲击","脚尖回勾 大腿前侧发力。膝盖不舒服就减重量。")
$actsC+=,@("3","臀桥(哑铃)","3组","15次","持哑铃6-10kg放髋部","60s","  零腰压 零脚踝","脚跟着地 推起时夹屁股。腰不受力——放心做。")
$actsC+=,@("4","坐姿髋外展","3组","15-20次","3-4片 轻-中","45s","  腰有靠背","身体贴紧靠背不晃 臀外侧发力。改善胯和大腿外侧线条。")
$actsC+=,@("5","坐姿提踵","3组","15-20次","轻-中重量","30s","  坐姿=脚踝比站姿安全","顶峰停1秒 慢放。脚踝不舒服就跳过——用椭圆机代替。")
$actsC+=,@("6","有氧","1组","快走/椭圆机 20min","心率130-140","—","脚踝不舒服=椭圆机","和推日拉日一样。腿日做完有氧=今天消耗最大化。")

$i=1
foreach($a in $actsC){
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
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=18; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=14; $ws2.Range("H:H").ColumnWidth=50
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进与规则
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进与规则"

SetC $ws3 1 1 "4周渐进  怎么判断加重还是休息  颜色周选择" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34
SetC $ws3 2 1 "你不追求大重量——追求'持续去练'。渐进从加组数开始 不是从加重量开始。判断规则帮你决定今天练不练。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 8

$r=4
SetC $ws3 $r 1 "4周渐进——加组数优先 加重量在后" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "重量" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "有氧" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "本周目标" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "完成标准" $true 10 $white $headerBlue $true; $r++

$prog=@()
$prog+=,@("1","适应","每个动作3组","中等——选一个做完15次有点累的","20min 快走/椭圆机","重建节奏——不求重量","去了就完成 练了几个动作不重要")
$prog+=,@("2","稳定","3组","比第1周重1片（如果轻松的话）","20min 保持","把频率稳住——上周3次这周试试4次","和上周同样次数=完成")
$prog+=,@("3","加量","每个动作+1组（3变4）","不加重——先加组数","20min 保持","身体适应——同样重量做更多组=进步","多加一组=完成 做不完=也正常")
$prog+=,@("4","减载","每个动作2组","轻重量——比第1周还轻","有氧减半或不做","让身体恢复 测体重腰围 和第1天对比","去了就完成")
foreach($pw in $prog){
    SetC $ws3 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pw[1] $true 10 $white $darkBg $false
    SetC $ws3 $r 3 $pw[2] $false 10 0 0 $false
    SetC $ws3 $r 4 $pw[3] $false 10 0 0 $false
    SetC $ws3 $r 5 $pw[4] $false 10 0 0 $false
    SetC $ws3 $r 6 $pw[5] $false 10 0 0 $false
    SetC $ws3 $r 7 $pw[6] $false 10 0 0 $true
    $ws3.Range("A$($r):H$($r)").RowHeight=30; SB $ws3 $r 1 8; $r++
}

# 颜色周选择
$r+=2
SetC $ws3 $r 1 "每周选颜色——根据你的状态 不是根据计划" $true 14 $white $softBlue $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "你的状态" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "选哪个" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "练几次" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "完成标准" $true 10 $white $headerBlue $false; MC $ws3 $r 6 8; $r++

$colors=@()
$colors+=,@("这周精力还行 睡眠还可以 工作不算太累","  绿色周","4次（A/B/C/A轮换）","至少3次=完成")
$colors+=,@("这周有点累 但还想去练","  黄色周","3次（A/B/C各一次）","至少2次=完成")
$colors+=,@("这周连续加班 累到不行 不想动","  红色周","2次（选两个最想练的）","1次=完成。0次也没关系。")
$colors+=,@("脚踝不舒服","不管什么颜色周——腿日跳过","只做推日和拉日","做了就完成")
$colors+=,@("腰不舒服","不管什么颜色周——用坐姿动作","全部动作都要有靠背","做了就完成")
foreach($cl in $colors){
    SetC $ws3 $r 1 $cl[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $cl[1] $false 10 0 0 $false
    SetC $ws3 $r 4 $cl[2] $false 10 0 0 $false
    SetC $ws3 $r 6 $cl[3] $false 10 0 0 $true; MC $ws3 $r 6 8
    $ws3.Range("A$($r):H$($r)").RowHeight=26; SB $ws3 $r 1 8; $r++
}

# 判断规则
$r+=2
SetC $ws3 $r 1 "怎么判断——训练中遇到这些情况" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false; MC $ws3 $r 1 2
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "怎么做" $true 10 $white $headerBlue $false; MC $ws3 $r 5 8; $r++

$rules=@()
$rules+=,@("训练太轻松 15次做完还能再做5个","可以加重了","下次+1片——先加组数 再加重")
$rules+=,@("训练后48小时还酸痛","正常——刚开始练时会酸。多喝水多睡","下次照样练——酸痛不是'受伤'")
$rules+=,@("今天完全不想动","—","去铁馆待20分钟——哪怕只做热身+1个动作+拉伸。做了就完成")
$rules+=,@("连续2周都是红色周","正常——你的工作压力是真实的","继续保持红色周——不要因为'这周又是红的'就自责")
$rules+=,@("腰/脚踝疼","疼≠酸。疼=停","停这个动作。腰疼=只做坐姿。脚踝疼=不练腿日")
$rules+=,@("今晚有应酬 要喝酒","—","今天去练。应酬前练=消耗一些热量 喝酒后第二天休息")
foreach($ru in $rules){
    SetC $ws3 $r 1 $ru[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $ru[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $ru[2] $false 10 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=28; SB $ws3 $r 1 8; $r++
}

$ws3.Range("A:A").ColumnWidth=28; $ws3.Range("B:H").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "热身流程  有氧方案  腰痛脚踝保护" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

$r=3
SetC $ws4 $r 1 "训练前热身（5分钟 必须做——上次小臂酸就是因为没热身）" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws4 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws4 $r 2 3
SetC $ws4 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws4 $r 5 "目的" $true 9 $white $headerBlue $false; MC $ws4 $r 5 8; $r++

$wu=@()
$wu+=,@("1","原地踏步+摆臂","1min","全身升温")
$wu+=,@("2","肩关节绕圈","前10次 后10次","肩关节润滑——推胸下拉前必做")
$wu+=,@("3","摆腿(前后+左右)","各10次/侧","髋关节活动——腿日前必做")
$wu+=,@("4","猫式伸展","8次","脊柱活动——腰舒服就做 不舒服跳过")
$wu+=,@("5","第一个动作轻重量试做","1组x10-12次","让身体知道接下来要做什么——最重要的一步！")
foreach($w in $wu){
    SetC $ws4 $r 1 $w[0] $false 9 0 0 $false
    SetC $ws4 $r 2 $w[1] $true 9 $headerBlue 0 $false; MC $ws4 $r 2 3
    SetC $ws4 $r 4 $w[2] $false 9 0 0 $false
    SetC $ws4 $r 5 $w[3] $false 9 0 0 $true; MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

$r+=2
SetC $ws4 $r 1 "有氧方案——力量后直接做" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$cardio=@()
$cardio+="  有氧方式：快走（跑步机坡度5-8% 速度4.5-5.5km/h）或椭圆机（阻力5-8 速度自调）。脚踝不舒服=椭圆机——脚不离踏板。"
$cardio+="  时间：每次力量后20min。如果今天时间不够——有氧可以砍 力量优先。绿色周=每次都做。黄色周=能做就做。红色周=散步就算。"
$cardio+="  强度：心率130-140bpm——能断续说话但不能流畅聊天。不需要跑到喘——快走就够了。"
$cardio+="  在家有氧：你之前在家练开合跳和深蹲——可以。但脚踝不舒服就改成快走或椭圆机。"
foreach($ca in $cardio){SetC $ws4 $r 1 $ca $false 10 0 0 $true; MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++}

$r+=2
SetC $ws4 $r 1 "腰痛和脚踝保护——每次训练前检查" $true 14 $white $accentRed $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentRed
$r++

$protect=@()
$protect+="  腰保护：不杠铃硬拉 不杠铃划船——用固定器械代替。所有动作首选坐姿（有靠背）。站姿动作（飞鸟 下压 弯举）=核心收紧 腰不弓。"
$protect+="  脚踝保护：不站姿提踵（改坐姿）。不跳跃（不开合跳 不深蹲跳）。不快走改椭圆机。训练前转转脚踝——不舒服就不练腿日。"
$protect+="  通风/痛风：多喝水 每天3L。啤酒海鲜火锅——知道但一次戒不掉 先减频次。喝完第二天多喝水 不练。"
$protect+="  训练后拉伸（5min）：胸拉伸+背拉伸+大腿前后拉伸+深呼吸。帮你第二天不酸痛——你之前没怎么做拉伸。"
foreach($prt in $protect){SetC $ws4 $r 1 $prt $false 10 0 0 $true; MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws4.Range("A:A").ColumnWidth=14; $ws4.Range("B:B").ColumnWidth=26
$ws4.Range("C:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食与生活
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与生活"

SetC $ws5 1 1 "饮食方向  睡眠调整  应酬应对" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

$r=3
SetC $ws5 $r 1 "饮食——你之前做到过的 不需要重新学" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你第一个月饮食坚持得很好——体重-1.8kg 腰围-3.5cm就是证明。最近工作压力大 饮食有点乱——正常。不需要一步到位回到之前的水准：先做一顿 再做一天 再做一周。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=32; $r++

SetC $ws5 $r 1 "餐次" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws5 $r 2 4
SetC $ws5 $r 5 "要点" $true 10 $white $headerBlue $false; MC $ws5 $r 5 8; $r++

$meals=@()
$meals+=,@("早餐","2个蛋+牛奶+全麦吐司1片/燕麦","蛋白吃够上午不饿——你之前这么吃的")
$meals+=,@("午餐","米饭一拳+肉/鱼一份+蔬菜多夹","食堂或外卖：选清蒸炖卤 避开油炸。先菜后肉后饭")
$meals+=,@("晚餐","粗粮一拳(糙米/红薯/玉米)+肉/豆腐+蔬菜","晚餐换粗粮——饱腹感更强 睡前不饿。没有粗粮=白米饭一拳也可以")
$meals+=,@("有应酬时","当天午餐做减法——少油少碳水","喝酒前吃蛋白质垫肚子。啤酒=热量炸弹。喝完第二天多喝水")
$meals+=,@("全天","戒含糖饮料——你第一个月做到了。水3L","痛风要多喝水——帮助尿酸代谢")
foreach($ml in $meals){
    SetC $ws5 $r 1 $ml[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $ml[1] $false 10 0 0 $true; MC $ws5 $r 2 4
    SetC $ws5 $r 5 $ml[2] $false 10 0 0 $true; MC $ws5 $r 5 8
    $ws5.Range("A$($r):H$($r)").RowHeight=36; SB $ws5 $r 1 8; $r++
}

$r+=2
SetC $ws5 $r 1 "睡眠——这个月比训练更重要的事" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$sleep=@()
$sleep+="你连续上班快一个月 精神压力大 睡眠不好——这些我都知道。睡眠不好=代谢受影响 减脂效率打折扣 训练恢复变慢。"
$sleep+="这个月做一件事：睡前一小时不看手机——这比多做两次训练更有效。"
$sleep+="如果晚上睡不着：起来坐一会儿 不要躺着焦虑。喝温水 做几个深呼吸。不强迫自己睡着——越强迫越睡不着。"
$sleep+="睡眠好的那天=练。睡眠差的那天=不练或只做有氧。睡眠优先于训练。"
foreach($sl in $sleep){SetC $ws5 $r 1 $sl $false 10 0 0 $true; MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++}

$r+=2
SetC $ws5 $r 1 "应酬——不做完美主义 只做80%" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$social=@()
$social+="你之前有几次和朋友出去喝酒——正常。你不是职业运动员 你有社交 有朋友——这些不需要放弃。"
$social+="应酬当天午餐做减法 应酬前先去练——练完再去喝酒 比你直接去喝酒强。第二天休息 多喝水。"
$social+="一周21顿饭 有2-3顿吃得不完美=只占10%。不影响整体。80%时间做好=成功。"
$social+="痛风注意：啤酒海鲜火锅——知道但一下子戒不掉 先减频次。喝一次啤酒 后面三天不喝。"
foreach($so in $social){SetC $ws5 $r 1 $so $false 10 0 0 $true; MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws5.Range("A:A").ColumnWidth=22; $ws5.Range("B:E").ColumnWidth=20; $ws5.Range("F:H").ColumnWidth=16
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用与复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  4周复盘  和第一个月对比" $true 16 $headerBlue 0 $true
MC $ws6 1 1 8; $ws6.Range("A1:H1").RowHeight=34

$r=3
SetC $ws6 $r 1 "如果完全不想去铁馆——在家版" $true 14 $white $accentOrange $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

$homeList=@()
$homeList+="完全不想出门的日子——在家做这些："
$homeList+=" 1. 自重深蹲（扶椅背）3x15次"
$homeList+=" 2. 跪姿俯卧撑 3x力竭（你做得了几个做几个）"
$homeList+=" 3. 仰卧举腿 3x15次"
$homeList+=" 4. 开合跳/原地踏步 10min + 拉伸"
$homeList+="花20分钟——做了就完成。你之前在家练过 这个你熟。"
foreach($hm in $homeList){SetC $ws6 $r 1 $hm $false 10 0 0 $true; MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++}

$r+=2
SetC $ws6 $r 1 "4周后复盘——和你第一个月的数据对比" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "你的第一个月数据" $true 12 $headerBlue $lightBlue $true; MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws6 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "5月7日" $true 10 $white $headerBlue $false
SetC $ws6 $r 3 "6月7日" $true 10 $white $headerBlue $false
SetC $ws6 $r 4 "变化" $true 10 $white $headerBlue $false
SetC $ws6 $r 5 "第4周目标" $true 10 $white $headerBlue $false; MC $ws6 $r 5 8; $r++

$data=@()
$data+=,@("体重","107kg","105.2kg","-1.8kg","再减2-3kg")
$data+=,@("腰围","116.5cm","113cm","-3.5cm","再减3-4cm")
$data+=,@("胸围","123cm","121cm","-2cm","保持或微减")
$data+=,@("手臂","40.5cm","39.8cm","-0.7cm","保持——不要掉太多")
$data+=,@("大腿","70.5cm","69.7cm","-0.8cm","保持")
foreach($dt in $data){
    SetC $ws6 $r 1 $dt[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $dt[1] $false 10 0 0 $false
    SetC $ws6 $r 3 $dt[2] $false 10 0 0 $false
    SetC $ws6 $r 4 $dt[3] $false 10 0 0 $false
    SetC $ws6 $r 5 $dt[4] $false 10 0 0 $true; MC $ws6 $r 5 8
    $ws6.Range("A$($r):H$($r)").RowHeight=26; SB $ws6 $r 1 8; $r++
}

$r++
SetC $ws6 $r 1 "第一个月体重-1.8kg 腰围-3.5cm——你已经证明了自己能做到。这个月不需要那么快 稳就行。" $false 10 $accentGreen 0 $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=24; $r++

$r+=2
SetC $ws6 $r 1 "4周后复盘——看你变了多少" $true 14 $white $darkBg $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 6 8; $r++

$review=@()
$review+=,@("体重","减了2-3kg？","减了=继续。没减=查饮食和睡眠——先修睡眠")
$review+=,@("腰围","小了3-4cm？","小了=继续。没小=检查有氧频率——是不是每次都跳了")
$review+=,@("训练频率","平均每周几次？绿色/黄色/红色的比例？","红色过多=工作压力太大——不是你的问题。下周期保持同样节奏就行")
$review+=,@("睡眠","比上月有改善吗？睡前一小时不看手机做到没？","做到了=最大进步。没做到=下周期继续试——不批评自己")
$review+=,@("照片","第1天vs第28天","镜子比秤真实——第一个月你没拍 这个月拍一组")
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 5
    SetC $ws6 $r 6 $rv[2] $false 10 0 0 $true; MC $ws6 $r 6 8
    $ws6.Range("A$($r):H$($r)").RowHeight=28; SB $ws6 $r 1 8; $r++
}

$r+=2
SetC $ws6 $r 1 "最后的话——给你自己" $true 14 $white $softPink $true
MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=26
$ws6.Range("A$($r):H$($r)").Font.Color=$white; $ws6.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

$final=@()
$final+="你第一个月证明了你能改变——数据不会骗人。"
$final+="最近工作太累了 连续上班快一个月 睡眠不好——这些不是你的错。"
$final+="不要因为'这周只练了2次'就觉得自己不行。这个方案的核心就是：做减法。能做多少做多少——做了就完成。"
$final+="你问过我'我是不是坚持不下去了'——你还在问这个问题 就说明你还在坚持。"
foreach($fn in $final){SetC $ws6 $r 1 $fn $false 10 0 0 $true; MC $ws6 $r 1 8; $ws6.Range("A$($r):H$($r)").RowHeight=22; $r++}

$ws6.Range("A:A").ColumnWidth=24; $ws6.Range("B:H").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\小冯\小冯_减脂方案_V1.xlsx"
New-Item -ItemType Directory -Force -Path "D:\Codex\members\小冯" | Out-Null
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
