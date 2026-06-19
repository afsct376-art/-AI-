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
    SetC $ws $ref 1 $title $true 13 $white $darkBg $true; MC $ws $ref 1 9
    $ws.Range("A$($ref):I$($ref)").RowHeight=26
    $ws.Range("A$($ref):I$($ref)").Font.Color=$white
    $ws.Range("A$($ref):I$($ref)").Interior.Color=$darkBg; $ref++
    SetC $ws $ref 1 "$focus  |  $time 分钟  |  $tips" $false 10 $headerBlue $lightGray $true; MC $ws $ref 1 9; $ref++
    SetC $ws $ref 1 "序" $true 10 $white $headerBlue $false
    SetC $ws $ref 2 "动作" $true 10 $white $headerBlue $false
    SetC $ws $ref 3 "组数" $true 10 $white $headerBlue $false
    SetC $ws $ref 4 "次数" $true 10 $white $headerBlue $false
    SetC $ws $ref 5 "强度/%1RM" $true 10 $white $headerBlue $false
    SetC $ws $ref 6 "休息" $true 10 $white $headerBlue $false
    SetC $ws $ref 7 "RPE" $true 10 $white $headerBlue $false
    SetC $ws $ref 8 "进阶/变化" $true 10 $white $headerBlue $false
    SetC $ws $ref 9 "技术要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):I$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        SetC $ws $ref 1 $i $false 10 0 $lightGray $false
        SetC $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 7;$j++){SetC $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 7)}
        $ws.Range("A$($ref):I$($ref)").RowHeight=50
        SB $ws $ref 1 9; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "吉  增肌分化方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  182cm / 82kg  |  增肌  |  给同行的一份方案" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你是教练——这份方案是写给你的" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$forCoach=@(
    "你每天在健身房教别人——但可能很久没有为自己系统做过计划了。这份方案按'同行标准'写。",
    "没有基础动作教学——你都知道。但增加了：DUP波动周期逻辑、RPE进阶体系、主动恢复策略、自主调节框架。",
    "你填了每周2-3练和4练——我按4练/周设计(推/拉/腿/上肢)。如果这周只能2-3练 用精简版。",
    "这套框架你可以用在自己身上——也可以迁移到你给会员做计划的思路里。",
    "",
    "这是为你写的——也是给同行看的一个参考：定制方案应该长这样。"
)
foreach($fc in $forCoach){
    SetC $ws0 $row 1 $fc $false 10 0 0 $true; MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=20
    if($fc -eq ""){$ws0.Range("A$($row):E$($row)").RowHeight=8}
    $row++
}

$row+=2
SetC $ws0 $row 1 "  方案概览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$overview=@(
    @("训练频率","4次/周 推/拉/腿/上肢分化（附2-3练精简版）"),
    @("每次时长","50-60分钟（热身10min+力量35-45+拉伸5min）"),
    @("周期结构","DUP波动周期——每周4天不同主题：大重量/容量/代谢/弱点"),
    @("加重原则","上肢+1-2.5kg / 下肢+2.5-5kg——你是教练 懂这个"),
    @("方案特点","同行标准 DUP波动 自主调节框架 主动恢复策略")
)
foreach($ov in $overview){
    SetC $ws0 $row 1 $ov[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $ov[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; SB $ws0 $row 1 5; $row++
}

# DUP 4-day
$row+=2
SetC $ws0 $row 1 "  DUP 4天分化——每天不同刺激" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "强度区间" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "RPE" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$dup=@(
    @("周一","推日(大重量)","80-87.5% 1RM","8-9","重而可控——主项杠铃卧推冲5x5"),
    @("周二","拉日(容量)","70-77.5% 1RM","7-8","中高强度 较高容量——背阔肌泵感+腘绳肌"),
    @("周四","腿日(代谢)","65-75% 1RM","7-8","短间歇 高次数——深蹲+腿举追求代谢压力"),
    @("周五","上肢日(弱点/容量)","67-77% 1RM","7-8","补推日不足 侧重肩和三头——弱点攻克")
)
foreach($d in $dup){
    SetC $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $d[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $d[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $d[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

# Weekly schedule
$row+=2
SetC $ws0 $row 1 "  每周安排" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$weekly=@(
    @("周一","推日(大重量)","周二","拉日(容量)"),
    @("周四","腿日(代谢)","周五","上肢日(弱点)"),
    @("周三/六/日","休息/主动恢复/有氧","","")
)
foreach($w in $weekly){
    SetC $ws0 $row 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $w[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $w[2] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 4 $w[3] $false 10 0 0 $false; MC $ws0 $row 4 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24; SB $ws0 $row 1 5; $row++
}

# 精简版
$row+=2
SetC $ws0 $row 1 "  如果这周只能2-3练——精简版" $true 14 $white $accentOrange $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentOrange
$row++

$simpleVer=@(
    @("2练","全身A(推+腿) 全身B(拉+上肢)","压缩辅助动作 保留核心复合"),
    @("3练","推+腿 拉+上肢 全身综合","每周每个肌群仍刺激2次")
)
foreach($sv in $simpleVer){
    SetC $ws0 $row 1 $sv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $sv[1] $false 10 0 0 $false; MC $ws0 $row 2 3
    SetC $ws0 $row 4 $sv[2] $false 10 0 0 $true; MC $ws0 $row 4 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+周期框架+RPE指南"," Sheet 方案概览"),
    @("推日 拉日 腿日 上肢日"," Sheet 训练计划"),
    @("渐进负荷+器械突破"," Sheet 渐进突破"),
    @("热身+主动恢复+有氧"," Sheet 恢复策略"),
    @("饮食方向+自主调节框架"," Sheet 饮食与调节"),
    @("精简版+4周复盘"," Sheet 备用与复盘")
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

SetC $ws 1 1 "吉  增肌方案  概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 9; $ws.Range("A1:I1").RowHeight=36

SetC $ws 2 1 "男  25-30岁  182cm/82kg  体脂20-25%  健身教练  目标增肌  商业健身房器械齐全" $false 11 $gold $darkBg $true
MC $ws 2 1 9; $ws.Range("A2:I2").Font.Color=$white; $ws.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 9

$ov=@(
    @("训练频率","4次/周(周一/二/四/五) 推/拉/腿/上肢分化","附2-3练精简版"),
    @("周期结构","DUP波动——大重量/容量/代谢/弱点 每天不同刺激","4周一小周期 每4周末减载"),
    @("每次时长","热身10min+力量35-45min+拉伸5min=50-60min","你是教练——热身可以自己调"),
    @("加重原则","上肢+1-2.5kg/下肢+2.5-5kg/次","RPE优先于重量百分比"),
    @("方案特点","同行标准 DUP波动 主动恢复 自主调节框架")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 6
    SetC $ws $r 7 $o[2] $false 10 0 0 $true; MC $ws $r 7 9
    SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "DUP波动周期——为什么这样设计" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$dupWhy=@(
    "你有训练基础——身体已经适应了线性周期（每周同样的分化方式）。DUP让神经系统始终保持对'新'刺激的敏感。",
    "周一推日大重量(80-87.5%)=神经系统高强度输出。周二拉日容量(70-77.5%)=背部和腘绳积累代谢压力。周四腿日代谢(65-75%短间歇)=下肢在疲劳状态下做功。周五上肢弱点=弥补推日不足+攻克弱项。",
    "每天不同刺激=适应性更慢 进步更持续。你作为教练应该能体会——让会员做同样的计划4周vs每周变节奏 哪个效果更好？",
    "减载周结束后 身体对4种刺激都保持敏感——恢复后同时突破多个维度。"
)
foreach($dw in $dupWhy){
    SetC $ws $r 1 $dw $false 10 0 0 $true; MC $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# RPE — coach level
$r++
SetC $ws $r 1 "RPE参考——你每天都在用 不需要解释 但统一标准" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "RPE" $true 10 $white $headerBlue $false
SetC $ws $r 2 "感觉" $true 10 $white $headerBlue $false
SetC $ws $r 4 "剩余次数" $true 10 $white $headerBlue $false
SetC $ws $r 6 "用在哪天" $true 10 $white $headerBlue $false; $r++

$rpeRef=@(
    @("6-7","中等——做完还能做4-6次","容量/弱点日——泵感和代谢积累"),
    @("7-8","中等偏强——做完还能做3-4次","拉日/腿日——理想的增肌区间"),
    @("8-9","很重——做完还能做1-2次","推日大重量——神经系统高强度输出"),
    @("10","极限力竭——做完一个都起不来","本方案几乎不用。只在第3周峰值周的最后一组偶尔触达。")
)
foreach($rp in $rpeRef){
    SetC $ws $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    SetC $ws $r 2 $rp[1] $false 10 0 0 $false
    SetC $ws $r 4 $rp[2] $false 10 0 0 $false
    SetC $ws $r 6 $rp[3] $false 10 0 0 $true; MC $ws $r 6 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

# Periodization
$r++
SetC $ws $r 1 "4周周期节奏" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws $r 4 "推日变化" $true 10 $white $headerBlue $false
SetC $ws $r 6 "拉/腿/上肢变化" $true 10 $white $headerBlue $false; $r++

$weeks4=@(
    @("1","积累","80%1RM 5x5 建基准","70%1RM 3-4x8-10 建容量基准"),
    @("2","递增","+2.5-5kg RPE8-9 保持5x5","缩短休息10s 或+2.5kg RPE7-8"),
    @("3","突破","再+2.5-5kg 尝试新极限","再缩短休息 或再+2.5kg RPE8"),
    @("4","减载","60%1RM 3x5 维持神经适应","50-60% 2-3x8-10 主动恢复")
)
foreach($w4 in $weeks4){
    SetC $ws $r 1 $w4[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $w4[1] $true 10 $white $darkBg $false
    SetC $ws $r 4 $w4[2] $false 10 0 0 $false; MC $ws $r 4 5
    SetC $ws $r 6 $w4[3] $false 10 0 0 $true; MC $ws $r 6 9
    $ws.Range("A$($r):I$($r)").RowHeight=28; SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$pr=@(
    @("1","这是你自己的训练","你是教练——动作技术不需要别人教。这份方案不是指导你'怎么做'——是给你一个周期框架和进阶逻辑 你在这个框架内自主调节。"),
    @("2","RPE优先于百分比","80% 1RM是理论值。你当天睡没睡好 会员多不多 站了一天累不累——这些都会影响实际表现。RPE让你和身体对话。"),
    @("3","这套框架可以迁移","DUP波动+4天分化+自主调节+RPE进阶——这套逻辑你可以用在自己身上 也可以用在给会员做计划的时候。")
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
# SHEET 2: 训练计划——推/拉/腿/上肢
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "DUP 4天分化  推/拉/腿/上肢" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "周一=推日(大重量) 周二=拉日(容量) 周四=腿日(代谢) 周五=上肢日(弱点)。第1周建基准 第2周递增 第3周突破 第4周减载。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 推日
$r=4
$actsPush=@(
    @("杠铃平板卧推(主项)","5组","5次","80% 1RM RPE 8-9","120-150s","8-9","第2周+2.5kg 第3周再+2.5kg","【呼吸】瓦式呼吸——下放吸气屏住 推起中途呼气。【要点】起桥 肩胛收紧 脚踩实。推日核心=重而可控 每组留1-2次力竭。"),
    @("上斜哑铃卧推","4组","8-10次","RPE 7-8 上胸补充","90s","7-8","上斜角度30-45度","【要点】和杠铃不同——哑铃自由轨迹 上胸和肩前束同时刺激。重量中等 追求上胸泵感。"),
    @("坐姿哑铃推举","4组","8-10次","RPE 7-8","75s","7-8","第2周+1-2kg","【要点】下背贴靠背 不反弓。推举是推日的肩部主项——三角肌前束+中束。"),
    @("双杠臂屈伸(负重可选)","3组","8-12次","自重或负重 RPE 7-8","60s","7-8","第2周+负重腰带","【要点】身体稍前倾=胸+三头。身体竖直=三头为主。推日自选角度。"),
    @("绳索飞鸟","3组","12-15次","轻重量 追求泵感","45s","6-7","—","【要点】手臂微屈 胸肌驱动。推日的收尾动作——在胸肌疲劳状态下继续刺激代谢。"),
    @("三头绳索下压","3组","12-15次","轻-中重量","45s","7","—","【要点】大臂固定 压到底微外旋。推日收尾——大重量卧推后三头已疲劳 这时候下压效果最好。")
)
$r=Write-TrainingBlock $ws2 $r "周一  推日（大重量 Heavy Push）" "杠铃卧推5x5主项+上斜哑铃+推举+臂屈伸+飞鸟+三头  |  45-55" "主项冲重量 辅助追求泵感" "45-55" $actsPush

# 拉日
$actsPull=@(
    @("杠铃俯身划船(主项)","4组","8-10次","RPE 7-8 中高强度","90s","7-8","第2周+2.5kg","【要点】俯身45-60度 核心收紧 杠铃贴大腿前侧拉至肚脐。拉日核心=中高强度+较高容量 背阔肌泵感。"),
    @("引体向上(负重可选)","4组","8-12次","自重或负重 RPE 7-8","90s","7-8","第2周+负重","【要点】全幅度——底部手臂伸直 顶部下巴过杠。做不了8次用弹力带辅助 能做12次以上负重。"),
    @("坐姿绳索划船","3组","10-12次","RPE 7 追求背阔肌收缩","60s","7","—","【要点】V柄拉至腹部 顶峰肩胛后缩1秒。和杠铃划船互补——杠铃=大重量 绳索=精细收缩。"),
    @("罗马尼亚硬拉","4组","8-10次","RPE 7-8 杠铃或哑铃","90s","7-8","第2周+5kg","【要点】膝盖微屈 全程直背。拉日的下肢动作——腘绳肌+臀。和腿日深蹲互补。"),
    @("面拉(绳索)","3组","15次","轻重量","45s","6-7","—","【要点】拉向额头 肘外展。肩后束+上背——拉日的肩部补充。"),
    @("杠铃/哑铃弯举","3组","10-12次","RPE 7","45s","7","—","【要点】大臂固定 顶峰外旋。二头是拉日的收尾——划船和引体后二头已充血 这时候弯举最有效。")
)
$r=Write-TrainingBlock $ws2 $r "周二  拉日（容量 Pull Volume）" "杠铃划船+引体+绳索划船+RDL+面拉+弯举  |  45-55" "中高强度 较高容量 背阔肌泵感+腘绳肌" "45-55" $actsPull

# 腿日
$actsLegs=@(
    @("杠铃深蹲(主项)","4组","8-10次","RPE 7-8 短间歇=90s","90s","7-8","第2周+5kg","【呼吸】瓦式——下蹲吸气屏住 站起呼气。【要点】深度至少大腿平行地面。腿日核心=短间歇+适当重量=代谢压力。"),
    @("腿举","4组","10-12次","RPE 7-8 短间歇=60s","60s","7-8","第2周+10kg","【要点】脚与肩同宽 膝盖不全蹬直。深蹲后股四头已疲劳——腿举在疲劳状态下继续做功。"),
    @("保加利亚分腿蹲","3组","10次/侧","哑铃对侧手持 RPE 7","60s","7","—","【要点】后脚架凳 前脚跟发力。单侧=纠正左右不平衡+核心抗旋转。"),
    @("坐姿腿弯举","3组","12-15次","RPE 7 离心3秒","45s","7","—","【要点】大腿后侧发力 慢放不砸片。腘绳肌孤立——和拉日的RDL互补。"),
    @("站姿提踵","4组","15-20次","RPE 7 膝盖微屈","30s","7","—","【要点】顶峰停1秒 慢放。小腿=代谢收尾——高频次短间歇。"),
    @("悬垂举腿","3组","12-15次","自重","45s","7","—","【要点】下腹带动骨盆上卷 不靠惯性。腿日的核心收尾。")
)
$r=Write-TrainingBlock $ws2 $r "周四  腿日（代谢 Legs Metabolic）" "杠铃深蹲+腿举+分腿蹲+腿弯举+提踵+举腿  |  45-55" "短间歇+中等重量=代谢压力训练" "45-55" $actsLegs

# 上肢日
$actsUpper=@(
    @("哑铃平板卧推(主项)","4组","8-10次","RPE 7-8 和周一不同角度","90s","7-8","第2周+1-2kg/边","【要点】和周一杠铃卧推不同——哑铃自由轨迹 更偏胸肌中束和稳定肌。上肢日=补充推日的容量不足。"),
    @("高位下拉(宽握)","4组","8-10次","RPE 7-8","75s","7-8","第2周+1片","【要点】宽握=上背阔肌。和周二划船不同——一个水平拉一个垂直拉 互补。"),
    @("哑铃侧平举","4组","15-20次","轻重量 RPE 7","45s","7","—","【要点】不耸肩 手臂微屈。上肢日的肩部专项——高次数追求三角肌中束泵感。推日做推举 上肢日做侧平举 肩部两个角度都覆盖。"),
    @("哑铃弯举+三头下压超级组","各3组","各12-15次","RPE 7 短间歇","不休息 组间60s","7-8","—","【要点】弯举→直接做三头下压→休息60秒。上肢日的手臂超级组——代谢+充血同时搞定。"),
    @("面拉(绳索)","3组","15次","轻重量","45s","6-7","—","【要点】肩后束+上背——和周二同样的动作 保持肩关节健康。"),
    @("农夫行走","3组","30-40m","大重量哑铃 单边30-40kg","60s","7-8","—","【要点】躯干正直不侧倾 核心全程收紧。上肢日收尾——握力+斜方+核心=功能性收尾。")
)
$r=Write-TrainingBlock $ws2 $r "周五  上肢日（弱点/容量 Upper Weakness）" "哑铃卧推+高位下拉+侧平举+手臂超级组+面拉+农夫行走  |  40-50" "补推日不足 侧重肩和三头——弱点攻克" "40-50" $actsUpper

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20; $ws2.Range("C:C").ColumnWidth=6
$ws2.Range("D:D").ColumnWidth=10; $ws2.Range("E:E").ColumnWidth=16; $ws2.Range("F:F").ColumnWidth=10
$ws2.Range("G:G").ColumnWidth=8; $ws2.Range("H:H").ColumnWidth=18; $ws2.Range("I:I").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进+突破
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进突破"

SetC $ws3 1 1 "渐进负荷  器械上限突破  进阶判断" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34

$r=3
SetC $ws3 $r 1 "4周渐进（DUP框架内 每周4天主题不变 逐周递增）" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "推日(大重量)" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "拉日(容量)" $true 10 $white $headerBlue $false
SetC $ws3 $r 6 "腿日(代谢)" $true 10 $white $headerBlue $false
SetC $ws3 $r 8 "上肢日(弱点)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("1","80%1RM 5x5","70% 3-4x8-10","65-70% 4x8-10 90s休息","67% 3-4x8-10"),
    @("2","+2.5-5kg 保持5x5","休息-10s或+2.5kg","休息-10s或+5kg","+1-2kg/边或休息-10s"),
    @("3","再+2.5-5kg 接近5RM","休息-10s 或再+2.5kg","休息-10s 或再+5kg 高次数力竭","维持或微加重"),
    @("4","60% 3x5 技术打磨","50-60% 2-3x8-10 换动作","50-60% 2x10 高脚杯/自重","轻量 弹力带 技术优先")
)
foreach($pw in $prog){
    SetC $ws3 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pw[1] $false 9 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pw[2] $false 9 0 0 $true; MC $ws3 $r 4 5
    SetC $ws3 $r 6 $pw[3] $false 9 0 0 $true; MC $ws3 $r 6 7
    SetC $ws3 $r 8 $pw[4] $false 9 0 0 $true; MC $ws3 $r 8 9
    $ws3.Range("A$($r):I$($r)").RowHeight=34; SB $ws3 $r 1 9; $r++
}

# Breakthrough
$r+=2
SetC $ws3 $r 1 "器械上限突破策略" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "策略" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "操作" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "适用" $true 10 $white $headerBlue $false; $r++

$bt=@(
    @("慢离心","3-4s下放+爆发推起","所有日子——增加TUT"),
    @("1.5次法","下放->半起->全下->全起=1次","推日/腿日——双倍刺激"),
    @("缩短休息","90s->60s->45s","腿日(代谢)——核心策略"),
    @("预疲劳","孤立力竭->复合动作","上肢日进阶——轻量化力竭")
)
foreach($b in $bt){
    SetC $ws3 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $b[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $b[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# Rules
$r+=2
SetC $ws3 $r 1 "进阶判断规则——你自己做决策" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 3 9; $r++

$rules=@(
    @("推日5x5轻松完成 RPE实际只有7","下次+2.5-5kg——RPE优先于百分比"),
    @("推日5x5第3组就变形","减重5-10%——今天的你不如上周的你 正常 不退步"),
    @("腿日间歇时间觉得太长","缩短到75s或60s——代谢压力不够就加密度"),
    @("上肢日某个动作连续2周没进步","换变式——同肌群不同角度重新刺激"),
    @("减载周感觉'太轻松'","这是对的——减载周就该轻松。不轻松说明前三周没练到位"),
    @("连续2周恢复不过来","延长减载至1周——你是教练 知道过度训练的代价")
)
foreach($r3 in $rules){
    SetC $ws3 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 3 $r3[1] $false 10 0 0 $true; MC $ws3 $r 3 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=26; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "热身  主动恢复  减载周  有氧" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

$r=3
SetC $ws4 $r 1 "训练前热身（10分钟 你是教练 自己调）" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "时间" $true 10 $white $headerBlue $false; $r++

$warmup=@(
    @("一般热身","划船机/单车/跳绳——轻度有氧升温","3-5min"),
    @("动态拉伸","猫式+最伟大伸展+摆腿+肩环绕","2-3min"),
    @("神经激活","跳箱低高度/药球砸地/弹力带侧走","1-2min"),
    @("专项热身","主项动作空杆/轻重量 1-2组x5-8次 递增到正式重量","2-3min")
)
foreach($wu in $warmup){
    SetC $ws4 $r 1 $wu[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $wu[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $wu[2] $false 10 0 0 $false; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=24; SB $ws4 $r 1 9; $r++
}

# Active recovery
$r+=2
SetC $ws4 $r 1 "休息日主动恢复（周三/六/日——选1-2天做）" $true 14 $white $softBlue $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$activeRecovery=@(
    "你不是普通会员——你每天在健身房站着教课 那不是休息。真正的休息日是：",
    "  泡沫轴全面放松：股四头+腘绳+臀+上背+胸+背阔——10-15min",
    "  低强度有氧：散步/游泳/休闲骑——20-30min 心率<120bpm",
    "  拉伸：全身主要肌群各30-45s",
    "  完全不练也比'轻重量练一下'好——休息日的目标是恢复 不是'顺便练练'"
)
foreach($ar in $activeRecovery){
    SetC $ws4 $r 1 $ar $false 10 0 0 $true; MC $ws4 $r 1 9
    $ws4.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# Deload
$r+=2
SetC $ws4 $r 1 "第4周减载——不是'做轻一点' 是'做不同的动作'" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "减载周替换" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 9; $r++

$deload=@(
    @("杠铃卧推","哑铃卧推(轻重量)或俯卧撑","2-3组x8-10 RPE4-5"),
    @("杠铃深蹲","高脚杯深蹲(轻哑铃)","2组x10-12 RPE4-5"),
    @("杠铃划船","哑铃单臂划船(轻重量)","2组x10/侧 RPE4-5"),
    @("硬拉/RDL","早安式体前屈(自重/轻量)","2组x10-12 RPE4-5"),
    @("引体向上","弹力带辅助引体或高位下拉轻量","2组x8-10 RPE4-5"),
    @("减载周有氧","低强度休闲骑/散步","20-30min 心率<120"),
    @("减载周心态","做完应该感觉'没练够'——这就对了。身体在偷偷变强。","")
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
# SHEET 5: 饮食+调节
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与调节"

SetC $ws5 1 1 "增肌饮食  自主调节框架" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

$r=3
SetC $ws5 $r 1 "增肌营养——你是教练 给你方向不用教程" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "目标" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "方向" $true 10 $white $headerBlue $false; $r++

$nutri=@(
    @("总热量","2700-2900 kcal","比维持高200-300——干净增肌 不做脏增肌"),
    @("蛋白质","160-180g (2.0-2.2g/kg)","每餐35-40g 练后一勺蛋白粉+正餐"),
    @("碳水","320-360g","训练前后集中摄入。你教一天课 碳水是能量来源——不吃碳水=训练表现打折"),
    @("脂肪","60-70g","不刻意控——正常饮食的脂肪就够了"),
    @("饮水","3-4L/天","上课说一天话——水分流失比会员大得多")
)
foreach($n in $nutri){
    SetC $ws5 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $n[2] $false 10 0 0 $true; MC $ws5 $r 4 9
    $ws5.Range("A$($r):I$($r)").RowHeight=26; SB $ws5 $r 1 9; $r++
}

# Training window
$r++
SetC $ws5 $r 1 "训练窗口——你是教练 应该已经在做了" $true 12 $headerBlue 0 $true; MC $ws5 $r 1 9; $r++

SetC $ws5 $r 1 "练前1-2h：碳水40-50g+蛋白20-25g。练后1h内：碳水50-60g+蛋白30-35g。你在健身房上班——练后可以直接去食堂或叫外卖 比普通会员方便得多。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

# Self-regulation framework
$r+=2
SetC $ws5 $r 1 "自主调节框架——这套你可以带走" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$selfReg=@(
    "你是教练——下面这套逻辑不只适用于你自己 也可以迁移到给会员做计划。",
    "",
    "1. 定分化方式：每周能练几天→推拉腿/上下肢/全身。你4天=推拉腿+上肢弱点日。",
    "2. 定DUP主题：每天不同刺激——大重量/容量/代谢/弱点。不要让身体适应同一种节奏。",
    "3. 定周期节奏：3周渐进+1周减载=1小周期。每3个月=1周全休。",
    "4. 定RPE范围：不只看重量——看当天的感觉。RPE优先于百分比。",
    "5. 每4周复盘：看力量数据+体重趋势+照片。判断该继续增肌还是短暂维持后再冲。",
    "",
    "这套框架——你在西宁健身房可以用在任何一个会员身上。它不依赖器械 不依赖会员水平——只依赖于你对自己身体的了解。"
)
foreach($sr in $selfReg){
    SetC $ws5 $r 1 $sr $false 10 0 0 $true; MC $ws5 $r 1 9
    $ws5.Range("A$($r):I$($r)").RowHeight=20
    if($sr -eq ""){$ws5.Range("A$($r):I$($r)").RowHeight=8}
    $r++
}

# Common issues
$r++
SetC $ws5 $r 1 "常见问题和判断" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$issues=@(
    @("今天上课站了一天 累 不想练","降级：热身+主项做3组轻重量+拉伸。保持训练习惯 不追求今天的强度。"),
    @("体重2周没涨","加碳水——每天多一碗饭的量。增肌慢是正常的 不涨说明热量缺口还存在。"),
    @("某个动作连续3周卡在同样重量","换变式——同肌群不同动作。或者减载1周再回来。"),
    @("推日卧推5x5最后一组只能做3次","正常——5x5本来就不是每组都能做满。减重或接受'最后一组只能做3次'。"),
    @("连续2天失眠 状态极差","不练或只做主动恢复。睡眠优先于训练——你是教练 知道这个。")
)
SetC $ws5 $r 1 "遇到什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "怎么做" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++
foreach($is in $issues){
    SetC $ws5 $r 1 $is[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 3 $is[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=34; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用+复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  4周复盘  下一步" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

# Simplified
$r=3
SetC $ws6 $r 1 "如果这周只能2-3练——精简版" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "精简原则：保留核心复合动作 压缩辅助动作。2练=推+腿/拉+上肢；3练=推/拉+腿/上肢。" $false 10 $headerBlue $lightGray $true
MC $ws6 $r 1 9; $r++

SetC $ws6 $r 1 "选项" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "训练日1" $true 10 $white $headerBlue $false; MC $ws6 $r 2 4
SetC $ws6 $r 5 "训练日2" $true 10 $white $headerBlue $false; MC $ws6 $r 5 7
SetC $ws6 $r 8 "训练日3(3练)" $true 10 $white $headerBlue $false; MC $ws6 $r 8 9; $r++

$simpleOpts=@(
    @("2练","推日(精简：卧推5x5+推举+三头)  腿日(精简：深蹲+RDL+提踵)","拉日(精简：划船+引体+弯举)  上肢日(精简：哑铃卧推+下拉+侧平举)",""),
    @("3练","推+腿(精简：卧推5x5+深蹲+推举+臂屈伸+提踵)","拉+上肢(精简：划船+引体+哑铃卧推+下拉+弯举+侧平举)","全身综合(精简：深蹲+卧推+划船+推举+面拉+核心)")
)
foreach($so in $simpleOpts){
    SetC $ws6 $r 1 $so[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $so[1] $false 10 0 0 $true; MC $ws6 $r 2 4
    SetC $ws6 $r 5 $so[2] $false 10 0 0 $true; MC $ws6 $r 5 7
    SetC $ws6 $r 8 $so[3] $false 10 0 0 $true; MC $ws6 $r 8 9
    $ws6.Range("A$($r):I$($r)").RowHeight=48; SB $ws6 $r 1 9; $r++
}

# Review
$r+=2
SetC $ws6 $r 1 "4周后复盘——用数据说话" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws6 $r 2 6
SetC $ws6 $r 7 "下周期调整" $true 10 $white $headerBlue $false; MC $ws6 $r 7 9; $r++

$review=@(
    @("卧推5x5重量","第1周vs第3周——涨了多少？","涨了=继续加重。没涨=换变式或增加容量。"),
    @("深蹲8-10RM","第1周vs第3周——下肢力量变化","涨了=加重。没涨=检查恢复和睡眠。"),
    @("体重趋势","4周涨了多少？目标是0.5-1kg/月净肌肉","涨超1kg/周=体脂也在涨 控热量。没涨=加碳水。"),
    @("照片","第1天vs第28天","镜子里的变化比体重秤真实"),
    @("DUP感受","哪个主题日进步最明显？推日/拉日/腿日/上肢日？","进步明显的=保持。停滞的=下周期调整。"),
    @("下周期方向","继续增肌还是短暂维持？","如果4周涨了1-2kg且力量涨了=继续。体脂明显上升=短暂维持2周再冲。")
)
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 6
    SetC $ws6 $r 7 $rv[2] $false 10 0 0 $true; MC $ws6 $r 7 9
    $ws6.Range("A$($r):I$($r)").RowHeight=30; SB $ws6 $r 1 9; $r++
}

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\吉_增肌方案_V1.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
