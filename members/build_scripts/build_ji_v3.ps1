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

SetC $ws0 1 1 "吉  增肌方案 V3" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  182cm / 82kg  |  增肌  |  V3 = 教科书级 完整框架" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  专业术语速查（方案里用到的词都在这里）" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

SetC $ws0 $row 1 "术语" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "全称" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "白话解释" $true 10 $white $headerBlue $false
MC $ws0 $row 3 5; $row++

$terms=@(
    @("DUP","每日波动周期","每天练的不一样。周一重、周二多、周四快、周五细——不让身体适应同一种节奏。"),
    @("AMRAP","尽可能多做一次","最后一组不限次数 做到动作快变形了就停。用来测试今天真实的极限。"),
    @("RPE","自感用力程度","做完一组凭感觉打分。1=躺着 10=拼老命。8-9=做完还剩1-2次力竭。"),
    @("超级组","Superset","两个动作连着做 中间不休息。A做完直接做B。压缩时间 增加代谢压力。"),
    @("预疲劳法","Pre-exhaustion","先孤立充血 再复合深度刺激。突破'做完了但目标肌群还没感觉'的瓶颈。"),
    @("减载周","Deload Week","每4周1周轻量 换动作。身体消化前三周刺激。超量恢复发生在休息时。"),
    @("皮质醇","Cortisol","压力激素。训练太久分泌它——分解肌肉。腿日45分钟限时就为控制它。")
)
foreach($t in $terms){
    SetC $ws0 $row 1 $t[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $t[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $t[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

# DUP 4-day
$row+=2
SetC $ws0 $row 1 "  DUP 4天分化——V3版" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "V3核心方法" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "RPE" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$dup=@(
    @("周一","推日(大重量)","4x5+1AMRAP 渐进决策树","8-9","前4组标准 最后一组探底"),
    @("周二","拉日(宽/厚轮换)","动作分类说明：宽度优先vs厚度优先","7-8","背的宽度+厚度双维度"),
    @("周四","腿日(代谢45min)","深蹲+超级组 45分钟限时=黄金窗口","7-8","短间歇 45min内完成"),
    @("周五","上肢日(预疲劳)","预疲劳法：孤立充血  复合深度","7-8","先充血再突破瓶颈")
)
foreach($d in $dup){
    SetC $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $d[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $d[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $d[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

# Week schedule
$row+=2
SetC $ws0 $row 1 "  每周安排+精简版" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$weekly=@(
    @("周一=推日(大重量) / 周二=拉日(宽厚轮换) / 周四=腿日(45min限时) / 周五=上肢日(预疲劳)"),
    @("周三/六/日=休息/主动恢复(泡沫轴+灵活性+低强度有氧)"),
    @("只能2练：推+腿 / 拉+上肢。只能3练：推+腿 / 拉+上肢 / 全身综合。")
)
foreach($w in $weekly){
    SetC $ws0 $row 1 $w[0] $false 10 0 0 $true; MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; SB $ws0 $row 1 5; $row++
}

# Annual framework
$row+=2
SetC $ws0 $row 1 "  年度周期框架——我现在在哪 接下来去哪" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "季度" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "目标" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "训练特点" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "收尾" $true 10 $white $headerBlue $false; $row++

$annual=@(
    @("Q1(1-3月)","积累期","建基准 提容量","DUP 4天 3周渐进+1周减载","第3月末=1周全休"),
    @("Q2(4-6月)","强化期","提力量 增肌肉","重量递增 RPE8-9","第6月末=1周全休"),
    @("Q3(7-9月)","峰值期","冲击极限 测1RM","强度优先 RPE9-10(偶尔)","第9月末=1周全休"),
    @("Q4(10-12月)","维持/过渡","维持力量 降压力","容量为主 RPE6-7","年底=2周全休(彻底重置)")
)
foreach($a in $annual){
    SetC $ws0 $row 1 $a[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $a[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $a[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $a[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $a[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row++
SetC $ws0 $row 1 "你现在在哪？如果刚开始——在Q1积累期。先跑完一个4周周期 把基准建起来。" $false 10 $softBlue 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+周期框架+RPE指南+全年框架"," Sheet 方案概览"),
    @("推日 拉日(宽/厚) 腿日(45min) 上肢日(预疲劳)"," Sheet 训练计划"),
    @("渐进决策树+器械突破+进阶判断"," Sheet 渐进决策树"),
    @("热身+主动恢复+减载(换动作)+季度全休"," Sheet 恢复策略"),
    @("增肌营养+训练窗口+自主调节+心理策略"," Sheet 饮食与调节"),
    @("精简版+4周复盘+年度周期"," Sheet 备用与复盘")
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

SetC $ws 1 1 "吉  增肌方案  概览 V3" $true 18 $headerBlue 0 $true
MC $ws 1 1 9; $ws.Range("A1:I1").RowHeight=36

SetC $ws 2 1 "男  25-30岁  182cm/82kg  体脂20-25%  健身教练  增肌  商业健身房  DUP+AMRAP+预疲劳+宽厚轮换+渐进决策树" $false 11 $gold $darkBg $true
MC $ws 2 1 9; $ws.Range("A2:I2").Font.Color=$white; $ws.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览 V3" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 9

$ov=@(
    @("训练频率","4次/周 推/拉/腿/上肢 DUP波动","2-3练精简版 附渐进决策树"),
    @("核心方法","4x5+1AMRAP 宽/厚轮换 45min限时 预疲劳法 换动作减载 季度全休","6大方法从'优秀'到'教科书级'"),
    @("周期结构","3周渐进+1周减载(换动作) 3小周期+1周全休=1中周期","年度框架：积累  强化  峰值  维持"),
    @("新增V3","专业术语速查 动作分类逻辑 渐进决策树 主动恢复 年度周期 心理策略","教科书级——完整的训练生态系统"),
    @("方案特点","教科书级完整框架 可迁移至任何会员")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 6
    SetC $ws $r 7 $o[2] $false 10 0 0 $true; MC $ws $r 7 9
    SB $ws $r 1 9; $r++
}

# V3 upgrades
$r++
SetC $ws $r 1 "V3 新增模块——从'好计划'到'教科书级'" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$v3New=@(
    @("专业术语速查","DUP/AMRAP/RPE/超级组/预疲劳/减载/皮质醇 七个术语的现场解释","封面+本Sheet"),
    @("动作分类逻辑","每个训练日拆成 主项/辅助/收尾——解释'为什么选这个动作' 以后能自己替换","训练计划Sheet"),
    @("渐进决策树","推日/拉日/腿日/上肢日——各自'何时加重/维持/退阶/换动作' 训练像做实验","渐进决策树Sheet"),
    @("主动恢复清单","泡沫轴/髋关节/肩关节/低强度有氧 四套具体动作 休息日不只是躺着","恢复策略Sheet"),
    @("年度周期框架","Q1积累 Q2强化 Q3峰值 Q4维持——只有战术没有战略=每月在重复","封面+本Sheet"),
    @("心理策略","教课太累不想练/连续没进步/觉得自己不如会员——5种状态的心理应对","饮食与调节Sheet")
)
SetC $ws $r 1 "模块" $true 10 $white $headerBlue $false
SetC $ws $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws $r 2 4
SetC $ws $r 5 "在哪" $true 10 $white $headerBlue $false; MC $ws $r 5 9; $r++
foreach($vn in $v3New){
    SetC $ws $r 1 $vn[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $vn[1] $false 10 0 0 $true; MC $ws $r 2 4
    SetC $ws $r 5 $vn[2] $false 10 0 0 $true; MC $ws $r 5 9
    $ws.Range("A$($r):I$($r)").RowHeight=28; SB $ws $r 1 9; $r++
}

# RPE
$r+=2
SetC $ws $r 1 "RPE参考" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "RPE" $true 10 $white $headerBlue $false
SetC $ws $r 2 "感觉" $true 10 $white $headerBlue $false
SetC $ws $r 4 "剩余次数" $true 10 $white $headerBlue $false
SetC $ws $r 6 "用在哪天" $true 10 $white $headerBlue $false; $r++

$rpeRef=@(
    @("6-7","中等——做完还能做4-6次","拉日/上肢日——泵感和代谢积累"),
    @("7-8","中等偏强——做完还能做3-4次","腿日——增肌的理想区间"),
    @("8-9","很重——做完还能做1-2次","推日AMRAP——最后一组探到RPE9"),
    @("10","极限力竭","本方案几乎不用。AMRAP最后一组最后1次可能触达——够了。")
)
foreach($rp in $rpeRef){
    SetC $ws $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    SetC $ws $r 2 $rp[1] $false 10 0 0 $false
    SetC $ws $r 4 $rp[2] $false 10 0 0 $false
    SetC $ws $r 6 $rp[3] $false 10 0 0 $true; MC $ws $r 6 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

# Annual
$r+=2
SetC $ws $r 1 "年度周期框架" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "季度" $true 10 $white $headerBlue $false
SetC $ws $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws $r 3 "目标" $true 10 $white $headerBlue $false
SetC $ws $r 5 "训练特点" $true 10 $white $headerBlue $false
SetC $ws $r 7 "收尾" $true 10 $white $headerBlue $false; $r++

$annualDetail=@(
    @("Q1(1-3月)","积累期","建基准 提容量","DUP 4天 3周渐进+1周减载","第3月末=1周全休"),
    @("Q2(4-6月)","强化期","提力量 增肌肉","重量递增 RPE 8-9","第6月末=1周全休"),
    @("Q3(7-9月)","峰值期","冲击极限 测1RM","强度优先 RPE 9-10(偶尔)","第9月末=1周全休"),
    @("Q4(10-12月)","维持/过渡","维持力量 降压力","容量为主 RPE 6-7","年底=2周全休(彻底重置)")
)
foreach($ad in $annualDetail){
    SetC $ws $r 1 $ad[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $ad[1] $true 10 $white $darkBg $false
    SetC $ws $r 3 $ad[2] $false 10 0 0 $false
    SetC $ws $r 5 $ad[3] $false 10 0 0 $false; MC $ws $r 5 6
    SetC $ws $r 7 $ad[4] $false 10 0 0 $true; MC $ws $r 7 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "没有年度视角的周期设计 = 只有战术没有战略。你需要知道自己在哪条线上——现在是Q1积累期 先跑完第一个4周 把基准建起来。" $false 10 $softBlue 0 $true
MC $ws $r 1 9; $ws.Range("A$($r):I$($r)").RowHeight=24; $r++

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$pr=@(
    @("1","动作质量 > 组数","5x5第五组变形=得不偿失。4x5+1AMRAP让你在质量不降的前提下探底。健美选手的底层逻辑——质量永远优先。"),
    @("2","RPE优先于百分比","80% 1RM是理论值。当天教课多久 睡没睡好——影响实际表现。RPE让你和身体对话。"),
    @("3","恢复=训练的一部分","教一天课+自己练=双倍消耗。季度全休不是偷懒 是职业保护。身体在恢复时变强。"),
    @("4","理解角色 才能替换","知道每个动作在框架中的角色（主项/辅助/收尾）——以后你自己替换动作不破坏训练逻辑。")
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
# SHEET 2: 训练计划 V3 (含动作分类)
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "DUP 4天分化 V3——推/拉(宽厚轮换)/腿(45min)/上肢(预疲劳)" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "每个动作标注了'角色'——理解角色后才能自己替换。V3新增：动作分类逻辑 渐进决策树见Sheet3。周一=推日 周二=拉日 周四=腿日 周五=上肢日。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 推日 V3
$r=4
$actsPush=@(
    @("杠铃平板卧推","主项(复合 神经驱动最高)","4组+1AMRAP","5次","前4组80%1RM 第5组AMRAP RPE8-9","120-150s","AMRAP=每周进步标尺","【呼吸】瓦式 下放吸气屏住 推起中途呼气。【要点】起桥 肩胛收紧 脚踩实。前4组RPE8-9 第5组做到技术变形前停。记录AMRAP次数——这是每周进步的标尺。"),
    @("上斜哑铃卧推","辅助1(上胸变式)","4组","8-10次","RPE7-8 上胸+肩前束","90s","补充主项的盲区","【要点】凳角30-45度 下放至胸锁骨水平。和平板不同角度——上胸和肩前束同时刺激。"),
    @("坐姿哑铃推举","辅助2(垂直推 补肩部)","4组","8-10次","RPE7-8 肩前束+中束","75s","卧推主练胸 肩前束不够——推肩补上","【要点】下背贴靠背 不反弓。推日肩部主项——三角肌前束+中束。"),
    @("双杠臂屈伸","辅助3(三头孤立)","3组","8-12次","自重或负重 RPE7-8","60s","卧推后三头已疲劳——此时做最有效","【要点】身体稍前倾=胸+三头。身体竖直=三头为主。"),
    @("绳索飞鸟","收尾1(代谢补充)","3组","12-15次","轻重量 追求泵感","45s","主项后继续刺激胸肌代谢","【要点】手臂微屈 胸肌驱动。胸肌在疲劳状态下继续做功——泵感收尾。"),
    @("三头绳索下压","收尾2(代谢补充)","3组","12-15次","轻-中重量","45s","AMRAP后三头高疲劳——下压效率最高","【要点】大臂固定 压到底微外旋。")
)
$r=Write-TrainingBlock $ws2 $r "周一  推日（大重量 4x5+1AMRAP）" "主项(杠铃卧推4x5+1AMRAP)+辅助(上斜哑铃+推肩+臂屈伸)+收尾(飞鸟+三头)" "45-55" $actsPush

# 拉日 宽度优先
$r++
SetC $ws2 $r 1 "拉日——V3动作分类：宽度优先版(A) 第1/3周使用" $true 14 $white $softBlue $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$actsPullW=@(
    @("宽握高位下拉","主项(垂直拉 宽度)","4组","10-12次","RPE7-8 宽度=视觉冲击力","75s","第1/3周=宽度优先 背阔肌外沿","【要点】宽握 杆拉至锁骨 顶峰背阔肌向外撑开。这是宽度日 不是重量日——感受背阔肌向外展。"),
    @("直臂下压","辅助1(拉伸角度)","3组","15次","轻重量 背阔肌全程拉伸","45s","和下拉互补——头顶下压至大腿","【要点】手臂几乎伸直 从头顶下压至大腿。背阔肌的全程拉伸——宽度日的补充。"),
    @("窄握坐姿划船","辅助2(水平拉 厚度补充)","3组","12-15次","RPE7 窄握=中背部","60s","宽度日也需厚度——轻量补充","【要点】V柄拉至腹部 顶峰肩胛后缩1秒。"),
    @("罗马尼亚硬拉","下肢(后链 固定动作)","4组","8-10次","RPE7-8 腘绳肌+臀","90s","拉日固定——不随宽/厚轮换","【要点】膝盖微屈 全程直背。和腿日深蹲互补。"),
    @("面拉","收尾1(肩袖保护)","3组","15次","轻重量","45s","肩后束+上背——拉日肩部补充","【要点】拉向额头 肘外展。"),
    @("杠铃弯举","收尾2(二头充血)","3组","10-12次","RPE7","45s","划船下拉后二头已充血——弯举收尾","【要点】大臂固定 顶峰外旋。")
)
$r=Write-TrainingBlock $ws2 $r "拉日A：宽度优先（Width Priority）" "主项(宽握下拉)+辅助(直臂下压+窄握划船)+下肢(RDL)+收尾(面拉+弯举)" "45-50" $actsPullW

# 厚度优先
SetC $ws2 $r 1 "拉日——V3动作分类：厚度优先版(B) 第2周使用" $true 14 $white $softBlue $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$actsPullT=@(
    @("杠铃俯身划船","主项(水平拉 厚度)","4组","8-10次","RPE7-8 厚度=立体感 可更重","90s","第2周=厚度优先 背阔肌厚度+中背部","【要点】俯身45-60度 核心收紧 杠铃贴大腿前侧拉至肚脐。重量可以比宽度日重。"),
    @("T杠/单臂哑铃划船","辅助1(不同角度划船)","3组","10-12次","RPE7-8 中背部厚度","75s","T杠=中性握 单臂=左右独立","【要点】T杠=对握 中背部厚度。单臂=左右独立 纠正不平衡。"),
    @("宽握高位下拉(轻量)","辅助2(宽度维持 轻量)","3组","15次","轻重量 RPE6","60s","厚度日保持宽度刺激 不追求力竭","【要点】和宽度日同动作 轻重量。保持背阔肌外沿刺激。"),
    @("罗马尼亚硬拉","下肢(后链 固定动作)","4组","8-10次","RPE7-8","90s","拉日固定 不随宽/厚轮换","【要点】同宽度日。"),
    @("面拉","收尾1(肩袖保护)","3组","15次","轻重量","45s","—","【要点】同宽度日。"),
    @("哑铃弯举","收尾2(二头充血)","3组","10-12次","RPE7","45s","—","【要点】同宽度日。")
)
$r=Write-TrainingBlock $ws2 $r "拉日B：厚度优先（Thickness Priority）" "主项(杠铃划船)+辅助(T杠/单臂+宽握下拉轻量)+下肢(RDL)+收尾(面拉+弯举)" "45-50" $actsPullT

# 腿日 V3
$actsLegs=@(
    @("杠铃深蹲","主项(复合 下肢核心)","4组","8-10次","RPE7-8 严格60s休息","60s","短间歇=代谢压力 非力量举长间歇","【呼吸】瓦式 下蹲吸气屏住 站起呼气。【要点】深度至少平行地面。V3核心=45分钟内完成。超时=皮质醇分解肌肉。"),
    @("腿举+腿弯举超级组","超级组(代谢密度翻倍)","A3x15-20 B3x15-20","交替不休息","A RPE7-8 B RPE7","组内不休息 组间60s","A做完直接B——股四+腘绳同时充血","【要点】A腿举做完直接做B腿弯举 不休息。A+B=1轮 3轮 轮间休息60s。如果45分钟不够 优先减轮数——宁可少做 不超时。"),
    @("站姿提踵","收尾1(小腿)","3组","15-20次","RPE7 膝盖微屈","30s","腿日疲劳状态下完成","【要点】顶峰停1秒 慢放。"),
    @("悬垂举腿","收尾2(核心)","3组","12-15次","自重","45s","代谢疲劳状态下稳定核心","【要点】下腹带动骨盆上卷 不靠惯性。"),
    @("腿日总时长","目标45分钟内","超时=皮质醇开始分解肌肉","—","—","45min=代谢压力黄金窗口","如果超时：减少超级组轮数 而非牺牲动作质量。")
)
$r=Write-TrainingBlock $ws2 $r "周四  腿日（代谢 45min限时）" "主项(深蹲)+超级组(腿举+弯举)+收尾(提踵+举腿)  45分钟内完成" "45" $actsLegs

# 上肢日 V3
$actsUpper=@(
    @("侧平举+三头下压超级组","预疲劳超级组(先做!)","A3x15 B3x15","交替不休息","轻-中重量","组内不休息 组间60s","V3预疲劳法：先充血  再深度刺激","【要点】A做完直接做B=1轮。3轮。目标：让侧肩和三头在复合动作前充血——突破瓶颈。"),
    @("窄距卧推/地板卧推","主项(三头主导复合)","4组","8-10次","RPE7-8 预疲劳后重量比平时轻=正常","90s","预疲劳后目标肌群已充血→深度刺激","【要点】握距比肩窄 肘贴身体。地板卧推=离地30cm启动 三头全程张力。重量不是指标 泵感才是。"),
    @("高位下拉(宽握)","辅助(垂直拉 保持背)","4组","10-12次","RPE7-8","75s","上肢日保持背阔肌刺激","【要点】和拉日同动作。"),
    @("面拉(绳索)","收尾1(肩袖保护)","3组","15次","轻重量","45s","肩后束+肩袖健康","【要点】拉向额头 肘外展。"),
    @("农夫行走","收尾2(功能性收尾)","3组","30-40m","大重量哑铃 单边30-40kg","60s","握力+斜方+核心=功能性","【要点】躯干正直不侧倾 核心全程收紧。")
)
$r=Write-TrainingBlock $ws2 $r "周五  上肢日（预疲劳法 Pre-Fatigue）" "预疲劳超级组  主项(窄距卧推)+辅助(下拉)+收尾(面拉+农夫行走)" "40-45" $actsUpper

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18; $ws2.Range("C:C").ColumnWidth=16
$ws2.Range("D:D").ColumnWidth=8; $ws2.Range("E:E").ColumnWidth=10; $ws2.Range("F:F").ColumnWidth=16
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=22; $ws2.Range("I:I").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进决策树 (V3 重头戏)
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进决策树"

SetC $ws3 1 1 "渐进决策树——每次训练前看  训练像做实验" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34
SetC $ws3 2 1 "没有决策标准的渐进 = 凭感觉训练。有了决策树 每次训练都是一个数据点 4周后复盘能看到清晰的进步轨迹。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 9

$r=4
# 推日决策树
SetC $ws3 $r 1 "周一推日（强度日）——渐进决策树" $true 14 $white $accentRed $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$accentRed
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$pushTree=@(
    @("前4组5次都完成 且RPE  8","  变强了","下次加重2.5-5kg"),
    @("前4组5次都完成 但RPE = 9","  最佳训练区间","维持当前重量——这就是最好的刺激"),
    @("前4组中有1组没完成5次","  今天状态或重量不对","退回上周重量 下周再试"),
    @("连续2周卡在同一重量","  同一刺激已适应","换变式：地板卧推/暂停卧推/窄距卧推"),
    @("AMRAP次数比上周+2次以上","  真的变强了","下次加重2.5-5kg"),
    @("AMRAP次数没变或下降","  累积疲劳或状态差","维持重量——不退步 下周再看")
)
foreach($pt in $pushTree){
    SetC $ws3 $r 1 $pt[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $pt[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $pt[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# 拉日决策树
$r+=2
SetC $ws3 $r 1 "周二拉日（容量日）——渐进决策树" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$pullTree=@(
    @("所有组完成目标次数 且RPE  7","  可以加量了","下次+1组或缩短休息15秒"),
    @("所有组完成 但RPE = 8","  最佳刺激区间","维持——刚好在理想强度"),
    @("有1-2组没完成目标次数","  容量或恢复不够","减重或增加休息时间"),
    @("宽度vs厚度——哪个进步更快？","  用数据说话","进步快的保持 慢的下周期多轮一次"),
    @("宽度日高位下拉RPE  6","  太轻松","下周加重1片——但保持'感受背阔肌外展'")
)
foreach($pl in $pullTree){
    SetC $ws3 $r 1 $pl[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $pl[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $pl[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# 腿日决策树
$r+=2
SetC $ws3 $r 1 "周四腿日（代谢日）——渐进决策树" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$legsTree=@(
    @("45分钟内完成全部","  代谢压力控制完美","下周缩短休息5-10秒 提高密度"),
    @("45分钟没完成","  容量太大或休息太长","减少超级组轮数——宁可做少 不超时"),
    @("深蹲RPE  7 轻松完成","  可以加重了","下周+2.5-5kg"),
    @("超级组做到第2轮就力竭","  代谢压力的目的达到了","维持——不追求更多轮数"),
    @("腿日结束后极度疲劳 想吐","  皮质醇可能超标了","下周减1轮超级组 或延长休息")
)
foreach($lg in $legsTree){
    SetC $ws3 $r 1 $lg[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $lg[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $lg[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# 上肢日决策树
$r+=2
SetC $ws3 $r 1 "周五上肢日（弱点日）——渐进决策树" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "情况" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$upperTree=@(
    @("预疲劳后窄距卧推重量比预期轻","  预疲劳的目的就是这个","正常。重量不是指标 泵感才是。"),
    @("预疲劳超级组做完 侧肩三头没泵感","  充血不够","加重或缩短休息——必须感受到充血"),
    @("窄距卧推连续2周没进步","  需要换刺激","换变式：JM Press/仰卧臂屈伸/双杠"),
    @("窄距卧推RPE  8 轻松完成","  可以加重","下次+1-2kg——但预疲劳后重量的参考价值有限")
)
foreach($up in $upperTree){
    SetC $ws3 $r 1 $up[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $up[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $up[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# Breakthrough
$r+=2
SetC $ws3 $r 1 "器械上限突破策略" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "策略" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "操作" $true 10 $white $headerBlue $false; MC $ws3 $r 2 4
SetC $ws3 $r 5 "适用" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$bt=@(
    @("慢离心","3-4s下放+爆发推起","所有日子——TUT增加"),
    @("1.5次法","下放->半起->全下->全起=1次","推日/腿日——双倍刺激"),
    @("缩短休息","90s->60s->45s","腿日代谢——核心策略"),
    @("预疲劳","孤立力竭->复合动作","上肢日——本方案已用")
)
foreach($b in $bt){
    SetC $ws3 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $b[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $b[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# Universal rules
$r+=2
SetC $ws3 $r 1 "通用进阶判断规则" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 3 9; $r++

$urules=@(
    @("连续2周睡不好+训练状态差","提前进入减载周——累积疲劳的信号"),
    @("连续4周某个动作没进步","换变式——同肌群不同角度重新刺激"),
    @("减载周感觉'太轻松'","这是对的——减载周不轻松=前三周没练到位"),
    @("季度全休周到期 身体渴望训练","新周期从全休前第3周70-75%起步"),
    @("连续2周恢复不过来","延长减载至1-2周——你是教练 知道过度训练的代价")
)
foreach($ur in $urules){
    SetC $ws3 $r 1 $ur[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $ur[1] $false 10 0 0 $true; MC $ws3 $r 3 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=30; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略 V3 (含主动恢复清单)
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "热身  主动恢复清单  减载(换动作)  季度全休" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

# Warmup
$r=3
SetC $ws4 $r 1 "训练前热身（10分钟）" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "时间" $true 10 $white $headerBlue $false; $r++

$wu=@(
    @("一般热身","划船机/单车/跳绳——轻度有氧升温","3-5min"),
    @("动态拉伸","猫式+最伟大伸展+摆腿+肩环绕","2-3min"),
    @("神经激活","跳箱(低)/药球砸地/弹力带侧走","1-2min"),
    @("专项热身","主项空杆/轻重量 1-2组x5-8次 递增至正式重量","2-3min")
)
foreach($w in $wu){
    SetC $ws4 $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $w[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $w[2] $false 10 0 0 $false; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=24; SB $ws4 $r 1 9; $r++
}

# V3: 主动恢复清单
$r+=2
SetC $ws4 $r 1 "休息日主动恢复清单——V3新增 必修课" $true 14 $white $accentGreen $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "你教一天课 身体一直在消耗。主动恢复不是可选项——是必修课。周三/六/日选1-2项做 15-20分钟。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=22; $r++

SetC $ws4 $r 1 "恢复项目" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "具体动作" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "时间" $true 10 $white $headerBlue $false; MC $ws4 $r 6 9; $r++

$arList=@(
    @("1. 泡沫轴全身放松","顺序：小腿->股四头肌->腘绳肌->臀->背->胸`n每个部位滚30秒 找到酸痛点停留15秒深呼吸","10分钟"),
    @("2. 髋关节灵活性","世界最伟大拉伸 每侧5次`n90-90髋拉伸 每侧30秒`n青蛙趴 30秒","10分钟"),
    @("3. 肩关节保养","YTWL 每个字母10次`n面拉 3x15(轻重量)`n弹力带肩外旋 每侧15次","10分钟"),
    @("4. 低强度有氧","散步/游泳/休闲骑行`n心率控制在120-130bpm——不是训练 是恢复","20分钟"),
    @("5. 完全不练也比'轻重量练一下'好","休息日的目标是恢复 不是'顺便练练'","—")
)
foreach($ar in $arList){
    SetC $ws4 $r 1 $ar[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $ar[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $ar[2] $false 10 0 0 $false; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=40; SB $ws4 $r 1 9; $r++
}

# Deload V3
$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 不是做轻一点" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "同样动作做轻量=关节应力路径相同 不是真休息。换动作=给关节不同应力分布 真正的休息。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 9; $r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "减载周换成" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 9; $r++

$deload=@(
    @("杠铃卧推","俯卧撑(3秒下放+爆发推起)","2-3组x8-10 RPE4-5"),
    @("杠铃划船","弹力带划船/TRX划船","2组x10-12 RPE4-5"),
    @("杠铃深蹲","高脚杯深蹲/箱式深蹲(轻量)","2组x10-12 RPE4-5"),
    @("硬拉/RDL","罗马尼亚硬拉(轻量 强调离心)","2组x10-12 RPE4-5"),
    @("高位下拉","弹力带高位下拉(单臂)","2组x10/侧 RPE4-5"),
    @("窄距卧推","弹力带推胸","2组x12-15 RPE4-5"),
    @("有氧","低强度休闲骑/散步","20-30min <120bpm"),
    @("心态","做完应该感觉'没练够'——这就对了","")
)
foreach($dl in $deload){
    SetC $ws4 $r 1 $dl[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $dl[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $dl[2] $false 10 0 0 $true; MC $ws4 $r 4 9
    $ws4.Range("A$($r):I$($r)").RowHeight=26; SB $ws4 $r 1 9; $r++
}

# Quarterly
$r+=2
SetC $ws4 $r 1 "季度全休周——每3个月1次 职业保护" $true 14 $white $accentOrange $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

$quarter=@(
    "即使有减载周 连续3个月DUP训练仍累积系统性疲劳。教课+训练=双倍恢复需求。",
    "全休周：散步/游泳/拉伸/泡沫轴。不碰器械。完全休息=系统重置。",
    "不是偷懒——是职业保护。长期坚持训练的人 是那些懂得何时休息的人。",
    "全休周结束后 身体和神经系统都'渴望'训练——新周期最佳起点。"
)
foreach($q in $quarter){
    SetC $ws4 $r 1 $q $false 10 0 0 $true; MC $ws4 $r 1 9
    $ws4.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$ws4.Range("A:A").ColumnWidth=26; $ws4.Range("B:I").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食+调节+心理策略 V3
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与调节"

SetC $ws5 1 1 "增肌营养  训练窗口  自主调节  心理策略" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

# Nutrition
$r=3
SetC $ws5 $r 1 "训练前后营养窗口——V3精确化" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你每天教课 身体一直处于低强度消耗。不注意训练前后营养窗口=练了但没恢复。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=22; $r++

SetC $ws5 $r 1 "时机" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws5 $r 2 5
SetC $ws5 $r 6 "逻辑" $true 10 $white $headerBlue $false; MC $ws5 $r 6 9; $r++

$peri=@(
    @("训练前2h","碳水40-50g+蛋白20-25g`n香蕉+蛋白粉+全麦吐司2片","提供能量 防分解。教课消耗大 不吃=练到一半没力气"),
    @("训练后1h内","碳水50-60g+蛋白30-35g`n米饭+鸡胸肉200g/蛋白粉2勺+香蕉+牛奶","补糖原 启动修复。训练后1h=蛋白质吸收最高效时段"),
    @("教课日(非训练日)","每4-5h吃一顿含蛋白质的餐","教课也在消耗——不要等到训练日才吃够"),
    @("休息日","碳水比训练日少20-30%","不练不需要那么多能量。蛋白质不变")
)
foreach($pw in $peri){
    SetC $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pw[1] $false 10 0 0 $true; MC $ws5 $r 2 5
    SetC $ws5 $r 6 $pw[2] $false 10 0 0 $true; MC $ws5 $r 6 9
    $ws5.Range("A$($r):I$($r)").RowHeight=38; SB $ws5 $r 1 9; $r++
}

$r++
SetC $ws5 $r 1 "全天：总热量2700-2900 蛋白质160-180g(2.0-2.2g/kg) 碳水320-360g 脂肪60-70g 水3-4L。干净增肌 不做脏增肌。教一天课 碳水是能量来源 不要怕吃碳水。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

# Self-regulation
$r+=2
SetC $ws5 $r 1 "自主调节框架——这套可以带走" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$selfReg=@(
    "1. 定分化：每周几天->推拉腿/上下肢/全身。你4天=推+拉(宽/厚)+腿(代谢)+上肢(预疲劳)",
    "2. 定DUP主题：每天不同刺激——不让身体适应",
    "3. 定周期：3周渐进+1周减载(换动作)+季度全休",
    "4. 定RPE：百分比是理论值 RPE是当天真实状态",
    "5. 每月复盘：AMRAP次数+力量+体重+照片->决定下周期方向",
    "在西宁健身房 任何一个会员身上都能用这套框架"
)
foreach($sr in $selfReg){
    SetC $ws5 $r 1 $sr $false 10 0 0 $true; MC $ws5 $r 1 9
    $ws5.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# V3: 心理策略
$r+=2
SetC $ws5 $r 1 "状态差时的心理策略——V3新增" $true 14 $white $softPink $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$softPink
$r++

SetC $ws5 $r 1 "你最大的挑战不是不知道怎么练——是教了一天课 已经没有精力给自己练了。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=22; $r++

SetC $ws5 $r 1 "状态" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "策略" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$psych=@(
    @("教课站一天 累 不想练","去健身房 只做热身+主项3组。如果还是不想练->回家 今天算完成。如果感觉来了->继续。你只是需要'开始'。"),
    @("连续几天状态差","检查睡眠和饮食。睡眠不足->优先补觉 训练降级。饮食没吃够->加碳水。"),
    @("连续2周没进步","不是你的问题 是计划需要调整。可能：累积疲劳/动作需换变式/热量不够。解决：减载1周+换变式+加碳水。"),
    @("教课太累 不想练自己的","把训练当成'给自己充电'。你每天在输出(教课) 训练是唯一属于你的输入时间。哪怕只练20分钟。"),
    @("觉得自己练得没会员好","你是教练 不是选手。训练不需要和别人比——只需要比上周的自己好。AMRAP+1次 深蹲+2.5kg 都是进步。")
)
foreach($ps in $psych){
    SetC $ws5 $r 1 $ps[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $ps[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=42; SB $ws5 $r 1 9; $r++
}

# FAQ
$r+=2
SetC $ws5 $r 1 "常见问题" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "问题" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "答案" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$faq=@(
    @("AMRAP次数连续2周卡住","换变式或提前减载——可能累积疲劳了"),
    @("腿日45分钟没完成","减超级组轮数——宁可做少 不要超时"),
    @("预疲劳后主项重量比预期低","正常——预疲劳目的就是这个。泵感才是指标"),
    @("体重2周没涨","加碳水——每天多一碗饭。干净增肌慢是正常的"),
    @("季度全休到期","新周期从全休前第3周70-75%起步")
)
foreach($fq in $faq){
    SetC $ws5 $r 1 $fq[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $fq[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=26; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用+复盘+年度
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  4周复盘  年度周期  季度全休提醒" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

# Simplified
$r=3
SetC $ws6 $r 1 "如果这周只能2-3练——精简版" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "选项" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "训练日1" $true 10 $white $headerBlue $false; MC $ws6 $r 2 4
SetC $ws6 $r 5 "训练日2" $true 10 $white $headerBlue $false; MC $ws6 $r 5 7
SetC $ws6 $r 8 "训练日3(3练)" $true 10 $white $headerBlue $false; MC $ws6 $r 8 9; $r++

$simple=@(
    @("2练","推+腿(精简)：卧推4x5+1AMRAP+深蹲4x8-10+推举+提踵","拉+上肢(精简)：划船+引体+窄距卧推+下拉+弯举+侧平举",""),
    @("3练","推+腿(精简)","拉+上肢(精简)","全身综合(精简)：深蹲+卧推+划船+推举+面拉+核心")
)
foreach($so in $simple){
    SetC $ws6 $r 1 $so[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $so[1] $false 10 0 0 $true; MC $ws6 $r 2 4
    SetC $ws6 $r 5 $so[2] $false 10 0 0 $true; MC $ws6 $r 5 7
    SetC $ws6 $r 8 $so[3] $false 10 0 0 $true; MC $ws6 $r 8 9
    $ws6.Range("A$($r):I$($r)").RowHeight=42; SB $ws6 $r 1 9; $r++
}

# Review V3
$r+=2
SetC $ws6 $r 1 "4周后复盘——用数据说话" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么看" $true 10 $white $headerBlue $false; MC $ws6 $r 2 6
SetC $ws6 $r 7 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 7 9; $r++

$review=@(
    @("AMRAP次数","第1周vs第3周——涨了多少？","涨=加重。平/降=维持或换变式。"),
    @("深蹲8-10RM","第1周vs第3周","涨=加重。没涨=检查恢复+睡眠。"),
    @("拉日宽vs厚","哪个进步更快？","进步快的保持 慢的下周期多轮一次。"),
    @("体重趋势","0.5-1kg/月净肌肉=理想。涨超1kg/周=控热量。","—"),
    @("照片","第1天vs第28天 镜子比体重秤真实","—"),
    @("季度全休","第3个月末 该全休了","全休1周 新季度从全休前70-75%起步")
)
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 6
    SetC $ws6 $r 7 $rv[2] $false 10 0 0 $true; MC $ws6 $r 7 9
    $ws6.Range("A$($r):I$($r)").RowHeight=26; SB $ws6 $r 1 9; $r++
}

# Annual reminder
$r+=2
SetC $ws6 $r 1 "年度周期提醒" $true 14 $white $softBlue $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws6 $r 1 "你现在在Q1积累期。先跑完第一个4周 把AMRAP基准和深蹲基准建起来。每3个月记得全休1周——季度全休和减载周一样 是不可跳过的必修课。" $false 10 0 0 $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=24

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\吉\吉_增肌方案_V3.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
