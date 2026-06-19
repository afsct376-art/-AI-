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

SetC $ws0 1 1 "瑞克  力型兼备方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  168cm / 67kg  |  力型兼备  高翻进修中  午间训练" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  你练了3-4年 三大项都会 正在学高翻" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$intro=@(
    "你不是新手——深蹲卧推硬拉的动作模式已经刻在身体里了。力型兼备：既要力量和肌肉 也要能看。",
    "你的时间窗口是中午——午休那一小时多一点。方案压缩在60分钟内：热身8min+力量45min+拉伸5min。",
    "高翻是你正在学的技能——方案里单独留了15分钟的练习窗口 放在下肢日 精力最好的时候练。",
    "每周3天=DUP波动：上体力(重)+下肢力+高翻(技能)+全身增肌(容量)。3年基础+3天/周=每个肌群2次刺激。"
)
foreach($it in $intro){
    SetC $ws0 $row 1 $it $false 10 0 0 $true; MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}

$row+=2
SetC $ws0 $row 1 "  DUP 3天分化——力型兼备" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "强度/RPE" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "核心内容" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$dup=@(
    @("周一","上体力(大重量)","85-90% 1RM RPE8-9","卧推主项+引体+推举+划船+手臂","重而可控——冲击上半身力量上限"),
    @("周三","下肢力+高翻(技能)","80-87.5% 1RM RPE7-8","高翻练习+深蹲主项+RDL+核心","先磨技术 再冲力量——高翻放在精力最好的时候"),
    @("周五","全身增肌(容量)","70-77.5% 1RM RPE7-8","上斜卧推+前蹲+划船+侧平举+手臂超级组","泵感——短间歇 高次数 全身代谢压力")
)
foreach($d in $dup){
    SetC $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $d[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $d[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $d[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周周期——积累→强化→峰值→减载" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "上体力" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "下肢力+高翻" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "全身增肌" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","积累","85% 4x5 建基准","高翻技术 深蹲80% 4x5","70% 3x10-12 建容量基准"),
    @("2","强化","+2.5-5kg 4x5 RPE8-9","高翻+2.5kg 深蹲+5kg","缩短休息10s 或+1组"),
    @("3","峰值","再+2.5-5kg 4x3 RPE9","高翻挑战新重量 深蹲87.5% 4x3","再缩短休息 或次数+2"),
    @("4","减载","60% 3x5 技术打磨","高翻轻重量练速度 深蹲60%","50-60% 2x10-12 换动作")
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
SetC $ws0 $row 1 "  午间训练——时间紧 效率优先" $true 14 $white $accentOrange $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentOrange
$row++

$noon=@(
    @("练前","11:30左右吃一顿轻食：香蕉+蛋白粉/全麦吐司+蛋——提供训练能量 不撑胃"),
    @("热身","压缩到8分钟——但不跳过：划船机3min+动态拉伸3min+专项热身2min"),
    @("力量","45分钟高效完成——组间休息严格计时。主项=120s 辅助=60-90s 孤立=45s"),
    @("练后","13:00前喝蛋白粉+快碳(香蕉/饭团)——下午上班前把恢复窗口关上")
)
foreach($nn in $noon){
    SetC $ws0 $row 1 $nn[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $nn[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("完整方案逻辑+DUP框架"," Sheet 方案概览"),
    @("上体力 下肢力+高翻 全身增肌"," Sheet 训练计划"),
    @("渐进决策树+高翻进阶+器械突破"," Sheet 渐进决策树"),
    @("热身+午间策略+减载+恢复"," Sheet 恢复策略"),
    @("增肌饮食+午间窗口+判断规则"," Sheet 饮食与规则"),
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
$ws1=$wb.Worksheets.Add();$ws1.Name="方案概览"

SetC $ws1 1 1 "瑞克  力型兼备  方案概览" $true 18 $headerBlue 0 $true
MC $ws1 1 1 9; $ws1.Range("A1:I1").RowHeight=36

SetC $ws1 2 1 "男 31-35岁 168cm/67kg 体脂20-25% 3-4年训练经验 三大项掌握 高翻进修中 午间训练 力型兼备" $false 11 $gold $darkBg $true
MC $ws1 2 1 9; $ws1.Range("A2:I2").Font.Color=$white; $ws1.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws1 $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws1 $row 1 9

$ov=@(
    @("训练频率","每周3练(周一/三/五) DUP波动 午间60min高效完成","上体力(重)+下肢力+高翻(技能)+全身增肌(容量)"),
    @("核心框架","DUP 3天不同主题 4周周期(积累→强化→峰值→减载)","每个肌群每周刺激2次——自然训练者最优频率"),
    @("高翻专项","周三下肢日的前15分钟 精力最好时练习","高翻是技能——不是'练到累' 是'练到准'"),
    @("午间策略","练前11:30轻食 热身8min压缩版 力量45min严格计时 练后13:00前补碳水+蛋白","打工人午休训练=效率第一 质量不妥协"),
    @("方案特点","力型兼备 DUP波动 高翻进修 午间高效 自主调节框架")
)
$r=5
foreach($o in $ov){
    SetC $ws1 $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws1 $r 3 $o[1] $false 10 0 0 $true; MC $ws1 $r 3 6
    SetC $ws1 $r 7 $o[2] $false 10 0 0 $true; MC $ws1 $r 7 9
    SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "为什么力型兼备选DUP 3天" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$why=@(
    "你练了3-4年——身体对单一刺激已经适应。DUP让神经系统每周面对3种不同刺激：大重量(上体力)+技能+力量(下肢+高翻)+代谢压力(全身增肌)。",
    "每周3天=每个肌群2次刺激：卧推(周一)+上斜/推举(周五)=胸肌2次。深蹲(周三)+前蹲(周五)=股四头2次。引体(周一)+高位下拉(周五)=背阔肌2次。",
    "力型兼备=力量周期(第1-3周渐进加重)+增肌容量(周五高次数短间歇)。既要有举重的能力 也要有好看的体型。",
    "高翻是你主动在学的技能——方案里把高翻放在周三精力最好的时候练。高翻不是'练到累' 是'练到准'——15分钟技术窗口 轻重量磨动作。"
)
foreach($w in $why){
    SetC $ws1 $r 1 $w $false 10 0 0 $true; MC $ws1 $r 1 9
    $ws1.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$r++
SetC $ws1 $r 1 "RPE参考——3年基础 你不需要解释 但统一标准" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

SetC $ws1 $r 1 "RPE" $true 10 $white $headerBlue $false
SetC $ws1 $r 2 "感觉" $true 10 $white $headerBlue $false
SetC $ws1 $r 4 "剩余次数" $true 10 $white $headerBlue $false
SetC $ws1 $r 6 "用在哪天" $true 10 $white $headerBlue $false; $r++

$rpeRef=@(
    @("7-8","中等偏强——做完还能做3-4次","周五全身增肌——泵感和代谢积累"),
    @("8-9","很重——做完还能做1-2次","周一上体力——冲击上半身力量上限"),
    @("7-8(技能)","不是力竭——是动作质量","周三高翻——磨技术 不追求RPE"),
    @("10","极限力竭","本方案偶尔用——第3周峰值周最后一组")
)
foreach($rp in $rpeRef){
    SetC $ws1 $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    SetC $ws1 $r 2 $rp[1] $false 10 0 0 $false
    SetC $ws1 $r 4 $rp[2] $false 10 0 0 $false
    SetC $ws1 $r 6 $rp[3] $false 10 0 0 $true; MC $ws1 $r 6 9
    $ws1.Range("A$($r):I$($r)").RowHeight=26; SB $ws1 $r 1 9; $r++
}

$r++
SetC $ws1 $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws1 $r 1 9; $r++

$pr=@(
    @("1","力量+美感 不矛盾","周一冲大重量 周五做高次数——同一个人 两个维度。力型兼备=力量和肌肉一起涨。"),
    @("2","高翻是技能 不是力量训练","15分钟技术窗口。轻重量磨动作 不求重量。技术对了 重量自然来。"),
    @("3","午间效率=质量不妥协","时间紧但动作质量不能降。热身压缩但不跳过。组间严格计时——午休不等人。")
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
# SHEET 2: 训练计划
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="训练计划"

SetC $ws2 1 1 "训练计划——DUP 3天 力型兼备" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "周一=上体力(重) 周三=下肢力+高翻(技能) 周五=全身增肌(容量)。午间60min高效完成。第1-3周渐进 第4周减载。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 周一 上体力
$r=4
SetC $ws2 $r 1 "周一  上体力（大重量 Upper Strength）" $true 13 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=28
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=85% 4x5 建基准 | 第2周=+2.5-5kg 4x5 | 第3周=再+2.5-5kg 4x3 | 第4周=60% 3x5" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "负荷/RPE" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "策略" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "力型兼备" $true 10 $white $headerBlue $false
SetC $ws2 $r 9 "技术要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):I$($r)").RowHeight=26; $r++

$actsPush=@(
    @("杠铃卧推(主项)","4组","5次","85%1RM RPE8-9","120-150s","第2周+2.5-5kg","力的来源——上半身力量的标尺","【要点】起桥 肩胛收紧 脚踩实 瓦式呼吸。每组留1-2次力竭余地 不做到完全力竭。"),
    @("负重引体向上","4组","5-8次","自重或负重 RPE8","120s","第2周+2.5kg负重","型的来源——背阔肌宽度=倒三角","【要点】全幅度——底部手臂伸直 顶部下巴过杠。做不了8次=减负重 能做8次+=加负重。"),
    @("站姿杠铃推举","3组","6-8次","RPE7-8","90s","第2周+1-2.5kg","力的辅助——肩部力量 和卧推互补","【要点】核心收紧 不塌腰。杠铃从锁骨前起推至头顶 不锁肘。"),
    @("杠铃俯身划船","3组","8-10次","RPE7-8","90s","第2周+2.5kg","型的辅助——中背部厚度 和引体互补","【要点】俯身45-60度 杠铃贴大腿前侧拉。引体=垂直拉 划船=水平拉 背阔全角度。"),
    @("绳索下压+哑铃弯举超级组","各3组","各12-15次","轻-中重量 RPE7","不休息 组间60s","手臂超级组——效率和泵感","力型兼备的收尾——三头+二头同时充血","【要点】下压做完直接做弯举→休息60秒→重复。大臂固定 不借力。")
)
$i=1
foreach($a in $actsPush){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 9 0 0 $false
    SetC $ws2 $r 9 $a[8] $false 9 0 0 $true
    $ws2.Range("A$($r):I$($r)").RowHeight=52; SB $ws2 $r 1 9; $r++; $i++
}
$r+=2

# 周三 下肢力+高翻
SetC $ws2 $r 1 "周三  下肢力+高翻（技能+力量 Lower+Clean）" $true 13 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=28
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=高翻技术 深蹲80% 4x5 | 第2周=高翻+2.5kg 深蹲+5kg | 第3周=挑战新重量 | 第4周=轻量练速度" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "负荷/RPE" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "策略" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "力型兼备" $true 10 $white $headerBlue $false
SetC $ws2 $r 9 "技术要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):I$($r)").RowHeight=26; $r++

$actsClean=@(
    @("高翻技术练习(先做!)","5-6组","3次","40-60%极限 轻重量磨动作","90s","第2周+2.5kg——技术对了才加","力的技能——举重能力=全身爆发力的标尺","【要点】髋爆发+耸肩+展体。不是用手臂拉——用髋和三倍伸展。轻重量练速度 不求重。"),
    @("杠铃深蹲(主项)","4组","5次","80%1RM RPE7-8","120s","第2周+5kg","力的核心——下肢力量的绝对标尺","【要点】深度至少平行地面 膝盖不内扣。高翻后的深蹲=在神经系统已激活的状态下做。"),
    @("杠铃罗马尼亚硬拉","3组","8-10次","RPE7-8","90s","第2周+5kg","型的辅助——腘绳肌+臀=下肢后侧线条","【要点】膝盖微屈 全程直背 哑铃/杠铃贴腿下滑。和深蹲互补——深蹲前侧 RDL后侧。"),
    @("悬垂举腿","3组","12-15次","自重","45s","—","核心稳定——所有大重量动作的地基","【要点】下腹带动骨盆上卷 不靠惯性。高翻和深蹲都需要核心稳定——这是地基训练。"),
    @("站姿提踵","3组","15-20次","RPE7 膝盖微屈","30s","—","小腿——下肢完整性的收尾","【要点】顶峰停1秒 慢放。")
)
$i=1
foreach($a in $actsClean){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 9 0 0 $false
    SetC $ws2 $r 9 $a[8] $false 9 0 0 $true
    $ws2.Range("A$($r):I$($r)").RowHeight=52; SB $ws2 $r 1 9; $r++; $i++
}
$r+=2

# 周五 全身增肌
SetC $ws2 $r 1 "周五  全身增肌（容量 Full Body Hypertrophy）" $true 13 $white $darkBg $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=28
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第1周=70% 3x10-12 建容量基准 | 第2周=缩短休息10s或+1组 | 第3周=再缩短或次数+2 | 第4周=50-60% 2x10-12" $false 9 $headerBlue $lightGray $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=20; $r++

SetC $ws2 $r 1 "序" $true 10 $white $headerBlue $false
SetC $ws2 $r 2 "动作" $true 10 $white $headerBlue $false
SetC $ws2 $r 3 "组数" $true 10 $white $headerBlue $false
SetC $ws2 $r 4 "次数" $true 10 $white $headerBlue $false
SetC $ws2 $r 5 "负荷/RPE" $true 10 $white $headerBlue $false
SetC $ws2 $r 6 "休息" $true 10 $white $headerBlue $false
SetC $ws2 $r 7 "策略" $true 10 $white $headerBlue $false
SetC $ws2 $r 8 "力型兼备" $true 10 $white $headerBlue $false
SetC $ws2 $r 9 "技术要点" $true 10 $white $headerBlue $true
$ws2.Range("A$($r):I$($r)").RowHeight=26; $r++

$actsHyper=@(
    @("哑铃上斜卧推","3组","10-12次","RPE7-8","60-75s","第2周缩短休息或+1组","型的核心——上胸+肩前束 和周一平板互补","【要点】凳角30-45度 下放至胸锁骨水平。和周一杠铃平板不同——哑铃自由轨迹 上胸主导。"),
    @("杠铃前蹲","3组","10-12次","RPE7-8","75-90s","第2周+2.5-5kg","型的核心——股四头+核心稳定 和周三后蹲互补","【要点】杠铃架在锁骨前 躯干更竖直。股四头主导——和周三杠铃后蹲不同。"),
    @("高位下拉(宽握)","3组","12-15次","RPE7","60s","—","型的辅助——背阔肌宽度 和周一引体互补","【要点】宽握 拉至锁骨 顶峰背阔肌向外展。容量日的高位下拉=轻量高次数 和引体互补。"),
    @("哑铃侧平举","4组","15-20次","轻重量 RPE7","45s","—","型的收尾——三角肌中束=肩宽","【要点】不耸肩 手臂微屈。高次数追求泵感——肩膀是增肌日的重头戏。"),
    @("绳索下压+哑铃弯举超级组","各3组","各12-15次","RPE7","不休息 组间45s","手臂超级组——泵感收尾","力型兼备的型=手臂线条。短间歇 高次数 充血拉满。")
)
$i=1
foreach($a in $actsHyper){
    SetC $ws2 $r 1 $i $false 10 0 $lightGray $false
    SetC $ws2 $r 2 $a[1] $true 10 $headerBlue 0 $false
    SetC $ws2 $r 3 $a[2] $false 10 0 0 $false
    SetC $ws2 $r 4 $a[3] $false 10 0 0 $false
    SetC $ws2 $r 5 $a[4] $false 10 0 0 $false
    SetC $ws2 $r 6 $a[5] $false 10 0 0 $false
    SetC $ws2 $r 7 $a[6] $false 10 0 0 $false
    SetC $ws2 $r 8 $a[7] $false 9 0 0 $false
    SetC $ws2 $r 9 $a[8] $false 9 0 0 $true
    $ws2.Range("A$($r):I$($r)").RowHeight=52; SB $ws2 $r 1 9; $r++; $i++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=18
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=10
$ws2.Range("E:E").ColumnWidth=16; $ws2.Range("F:F").ColumnWidth=10
$ws2.Range("G:G").ColumnWidth=16; $ws2.Range("H:H").ColumnWidth=18
$ws2.Range("I:I").ColumnWidth=48
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进决策树
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进决策树"

SetC $ws3 1 1 "渐进决策树  高翻进阶  器械突破" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34
SetC $ws3 2 1 "训练像做实验——每次都是数据点。以下决策树覆盖三个训练日+高翻专项。" $false 10 $headerBlue $lightGray $true
MC $ws3 2 1 9

# 4周渐进
$r=4
SetC $ws3 $r 1 "4周渐进（DUP框架内 每周3天主题不变 逐周递进）" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "上体力(周一)" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "高翻(周三)" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "下肢力(周三)" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "全身增肌(周五)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("1","85% 4x5 建基准","技术练习 40-60%","80% 4x5 建基准","70% 3x10-12 建容量基准"),
    @("2","+2.5-5kg 4x5","+2.5kg(技术对才加)","+5kg","缩短休息10s 或+1组"),
    @("3","再+2.5-5kg 4x3","挑战新重量","87.5% 4x3","再缩短休息 或次数+2"),
    @("4","60% 3x5 技术打磨","轻重量练速度","60% 3x5","50-60% 2x10-12 换动作")
)
foreach($pw in $prog){
    SetC $ws3 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pw[1] $false 9 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pw[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $pw[3] $false 9 0 0 $false
    SetC $ws3 $r 6 $pw[4] $false 9 0 0 $true; MC $ws3 $r 6 9
    $ws3.Range("A$($r):I$($r)").RowHeight=30; SB $ws3 $r 1 9; $r++
}

# 高翻进阶
$r+=2
SetC $ws3 $r 1 "高翻进阶——技能优先 不是重量优先" $true 14 $white $softBlue $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 5 9; $r++

$cleanTree=@(
    @("动作流畅 杠铃轨迹直 接杠稳","技术对了——可以+2.5kg","每次只加2.5kg 不求多。动作变形→退回。"),
    @("手臂过早发力弯肘 用二头拉","技术不对——不要加重","减重练动作。空杆/轻重量磨三倍伸展——髋-膝-踝。"),
    @("接杠时杠铃砸锁骨/手腕疼","接杠位置不对——不加重","练前蹲+杠铃耸肩 找锁骨前架杠的感觉。"),
    @("连续3次训练 动作流畅+接杠稳","可以挑战次极限重量(85-90%)","但只在技术完美时尝试。高翻不是1RM——是3RM。" ),
    @("今天状态差/下午累了","不练高翻——改练杠铃耸肩+前蹲","高翻是最高神经需求的动作——状态不对练了反而退步。")
)
foreach($ct in $cleanTree){
    SetC $ws3 $r 1 $ct[0] $true 10 $headerBlue $lightGray $false; MC $ws3 $r 1 2
    SetC $ws3 $r 3 $ct[1] $false 10 0 0 $false
    SetC $ws3 $r 5 $ct[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=28; SB $ws3 $r 1 9; $r++
}

# Decision trees for each day
$r+=2
SetC $ws3 $r 1 "各训练日渐进决策树" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "情况" $true 10 $white $headerBlue $false; MC $ws3 $r 2 4
SetC $ws3 $r 5 "判断" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 7 9; $r++

$treeAll=@(
    @("周一上体力","4x5轻松完成 RPE实际只有7","太轻了","下次+2.5-5kg"),
    @("周一上体力","4x5第3组就变形 RPE到10","太重了","退回上周重量——RPE优先"),
    @("周三深蹲","4x5完成后还有余力","可以加重","下周+2.5-5kg"),
    @("周五全身增肌","组间休息觉得太长","代谢压力不够","缩短10-15秒"),
    @("周五全身增肌","某动作连续2周没进步","需要换刺激","换变式——同肌群不同角度"),
    @("连续2周恢复不过来","训练+上班 双重疲劳","该减载了","提前进入第4周减载")
)
foreach($ta in $treeAll){
    SetC $ws3 $r 1 $ta[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $ta[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $ta[2] $false 10 0 0 $false
    SetC $ws3 $r 7 $ta[3] $false 10 0 0 $true; MC $ws3 $r 7 9
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
    @("1.5次法","下放 半起 全下 全起=1次","周一周三——双倍刺激"),
    @("缩短休息","90s 60s 45s","周五增肌——核心策略"),
    @("暂停卧推/深蹲","底部停顿2秒 爆发起","突破粘滞点")
)
foreach($b in $bt){
    SetC $ws3 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $b[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $b[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=22; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "热身(午间版)  主动恢复  减载周" $true 16 $headerBlue 0 $true
MC $ws4 1 1 9; $ws4.Range("A1:I1").RowHeight=34

$r=3
SetC $ws4 $r 1 "午间热身——压缩到8分钟 但不跳过" $true 14 $white $darkBg $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "阶段" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "时间" $true 10 $white $headerBlue $false; $r++

$wu=@(
    @("一般热身","划船机/单车——轻度有氧升温","3min"),
    @("动态拉伸","猫式+最伟大伸展+摆腿+肩环绕","2min"),
    @("神经激活","药球砸地/跳箱(低)/弹力带侧走","1min"),
    @("专项热身","主项空杆/轻重量 1-2组x5-8次 递增至正式","2min")
)
foreach($w in $wu){
    SetC $ws4 $r 1 $w[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $w[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $w[2] $false 10 0 0 $false; MC $ws4 $r 6 9
    $ws4.Range("A$($r):I$($r)").RowHeight=24; SB $ws4 $r 1 9; $r++
}

# Active recovery
$r+=2
SetC $ws4 $r 1 "休息日主动恢复（周二/四/六/日 选1-2天）" $true 14 $white $softBlue $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$active=@(
    "普通打工人——下班后可能没精力专门做恢复。不需要复杂：",
    "  泡沫轴10min：小腿→股四→腘绳→臀→背→胸 每部位30秒",
    "  散步20-30min：心率<120bpm 不是训练 是恢复",
    "  拉伸5min：髋屈肌+腘绳肌+胸+背 各30秒",
    "  睡够7-8小时：这比任何恢复手段都重要——你是午间训练 睡眠质量直接影响训练表现"
)
foreach($ar in $active){
    SetC $ws4 $r 1 $ar $false 10 0 0 $true; MC $ws4 $r 1 9
    $ws4.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# Deload
$r+=2
SetC $ws4 $r 1 "第4周减载——换动作 不是做轻一点" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "原动作" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "第4周换成" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "参数" $true 10 $white $headerBlue $false; MC $ws4 $r 4 9; $r++

$deload=@(
    @("杠铃卧推","哑铃卧推(轻量)或俯卧撑","2-3组x8-10 RPE4-5"),
    @("杠铃深蹲","高脚杯深蹲(轻哑铃)","2组x10-12 RPE4-5"),
    @("杠铃划船","哑铃单臂划船(轻量)","2组x10/侧"),
    @("高翻","杠铃耸肩+前蹲(轻量 练分解动作)","2组x8-10"),
    @("负重引体","高位下拉(轻量)","2组x10-12"),
    @("有氧","低强度散步/休闲骑","20-30min <120bpm"),
    @("心态","做完应该感觉没练够——这就对了","")
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
# SHEET 5: 饮食与规则
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与规则"

SetC $ws5 1 1 "增肌饮食  午间窗口  判断规则" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

$r=3
SetC $ws5 $r 1 "午间训练——营养窗口策略" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你中午训练——不同于晚上练完回家吃饭。营养窗口需要精确：练前不能空腹 练后要在下午上班前吃完。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

SetC $ws5 $r 1 "时机" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws5 $r 2 5
SetC $ws5 $r 6 "为什么" $true 10 $white $headerBlue $false; MC $ws5 $r 6 9; $r++

$peri=@(
    @("练前 11:30","香蕉1根+蛋白粉/全麦吐司2片+2蛋","轻食 不撑胃——提供训练能量 防止训练中分解肌肉"),
    @("练后 13:00前","米饭+鸡胸肉/蛋白粉2勺+香蕉+牛奶","补糖原 启动肌肉修复。下午上班前把恢复窗口关上"),
    @("下午加餐 15:00","小份坚果/水果/无糖酸奶","如果下午饿了就吃——不饿可以不吃"),
    @("晚餐","正常吃：碳水一拳+蛋白质+蔬菜","和平时一样——不需要因为中午训练就晚餐吃特殊")
)
foreach($pw in $peri){
    SetC $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pw[1] $false 10 0 0 $true; MC $ws5 $r 2 5
    SetC $ws5 $r 6 $pw[2] $false 10 0 0 $true; MC $ws5 $r 6 9
    $ws5.Range("A$($r):I$($r)").RowHeight=36; SB $ws5 $r 1 9; $r++
}

$r++
SetC $ws5 $r 1 "全天营养方向" $true 12 $headerBlue 0 $true; MC $ws5 $r 1 9; $r++

SetC $ws5 $r 1 "总热量2500-2700 蛋白质130-150g(2.0-2.2g/kg) 碳水300-340g 脂肪55-65g 水3L。力型兼备=既要长力量也要长肌肉。热量不能少——但干净增肌 不做脏增肌。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

# Rules
$r+=2
SetC $ws5 $r 1 "判断规则——训练和身体信号" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "信号" $true 10 $white $headerBlue $false; MC $ws5 $r 1 2
SetC $ws5 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$rules=@(
    @("卧推/深蹲4x5轻松完成 RPE实际7","下次+2.5-5kg——你变强了"),
    @("4x5第3组就变形 RPE到10","退回上周重量——今天的你不如上周 正常"),
    @("高翻技术连续3次训练流畅","+2.5kg——但只在技术完美时加"),
    @("午间训练感觉没力气","检查练前那顿——是不是没吃够？香蕉+吐司不能省"),
    @("中午练完下午犯困","练后碳水可能偏多——减半碗米饭 蛋白不变"),
    @("体重2周没涨","加碳水——每天多一碗饭。力型兼备需要热量盈余"),
    @("某天状态极差","降级：热身+主项3组轻重量+拉伸。不跳训练——但降强度")
)
foreach($r3 in $rules){
    SetC $ws5 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false; MC $ws5 $r 1 2
    SetC $ws5 $r 3 $r3[1] $false 10 0 0 $true; MC $ws5 $r 3 9
    $ws5.Range("A$($r):I$($r)").RowHeight=28; SB $ws5 $r 1 9; $r++
}

$ws5.Range("A:A").ColumnWidth=30; $ws5.Range("B:I").ColumnWidth=18
Write-Host "Sheet 5 done"

# ============================================
# SHEET 6: 备用与复盘
# ============================================
$ws6=$wb.Worksheets.Add();$ws6.Name="备用与复盘"

SetC $ws6 1 1 "精简版  4周复盘" $true 16 $headerBlue 0 $true
MC $ws6 1 1 9; $ws6.Range("A1:I1").RowHeight=34

$r=3
SetC $ws6 $r 1 "如果这周只能2练——精简版" $true 14 $white $accentOrange $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

SetC $ws6 $r 1 "选项" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "训练日1" $true 10 $white $headerBlue $false; MC $ws6 $r 2 4
SetC $ws6 $r 5 "训练日2" $true 10 $white $headerBlue $false; MC $ws6 $r 5 9; $r++

$simple=@(
    @("2练","上体力(精简)：卧推4x5+引体4x5+推举+划船+手臂","下肢力+全身(精简)：高翻+深蹲4x5+RDL+前蹲3x10+侧平举+手臂"),
    @("原则","保留核心复合动作 压缩辅助。高翻不跳——放在精力最好的时候 哪怕只做3组。","")
)
foreach($so in $simple){
    SetC $ws6 $r 1 $so[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $so[1] $false 10 0 0 $true; MC $ws6 $r 2 4
    SetC $ws6 $r 5 $so[2] $false 10 0 0 $true; MC $ws6 $r 5 9
    $ws6.Range("A$($r):I$($r)").RowHeight=42; SB $ws6 $r 1 9; $r++
}

$r+=2
SetC $ws6 $r 1 "4周后复盘——力型兼备的双重指标" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "指标" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "第1周vs第3周" $true 10 $white $headerBlue $false; MC $ws6 $r 2 5
SetC $ws6 $r 6 "下周期" $true 10 $white $headerBlue $false; MC $ws6 $r 6 9; $r++

$review=@(
    @("力-卧推/深蹲","卧推4x5重量涨了多少？深蹲4x5涨了多少？","涨=加重。平/降=维持或换变式"),
    @("力-高翻","技术流畅了吗？重量能加吗？","流畅=继续加2.5kg。仍卡=继续磨技术"),
    @("型-体重","涨了0.5-1kg净肌肉？","涨=继续。没涨=加碳水"),
    @("型-照片","第1天vs第28天——镜子比秤真实","—"),
    @("力型兼备","力量涨了+体型变了=力型兼备在发生","继续这个方向——DUP 3天框架可以一直用")
)
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 5
    SetC $ws6 $r 6 $rv[2] $false 10 0 0 $true; MC $ws6 $r 6 9
    $ws6.Range("A$($r):I$($r)").RowHeight=28; SB $ws6 $r 1 9; $r++
}

$r+=2
SetC $ws6 $r 1 "自主调节框架——3年基础 你可以自己迭代" $true 14 $white $accentGreen $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

$self=@(
    "1. 定分化：每周3天 → DUP(上体力+下肢力+高翻+全身增肌)。想换成推拉腿？可以——但3天频率DUP最有效。",
    "2. 定周期：3周渐进+1周减载 4周一循环。每3个月全休1周。",
    "3. 定RPE：周一8-9(重) 周三7-8(重+技能) 周五7-8(泵感)。百分比是理论值 RPE是真实状态。",
    "4. 每月复盘：卧推/深蹲力量+体重+照片。力型兼备——两个维度都要看。",
    "你现在已经有了一个能跑一辈子的训练框架——它会跟着你一起进化。"
)
foreach($s in $self){
    SetC $ws6 $r 1 $s $false 10 0 0 $true; MC $ws6 $r 1 9
    $ws6.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\瑞克\瑞克_力型兼备方案_V1.xlsx"
New-Item -ItemType Directory -Force -Path "D:\Codex\members\瑞克" | Out-Null
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
