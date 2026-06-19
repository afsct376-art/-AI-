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
    SetC $ws $ref 8 "V2变化/策略" $true 10 $white $headerBlue $false
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

SetC $ws0 1 1 "吉  增肌方案 V2" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  182cm / 82kg  |  增肌  |  V2 = 奥赛级底层逻辑" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  V2 升级了什么——从'优秀教练'到'奥赛逻辑'" $true 16 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$upgrades=@(
    @("  推日","5x5  4x5+1xAMRAP——最后一组探底 找到当天真实极限。减少神经疲劳 不牺牲质量。"),
    @("  拉日","拆成'宽度优先'和'厚度优先'两个版本 每周轮换。背=视觉冲击力+立体感 两个维度都要。"),
    @("  腿日","代谢压力限时45分钟——刺激够了 皮质醇还没起来。超时=分解肌肉 得不偿失。"),
    @("  上肢日","从'补弱点'升级为'预疲劳法'：孤立力竭  复合动作。突破'做完了但目标肌群还没感觉'的瓶颈。"),
    @("  减载周","换动作 不是做轻一点。同样动作做轻量=关节应力路径相同。换动作=给关节真正的休息。"),
    @("  营养","训练前后窗口精确化。季度全休周——每3个月1周完全休息 不是偷懒 是职业保护。")
)
foreach($up in $upgrades){
    SetC $ws0 $row 1 $up[0] $true 10 $softPurple 0 $false
    SetC $ws0 $row 2 $up[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=30; $row++
}

$row+=2
SetC $ws0 $row 1 "  DUP 4天分化——V2版" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

SetC $ws0 $row 1 "训练日" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "V2核心变化" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "RPE" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "训练感觉" $true 10 $white $headerBlue $false; $row++

$dup=@(
    @("周一","推日(大重量)","4x5+1xAMRAP 卧推","8-9","前4组标准 最后一组探底"),
    @("周二","拉日(宽度/厚度轮换)","每周轮换 宽度优先或厚度优先","7-8","背的宽度+厚度双维度发展"),
    @("周四","腿日(代谢 45min限时)","深蹲+腿举+腿弯举超级组","7-8","短间歇 45min内完成=黄金窗口"),
    @("周五","上肢日(预疲劳)","孤立力竭  复合动作","7-8","先充血再深度刺激 突破瓶颈")
)
foreach($d in $dup){
    SetC $ws0 $row 1 $d[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $d[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $d[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $d[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $d[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  四周周期" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "拉日版本" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "推/腿/上肢" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "每周目标" $true 10 $white $headerBlue $false; $row++

$weeks=@(
    @("1","积累","宽度优先","80% 4x5+1AMRAP 建基准","建立基准+宽度刺激"),
    @("2","递增","厚度优先","+2.5-5kg RPE8-9 保持4x5+1AMRAP","递增+厚度刺激"),
    @("3","突破","宽度优先 加重","再+2.5-5kg 冲击AMRAP次数","突破+宽度冲击"),
    @("4","减载","弹力带+自重 换动作","60% 3x5 换动作","关节休息 身体偷偷变强")
)
foreach($w4 in $weeks){
    SetC $ws0 $row 1 $w4[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $w4[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $w4[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $w4[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $w4[4] $false 10 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  如果这周只能2-3练——精简版" $true 14 $white $accentOrange $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentOrange
$row++

$simpleVer=@(
    @("2练","推+腿(精简)：卧推4x5+1AMRAP+深蹲4x8-10+推举+提踵","拉+上肢(精简)：划船+引体+哑铃卧推+下拉+弯举+侧平举"),
    @("3练","推+腿(精简)+拉+上肢(精简)+全身综合(精简)","每周每个肌群仍2次刺激")
)
foreach($sv in $simpleVer){
    SetC $ws0 $row 1 $sv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $sv[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=32; SB $ws0 $row 1 5; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+周期框架+RPE指南"," Sheet 方案概览"),
    @("推日 拉日(宽/厚) 腿日(45min) 上肢日(预疲劳)"," Sheet 训练计划"),
    @("渐进负荷+器械突破"," Sheet 渐进突破"),
    @("热身+主动恢复+减载(换动作)+季度全休"," Sheet 恢复策略"),
    @("增肌营养+训练前后窗口+自主调节"," Sheet 饮食与调节"),
    @("精简版+4周复盘+季度全休"," Sheet 备用与复盘")
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

SetC $ws 1 1 "吉  增肌方案  概览 V2" $true 18 $headerBlue 0 $true
MC $ws 1 1 9; $ws.Range("A1:I1").RowHeight=36

SetC $ws 2 1 "男  25-30岁  182cm/82kg  体脂20-25%  健身教练  增肌  商业健身房  DUP+预疲劳+宽度/厚度轮换" $false 11 $gold $darkBg $true
MC $ws 2 1 9; $ws.Range("A2:I2").Font.Color=$white; $ws.Range("A2:I2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 9

$ov=@(
    @("训练频率","4次/周 推/拉/腿/上肢","DUP波动——每天不同刺激"),
    @("核心升级","4x5+1AMRAP 宽度/厚度轮换 45min限时 预疲劳法 换动作减载 季度全休","6项升级从'优秀'到'奥赛级'"),
    @("周期结构","4周一小周期(3周渐进+1周减载) 3小周期+1周全休=1中周期","细水长流——不是冲一个周期就停"),
    @("加重原则","上肢+1-2.5kg/下肢+2.5-5kg RPE优先于百分比","—"),
    @("方案特点","奥赛级底层逻辑 预疲劳/AMRAP/宽度厚度轮换/45min窗口/季度全休")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 6
    SetC $ws $r 7 $o[2] $false 10 0 0 $true; MC $ws $r 7 9
    SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "DUP + V2升级逻辑" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$dupWhy=@(
    "DUP波动——每周4天不同刺激 神经系统始终保持对新挑战的敏感。V2在此基础上加了6项奥赛级微调。",
    "推日4x5+1AMRAP：前四组积累标准刺激 最后一组探底。5x5的第五组往往因疲劳变形——AMRAP让你在质量不降的前提下找到真实极限。",
    "拉日宽度/厚度每周轮换：宽度=视觉冲击力(宽握下拉+飞鸟) 厚度=立体感(杠铃划船+T杠)。每周换一次 两个维度都发展。",
    "腿日45分钟限时：代谢压力增肌的黄金窗口。超时=皮质醇开始分解肌肉。短间歇+超级组 在45分钟内完成。",
    "上肢日预疲劳法：先用孤立动作让目标肌群充血 再用复合动作深度刺激。突破'做完了但没感觉'的瓶颈。",
    "减载周换动作：不是做轻一点——是换动作。同样动作做轻量=关节应力路径相同。换动作=真正的关节休息。",
    "季度全休：每3个月1周完全休息。你教一天课 恢复压力比普通人大一倍——季度全休是职业保护。"
)
foreach($dw in $dupWhy){
    SetC $ws $r 1 $dw $false 10 0 0 $true; MC $ws $r 1 9
    $ws.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# RPE
$r++
SetC $ws $r 1 "RPE参考" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "RPE" $true 10 $white $headerBlue $false
SetC $ws $r 2 "感觉" $true 10 $white $headerBlue $false
SetC $ws $r 4 "剩余次数" $true 10 $white $headerBlue $false
SetC $ws $r 6 "用在哪天" $true 10 $white $headerBlue $false; $r++

$rpeRef=@(
    @("6-7","中等——做完还能做4-6次","拉日/上肢日——泵感和代谢积累"),
    @("7-8","中等偏强——做完还能做3-4次","腿日——增肌的理想区间"),
    @("8-9","很重——做完还能做1-2次","推日AMRAP——最后一组探到RPE9"),
    @("10","极限力竭","本方案几乎不用。AMRAP最后一组的最后一次可能触达——够了。")
)
foreach($rp in $rpeRef){
    SetC $ws $r 1 $rp[0] $true 12 $headerBlue $lightGray $false
    SetC $ws $r 2 $rp[1] $false 10 0 0 $false
    SetC $ws $r 4 $rp[2] $false 10 0 0 $false
    SetC $ws $r 6 $rp[3] $false 10 0 0 $true; MC $ws $r 6 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "全年周期框架" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

SetC $ws $r 1 "层级" $true 10 $white $headerBlue $false
SetC $ws $r 2 "时长" $true 10 $white $headerBlue $false
SetC $ws $r 4 "结构" $true 10 $white $headerBlue $false; MC $ws $r 4 6
SetC $ws $r 7 "关键" $true 10 $white $headerBlue $false; MC $ws $r 7 9; $r++

$yearly=@(
    @("大周期","1年","增肌至85-88kg 体脂维持18-22%","每季度评估 调整方向"),
    @("中周期","1季度","3小周期(每月)+1周全休","全休周=系统重置——不是偷懒"),
    @("小周期","1个月","3周DUP+1周减载(换动作)","3+1=标准周期单元"),
    @("全休周","每3个月1次","散步/游泳/拉伸/泡沫轴 不碰器械","教课+训练=双倍恢复需求 全休是职业保护")
)
foreach($y in $yearly){
    SetC $ws $r 1 $y[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $y[1] $false 10 0 0 $false
    SetC $ws $r 4 $y[2] $false 10 0 0 $true; MC $ws $r 4 6
    SetC $ws $r 7 $y[3] $false 10 0 0 $true; MC $ws $r 7 9
    $ws.Range("A$($r):I$($r)").RowHeight=26; SB $ws $r 1 9; $r++
}

$r++
SetC $ws $r 1 "核心原则" $true 14 $headerBlue $lightBlue $true; MC $ws $r 1 9; $r++

$pr=@(
    @("1","动作质量 > 组数","5x5的第五组变形=得不偿失。4x5+1AMRAP让你在质量不降的前提下探底。健美选手的底层逻辑——质量永远优先。"),
    @("2","RPE优先于百分比","80% 1RM是理论值。你当天教课站了多久 会员多不多 睡没睡好——这些影响实际表现。RPE让你和身体对话。"),
    @("3","恢复=训练的一部分","你教一天课+自己练=双倍消耗。季度全休周不是偷懒 是职业保护。身体不是在训练时变强——是在恢复时。")
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

SetC $ws2 1 1 "DUP 4天分化 V2——推/拉(宽厚轮换)/腿(45min)/上肢(预疲劳)" $true 16 $headerBlue 0 $true
MC $ws2 1 1 9; $ws2.Range("A1:I1").RowHeight=34
SetC $ws2 2 1 "周一=推日(4x5+1AMRAP) 周二=拉日(宽度/厚度每周轮换) 周四=腿日(45min限时) 周五=上肢日(预疲劳法)。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 9

# 推日 V2
$r=4
$actsPush=@(
    @("杠铃平板卧推(主项)","4组+1AMRAP","5次","前4组80%1RM 第5组AMRAP","120-150s","8-9","V2:5x5  4x5+1AMRAP——第5组探底 找真实极限。不追求组数 追求质量。","【呼吸】瓦式——下放吸气屏住 推起中途呼气。【要点】起桥 肩胛收紧 脚踩实 前4组RPE8-9 第5组做到技术变形前停。记录AMRAP次数——这是每周进步的标尺。"),
    @("上斜哑铃卧推","4组","8-10次","RPE 7-8 上胸补充","90s","7-8","和V1相同——哑铃自由轨迹 上胸+肩前束","【要点】凳角30-45度 下放至胸锁骨水平。推日第二个推类动作——在胸肌疲劳状态下追求上胸泵感。"),
    @("坐姿哑铃推举","4组","8-10次","RPE 7-8","75s","7-8","第2周+1-2kg","【要点】下背贴靠背 不反弓。推日肩部主项——三角肌前束+中束。"),
    @("双杠臂屈伸","3组","8-12次","自重或负重 RPE 7-8","60s","7-8","第2周+负重腰带","【要点】身体稍前倾=胸+三头。身体竖直=三头为主。"),
    @("绳索飞鸟","3组","12-15次","轻重量 追求泵感","45s","6-7","—","【要点】手臂微屈 胸肌驱动。推日在主项胸肌疲劳后继续刺激代谢。"),
    @("三头绳索下压","3组","12-15次","轻-中重量","45s","7","—","【要点】大臂固定 压到底微外旋。AMRAP后三头已高疲劳——下压收尾效率最高。")
)
$r=Write-TrainingBlock $ws2 $r "周一  推日（大重量 4x5+1AMRAP）" "杠铃卧推4x5+1AMRAP主项+上斜哑铃+推举+臂屈伸+飞鸟+三头  |  45-55" "AMRAP=每周进步的标尺 记录次数" "45-55" $actsPush

# 拉日 V2 (宽度优先版)
$r++
SetC $ws2 $r 1 "拉日——V2新增：宽度优先版和厚度优先版 每周轮换" $true 14 $white $softBlue $true
MC $ws2 $r 1 9; $ws2.Range("A$($r):I$($r)").RowHeight=26
$ws2.Range("A$($r):I$($r)").Font.Color=$white; $ws2.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$actsPullWidth=@(
    @("宽握高位下拉(主项)","4组","10-12次","RPE 7-8 宽度优先","75s","7-8","第1/3周=宽度优先。第2周=杠铃划船见下方。","【要点】宽握 杆拉至锁骨 顶峰背阔肌向外撑开。宽度=视觉冲击力——背阔肌的外沿。  这是宽度日 不是重量日——感受背阔肌向外展。"),
    @("哑铃飞鸟/直臂下压","3组","15次","轻重量 追求背阔肌拉伸","45s","7","宽度日的辅助——背阔肌不同角度拉伸","【要点】直臂下压：手臂几乎伸直 从头顶下压至大腿。背阔肌的全程拉伸——宽度日的补充。"),
    @("窄握坐姿划船","3组","12-15次","RPE 7","60s","7","宽度日的厚度补充——窄握偏中背部","【要点】V柄拉至腹部 顶峰肩胛后缩1秒。宽度日也需要厚度——窄握划船补充中背部。"),
    @("罗马尼亚硬拉","4组","8-10次","RPE 7-8","90s","7-8","第2周+5kg","【要点】膝盖微屈 全程直背。拉日的下肢——腘绳肌+臀。和腿日深蹲互补。"),
    @("面拉(绳索)","3组","15次","轻重量","45s","6-7","—","【要点】拉向额头 肘外展。肩后束+上背——拉日的肩部补充。"),
    @("杠铃弯举","3组","10-12次","RPE 7","45s","7","—","【要点】大臂固定 顶峰外旋。拉日收尾——划船和下拉后二头已充血。")
)
$r=Write-TrainingBlock $ws2 $r "拉日A：宽度优先（Width Priority）" "宽握下拉主项+直臂下压+窄握划船+RDL+面拉+弯举  |  45-50" "第1/3周用这版——背阔肌外沿 视觉宽度" "45-50" $actsPullWidth

# 厚度优先版
$actsPullThick=@(
    @("杠铃俯身划船(主项)","4组","8-10次","RPE 7-8 厚度优先","90s","7-8","第2周用这版。第1/3周=宽度优先。","【要点】俯身45-60度 核心收紧 杠铃贴大腿前侧拉至肚脐。厚度=立体感——背阔肌的厚度和中背部。  这是厚度日——重量可以比宽度日重。"),
    @("T杠划船/单臂哑铃划船","3组","10-12次","RPE 7-8","75s","7-8","厚度日的辅助——不同角度划船","【要点】T杠=对握中性位 中背部厚度。单臂=左右独立 纠正不平衡。"),
    @("宽握高位下拉(轻量)","3组","15次","轻重量 RPE 6","60s","6","厚度日的宽度补充——轻重量保持宽度刺激","【要点】和宽度日同样动作 但轻重量。保持背阔肌外沿的刺激 不追求力竭。"),
    @("罗马尼亚硬拉","4组","8-10次","RPE 7-8","90s","7-8","第2周+5kg","【要点】同宽度日。RDL是拉日固定动作——不随宽/厚轮换。"),
    @("面拉(绳索)","3组","15次","轻重量","45s","6-7","—","【要点】同宽度日。"),
    @("哑铃弯举","3组","10-12次","RPE 7","45s","7","—","【要点】同宽度日——二头收尾。")
)
$r=Write-TrainingBlock $ws2 $r "拉日B：厚度优先（Thickness Priority）" "杠铃划船主项+T杠/单臂划船+宽握下拉(轻)+RDL+面拉+弯举  |  45-50" "第2周用这版——背阔肌厚度 立体感" "45-50" $actsPullThick

# 腿日 V2 45min限时
$actsLegs=@(
    @("杠铃深蹲(主项)","4组","8-10次","RPE 7-8 严格60s休息","60s","7-8","V2:休息从90s缩到60s——限时策略","【呼吸】瓦式——下蹲吸气屏住 站起呼气。【要点】深度至少平行地面。V2腿日核心=45分钟内完成。短间歇 中等重量 追求代谢压力。"),
    @("腿举+腿弯举超级组","A3组15-20次 B3组15-20次","交替 不休息","A RPE7-8 B RPE7","超级组内不休息 组间60s","7-8","V2新增超级组——代谢密度翻倍","【要点】A腿举做完直接做B腿弯举 不休息。A+B=1轮。3轮 每轮之间休息60s。股四头+腘绳肌同时充血——代谢压力拉满。"),
    @("站姿提踵","3组","15-20次","RPE 7 膝盖微屈","30s","7","—","【要点】顶峰停1秒 慢放。小腿=腿日收尾。"),
    @("悬垂举腿","3组","12-15次","自重","45s","7","—","【要点】下腹带动骨盆上卷。腿日的核心——在代谢疲劳状态下完成。"),
    @("腿日总时长","目标45分钟内完成","超时=皮质醇开始分解肌肉","—","—","—","如果超时：减少腿举或弯举的组数。宁可少做一组 不要超时。")
)
$r=Write-TrainingBlock $ws2 $r "周四  腿日（代谢 45min限时）" "深蹲+腿举/弯举超级组+提踵+举腿  |  45" "严格控制45分钟内完成——超时=得不偿失" "45" $actsLegs

# 上肢日 V2 预疲劳
$actsUpper=@(
    @("预疲劳超级组(先做!)","A侧平举3x15 + B三头下压3x15","交替不休息","轻-中重量","组内不休息 组间60s","7-8","V2新增预疲劳——先让侧肩和三头充血","【要点】A做完直接做B 1轮。3轮。目标：让三角肌中束和三头肌在复合动作前'预热'——突破'做完了但目标肌群还没感觉'的瓶颈。"),
    @("窄距卧推/地板卧推(主项)","4组","8-10次","RPE 7-8 三头主导","90s","7-8","V2:从哑铃卧推改为窄距/地板卧推——三头深度刺激","【要点】握距比肩窄 肘贴身体。地板卧推=离地30cm启动 三头全程张力。预疲劳后三头已经充血——这时候窄距卧推 三头会第一个力竭。"),
    @("高位下拉(宽握)","4组","10-12次","RPE 7-8","75s","7-8","第2周+1片","【要点】和拉日同动作——上肢日保持背阔肌刺激。"),
    @("面拉(绳索)","3组","15次","轻重量","45s","6-7","—","【要点】肩后束+肩袖健康——上肢日的肩部保护。"),
    @("农夫行走","3组","30-40m","大重量哑铃 单边30-40kg","60s","7-8","—","【要点】躯干正直不侧倾 核心全程收紧。上肢日收尾——握力+斜方+功能性。")
)
$r=Write-TrainingBlock $ws2 $r "周五  上肢日（预疲劳法 Pre-Fatigue）" "预疲劳超级组+窄距卧推+下拉+面拉+农夫行走  |  40-45" "先孤立充血  再复合深度刺激——突破瓶颈" "40-45" $actsUpper

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20; $ws2.Range("C:C").ColumnWidth=6
$ws2.Range("D:D").ColumnWidth=10; $ws2.Range("E:E").ColumnWidth=16; $ws2.Range("F:F").ColumnWidth=10
$ws2.Range("G:G").ColumnWidth=8; $ws2.Range("H:H").ColumnWidth=20; $ws2.Range("I:I").ColumnWidth=52
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 渐进+突破
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="渐进突破"

SetC $ws3 1 1 "渐进负荷  器械上限突破  进阶判断" $true 16 $headerBlue 0 $true
MC $ws3 1 1 9; $ws3.Range("A1:I1").RowHeight=34

$r=3
SetC $ws3 $r 1 "4周渐进（DUP框架内 每周4天主题不变 拉日版本轮换）" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "推日(AMRAP)" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "拉日版本" $true 10 $white $headerBlue $false
SetC $ws3 $r 5 "腿日(45min)" $true 10 $white $headerBlue $false
SetC $ws3 $r 7 "上肢日(预疲劳)" $true 10 $white $headerBlue $false; $r++

$prog=@(
    @("1","80% 4x5+AMRAP 记录次数","宽度优先","深蹲4x8-10 60s休息 超级组标准","预疲劳+窄距卧推 建基准"),
    @("2","+2.5-5kg 保持4x5 AMRAP次数  超越第1周","厚度优先","维持重量 休息缩至50s","窄距卧推+1-2kg"),
    @("3","再+2.5-5kg AMRAP冲击新次数","宽度优先 加重","维持 超级组再加1轮","窄距卧推再+1-2kg"),
    @("4","60% 3x5 换俯卧撑","弹力带+自重 换动作","高脚杯深蹲 自重 2x10","轻量 弹力带 技术优先")
)
foreach($pw in $prog){
    SetC $ws3 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pw[1] $false 9 0 0 $true; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $pw[2] $false 9 0 0 $true
    SetC $ws3 $r 5 $pw[3] $false 9 0 0 $true; MC $ws3 $r 5 7
    SetC $ws3 $r 7 $pw[4] $false 9 0 0 $true; MC $ws3 $r 7 9
    $ws3.Range("A$($r):I$($r)").RowHeight=34; SB $ws3 $r 1 9; $r++
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
    @("1.5次法","下放 半起 全下 全起=1次","推日/腿日——双倍刺激"),
    @("缩短休息","90s 60s 45s","腿日代谢——核心策略"),
    @("预疲劳","孤立力竭 复合动作","上肢日——本方案已用")
)
foreach($b in $bt){
    SetC $ws3 $r 1 $b[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $b[1] $false 10 0 0 $true; MC $ws3 $r 2 4
    SetC $ws3 $r 5 $b[2] $false 10 0 0 $true; MC $ws3 $r 5 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

# Rules
$r+=2
SetC $ws3 $r 1 "进阶判断规则——V2版" $true 14 $white $darkBg $true
MC $ws3 $r 1 9; $ws3.Range("A$($r):I$($r)").RowHeight=26
$ws3.Range("A$($r):I$($r)").Font.Color=$white; $ws3.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "信号" $true 10 $white $headerBlue $false
SetC $ws3 $r 3 "做法" $true 10 $white $headerBlue $false; MC $ws3 $r 3 9; $r++

$rules=@(
    @("AMRAP次数比上周+2次以上","下周推日重量+2.5-5kg——你在变强"),
    @("AMRAP次数没变或下降","维持重量——今天的你不是上周的你 正常。不退步。"),
    @("腿日45分钟内没完成","减少超级组轮数而非牺牲动作质量——宁可做少 不要做差"),
    @("拉日宽度和厚度哪个进步更快","进步快的保持 进步慢的下周期多轮一次——用数据调整"),
    @("预疲劳后窄距卧推重量比预期轻","正常——预疲劳的目的就是让目标肌群先力竭。重量不是指标 泵感才是。"),
    @("减载周感觉'太轻松'","这是对的——减载周不轻松=前三周没练到位"),
    @("连续2周睡不好+训练状态差","提前进入减载周——你是教练 知道过度训练的代价")
)
foreach($r3 in $rules){
    SetC $ws3 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 3 $r3[1] $false 10 0 0 $true; MC $ws3 $r 3 9
    $ws3.Range("A$($r):I$($r)").RowHeight=26; SB $ws3 $r 1 9; $r++
}

$ws3.Range("A:A").ColumnWidth=30; $ws3.Range("B:I").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 恢复策略 V2
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="恢复策略"

SetC $ws4 1 1 "热身  主动恢复  减载(换动作)  季度全休" $true 16 $headerBlue 0 $true
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

# Active recovery
$r+=2
SetC $ws4 $r 1 "休息日主动恢复（周三/六/日 选1-2天）" $true 14 $white $softBlue $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softBlue
$r++

$active=@(
    "你在健身房教课=那不是休息。真正的休息日：泡沫轴10-15min+低强度有氧20-30min+拉伸全身。",
    "低强度有氧=散步/游泳/休闲骑 心率<120bpm。高强度不行——那是训练不是恢复。",
    "完全不练也比'轻重量练一下'好——休息日的目标是恢复 不是'顺便练练'。"
)
foreach($ar in $active){
    SetC $ws4 $r 1 $ar $false 10 0 0 $true; MC $ws4 $r 1 9
    $ws4.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# Deload V2 (换动作)
$r+=2
SetC $ws4 $r 1 "第4周减载——V2:换动作 不是做轻一点" $true 14 $white $softPurple $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$softPurple
$r++

SetC $ws4 $r 1 "同样动作做轻量=关节应力路径相同 不是真正的休息。换动作=给关节不同应力分布 真正的休息。" $false 10 $headerBlue $lightGray $true
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

# Quarterly full rest V2
$r+=2
SetC $ws4 $r 1 "季度全休周——每3个月1次  V2新增" $true 14 $white $accentOrange $true
MC $ws4 $r 1 9; $ws4.Range("A$($r):I$($r)").RowHeight=26
$ws4.Range("A$($r):I$($r)").Font.Color=$white; $ws4.Range("A$($r):I$($r)").Interior.Color=$accentOrange
$r++

$quarter=@(
    "即使有减载周 连续3个月DUP训练仍会累积系统性疲劳。你每天教课=恢复压力比普通人大一倍。",
    "全休周做什么：散步/游泳/拉伸/泡沫轴。不碰器械 不练力量。",
    "这不是偷懒——是职业保护。长期坚持训练的人 是那些懂得何时休息的人。",
    "全休周结束后 身体和神经系统都'渴望'训练——这就是新周期的最佳起点。"
)
foreach($q in $quarter){
    SetC $ws4 $r 1 $q $false 10 0 0 $true; MC $ws4 $r 1 9
    $ws4.Range("A$($r):I$($r)").RowHeight=22; $r++
}

$ws4.Range("A:A").ColumnWidth=22; $ws4.Range("B:I").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 饮食+调节 V2
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="饮食与调节"

SetC $ws5 1 1 "增肌营养  训练窗口  自主调节框架" $true 16 $headerBlue 0 $true
MC $ws5 1 1 9; $ws5.Range("A1:I1").RowHeight=34

$r=3
SetC $ws5 $r 1 "训练前后营养窗口——V2精确化" $true 14 $white $accentGreen $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$accentGreen
$r++

SetC $ws5 $r 1 "你每天在健身房教课 身体一直处于低强度消耗状态。如果不注意训练前后营养窗口 可能会陷入'练了但没恢复'的循环。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=24; $r++

SetC $ws5 $r 1 "时机" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "内容" $true 10 $white $headerBlue $false; MC $ws5 $r 2 4
SetC $ws5 $r 5 "逻辑" $true 10 $white $headerBlue $false; MC $ws5 $r 5 9; $r++

$peri=@(
    @("训练前2小时","碳水40-50g+蛋白20-25g`n例：香蕉+蛋白粉+全麦吐司2片","提供训练能量 防止训练中分解肌肉。你教课消耗大 练前不吃=练到一半没力气。"),
    @("训练中","水+电解质 500-750ml","维持表现 防抽筋——尤其腿日和推日。"),
    @("训练后1小时内","碳水50-60g+蛋白30-35g`n例：米饭一碗+鸡胸肉200g`n或：蛋白粉2勺+香蕉+牛奶","补充糖原 启动肌肉修复。训练后1h=蛋白质吸收效率最高时段。"),
    @("教课日(非训练日)","每4-5小时吃一顿含蛋白质的餐","你教课也在消耗——不要等到训练日才吃够。保持蛋白质摄入稳定。"),
    @("休息日","碳水比训练日少20-30%","不练不需要那么多能量——自然少一点。蛋白质不变。")
)
foreach($pw in $peri){
    SetC $ws5 $r 1 $pw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pw[1] $false 10 0 0 $true; MC $ws5 $r 2 4
    SetC $ws5 $r 5 $pw[2] $false 10 0 0 $true; MC $ws5 $r 5 9
    $ws5.Range("A$($r):I$($r)").RowHeight=40; SB $ws5 $r 1 9; $r++
}

# Total
$r++
SetC $ws5 $r 1 "全天营养方向" $true 12 $headerBlue 0 $true; MC $ws5 $r 1 9; $r++

SetC $ws5 $r 1 "总热量2700-2900 蛋白质160-180g(2.0-2.2g/kg) 碳水320-360g 脂肪60-70g 水3-4L。干净增肌——不做脏增肌。你教一天课 碳水是真正的能量来源 不要害怕吃碳水。" $false 10 0 0 $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26; $r++

# Self-regulation V2
$r+=2
SetC $ws5 $r 1 "自主调节框架 V2——这套可以带走" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

$selfReg=@(
    "你是教练——下面这套逻辑不只适用于你自己 也可以迁移到给会员做计划。",
    "1. 定分化：每周几天→推拉腿/上下肢/全身。你4天=推+拉(宽/厚)+腿(代谢)+上肢(预疲劳)。",
    "2. 定DUP主题：大重量/容量/代谢/弱点。每天不同刺激。",
    "3. 定周期：3周渐进+1周减载(换动作)+季度全休=长期可持续。",
    "4. 定RPE：百分比是理论值 RPE是当天真实状态。",
    "5. 每月复盘：AMRAP次数+力量数据+体重+照片→决定下周期方向。",
    "这套框架你在西宁任何一个会员身上都能用——不依赖器械 不依赖水平。"
)
foreach($sr in $selfReg){
    SetC $ws5 $r 1 $sr $false 10 0 0 $true; MC $ws5 $r 1 9
    $ws5.Range("A$($r):I$($r)").RowHeight=20; $r++
}

# Issues
$r+=2
SetC $ws5 $r 1 "常见问题和判断" $true 14 $white $darkBg $true
MC $ws5 $r 1 9; $ws5.Range("A$($r):I$($r)").RowHeight=26
$ws5.Range("A$($r):I$($r)").Font.Color=$white; $ws5.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "遇到什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 3 "怎么做" $true 10 $white $headerBlue $false; MC $ws5 $r 3 9; $r++

$issues=@(
    @("教课站一天 累 不想练","降级：热身+主项做3组(不AMRAP)+拉伸。保持习惯 不追求强度。"),
    @("体重2周没涨","加碳水——每天多一碗饭。干净增肌慢是正常的。"),
    @("AMRAP次数连续2周卡在同样数字","换变式或提前减载。可能累积疲劳了。"),
    @("腿日45分钟没完成","减超级组轮数——宁可做少 不要超时。超时=皮质醇开始分解肌肉。"),
    @("预疲劳后主项重量比预期低","预疲劳的目的就是这个——重量不是指标 泵感才是。"),
    @("季度全休周到期 身体渴望训练","进入新周期——从第1周重新开始 但重量以全休前第3周的70-75%起步。")
)
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

SetC $ws6 1 1 "精简版  4周复盘  季度全休" $true 16 $headerBlue 0 $true
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

# Review V2
$r+=2
SetC $ws6 $r 1 "4周后复盘——用数据说话" $true 14 $white $darkBg $true
MC $ws6 $r 1 9; $ws6.Range("A$($r):I$($r)").RowHeight=26
$ws6.Range("A$($r):I$($r)").Font.Color=$white; $ws6.Range("A$($r):I$($r)").Interior.Color=$darkBg
$r++

SetC $ws6 $r 1 "V2看什么" $true 10 $white $headerBlue $false
SetC $ws6 $r 2 "怎么评判" $true 10 $white $headerBlue $false; MC $ws6 $r 2 6
SetC $ws6 $r 7 "下周期调整" $true 10 $white $headerBlue $false; MC $ws6 $r 7 9; $r++

$review=@(
    @("AMRAP次数","第1周vs第3周——涨了多少？","涨=加重。平/降=维持或换变式。"),
    @("深蹲8-10RM","第1周vs第3周","涨=加重。没涨=检查恢复。"),
    @("体重趋势","0.5-1kg/月净肌肉=理想。涨超1kg/周=控热量。","—"),
    @("拉日宽度vs厚度","哪个进步更快？","进步快的保持 慢的下周期多轮一次。"),
    @("照片","第1天vs第28天——镜子比体重秤真实","—"),
    @("季度全休","第3个月末 该全休了","全休1周 然后进入新季度周期")
)
foreach($rv in $review){
    SetC $ws6 $r 1 $rv[0] $true 10 $headerBlue $lightGray $false
    SetC $ws6 $r 2 $rv[1] $false 10 0 0 $true; MC $ws6 $r 2 6
    SetC $ws6 $r 7 $rv[2] $false 10 0 0 $true; MC $ws6 $r 7 9
    $ws6.Range("A$($r):I$($r)").RowHeight=26; SB $ws6 $r 1 9; $r++
}

$ws6.Range("A:A").ColumnWidth=26; $ws6.Range("B:I").ColumnWidth=18
Write-Host "Sheet 6 done"

# Save
$savePath="D:\Codex\members\吉\吉_增肌方案_V2.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
