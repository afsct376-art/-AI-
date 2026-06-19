$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# Colors
$darkBg = 0x1a1a2e; $gold = 0xd4a574; $white = 0xffffff
$lightGray = 0xf5f5f5; $headerBlue = 0x2c3e50; $accentGreen = 0x27ae60
$accentRed = 0xe74c3c; $accentOrange = 0xf39c12; $lightBlue = 0xd6eaf8
$softPurple = 0x8e44ad; $softPink = 0xe91e63; $warmPink = 0xff6b81

function Set-Cell($ws,$r,$c,$v,$b,$s,$cl,$bg,$wr) {
    $cell=$ws.Cells.Item($r,$c);$cell.Value="$v";$cell.Font.Bold=$b
    $cell.Font.Size=$s;if($cl){$cell.Font.Color=$cl}
    if($bg){$cell.Interior.Color=$bg};$cell.WrapText=$wr
}
function MC($ws,$r,$c1,$c2){$ws.Range($ws.Cells.Item($r,$c1),$ws.Cells.Item($r,$c2)).Merge()}
function SB($ws,$r,$c1,$c2){$rng=$ws.Range($ws.Cells.Item($r,$c1),$ws.Cells.Item($r,$c2));$rng.Borders.LineStyle=1;$rng.Borders.Weight=2}

function Write-TrainingBlock($ws,$ref,$dayLabel,$focus,$tips,$estTime,$acts){
    Set-Cell $ws $ref 1 $dayLabel $true 13 $white $darkBg $true
    MC $ws $ref 1 8
    $ws.Range("A$($ref):H$($ref)").RowHeight=26
    $ws.Range("A$($ref):H$($ref)").Font.Color=$white
    $ws.Range("A$($ref):H$($ref)").Interior.Color=$darkBg
    $ref++
    Set-Cell $ws $ref 1 "$focus  |  $estTime 分钟  |  $tips" $false 10 $headerBlue $lightGray $true
    MC $ws $ref 1 8; $ref++
    Set-Cell $ws $ref 1 "序" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 2 "动作" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 3 "组数" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 4 "次数" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 5 "重量建议" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 6 "休息" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 7 "感觉" $true 10 $white $headerBlue $false
    Set-Cell $ws $ref 8 "动作要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):H$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        Set-Cell $ws $ref 1 $i $false 10 0 $lightGray $false
        Set-Cell $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 6;$j++){Set-Cell $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 6)}
        $ws.Range("A$($ref):H$($ref)").RowHeight=58
        SB $ws $ref 1 8; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup done"

# ============================================
# SHEET 0: 封面
# ============================================
$ws0=$wb.Worksheets.Item(1); $ws0.Name=" 从这里开始"

Set-Cell $ws0 1 1 "王丽颖  减脂塑形方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5
$ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold
$ws0.Range("A1:E1").Interior.Color=$darkBg

Set-Cell $ws0 2 1 "崔知行教练  2026-06-18  |  175cm / 97kg  |  减脂+塑形，从第一步开始" $false 11 $white $darkBg $true
MC $ws0 2 1 5
$ws0.Range("A2:E2").Font.Color=$white
$ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
Set-Cell $ws0 $row 1 "  第一周——只做这三件事" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

Set-Cell $ws0 $row 1 "我不是让你立刻泡在健身房。第一周，只做这三件事。做到任意两件，就是成功的第一周。" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$row++

$firstWeek=@(
    @("1","每天8000步","上班路上多走一站、午休溜达一圈、晚饭后散步20分钟。不要求跑步——走就行。"),
    @("2","每天3L水","买一个大水壶（1.5L），上午喝完一壶、下午喝完一壶。戒掉含糖饮料——这是你第一周掉秤的最大来源。"),
    @("3","每天睡够7小时","美容院上班站着多，下班累。睡眠是脂肪燃烧的底层——睡不够，减脂效率打五折。")
)
foreach($fw in $firstWeek){
    Set-Cell $ws0 $row 1 $fw[0] $true 18 $accentGreen 0 $false
    Set-Cell $ws0 $row 2 $fw[1] $true 12 $headerBlue 0 $false
    Set-Cell $ws0 $row 3 $fw[2] $false 11 0 0 $true
    MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=44
    SB $ws0 $row 1 5
    $row++
}

$row++
Set-Cell $ws0 $row 1 "  你可能会问" $true 14 $white $darkBg $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$faq=@(
    @("Q: 不用练吗？","A: 第一周先建立饮食和步行的习惯。这两件事做好了，体重就会开始动。第二周再走进健身房。相信我——急不来，但不用急。"),
    @("Q: 为什么是8000步而不是10000？","A: 8000步是你现在能做到的。10000步可能第一天就劝退。做到了再加。每次只提一个自己能完成的目标。"),
    @("Q: 饮料真的一口都不能喝？","A: 第一周先试试换：奶茶换无糖茶、可乐换气泡水+柠檬。不用一步到位戒掉——先减半。"),
    @("Q: 我要是第一周就做不到怎么办？","A: 做到一件也算赢。8000步没走到？走5000步也比不动强。水没喝够3L？喝2L也比之前多了。第一天没做到第二天补。不追求完美——追求持续。")
)
foreach($f in $faq){
    Set-Cell $ws0 $row 1 $f[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $f[1] $false 10 0 0 $true
    MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=42
    SB $ws0 $row 1 5
    $row++
}

# V2: 8000步拆解
$row++
Set-Cell $ws0 $row 1 "  8000步拆解——从哪里来（示例）" $true 14 $white $darkBg $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$stepBreakdown=@(
    @("早上通勤多走一站公交/地铁","+1500步","早出门10分钟 提前一站下车走过去"),
    @("午休溜达10分钟","+2000步","吃完午饭别马上坐 绕着美容院/商圈走一圈"),
    @("下午接水/上厕所多走两趟","+1000步","每次绕远路 多走几步攒起来"),
    @("晚饭后散步20分钟","+2500步","吃完饭别马上坐下——出去走 顺便完成一天目标"),
    @("晚上在家走动/收拾","+1000步","做家务也是步数 擦桌子 收拾衣服 都算"),
    @("全天合计","~8000步  如果今天实在走不够——走5000步也比不走强")
)
Set-Cell $ws0 $row 1 "场景" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "步数" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "怎么做" $true 10 $white $headerBlue $false
MC $ws0 $row 3 5; $row++
foreach($sb in $stepBreakdown){
    Set-Cell $ws0 $row 1 $sb[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $sb[1] $false 10 0 0 $false
    Set-Cell $ws0 $row 3 $sb[2] $false 10 0 0 $true
    MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=28
    SB $ws0 $row 1 5; $row++
}

# V2: 每周只加一件事路径
$row++
Set-Cell $ws0 $row 1 "  如果同时做四件事太难——试试这个节奏" $true 14 $white $softPurple $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$weeklyFocus=@(
    @("第1周","只做一件事：戒含糖饮料","其他三件事能做就做 做不到不批评自己"),
    @("第2周","戒含糖饮料 + 每天8000步","第一个习惯稳了 再加一个"),
    @("第3周","戒含糖饮料 + 8000步 + 3L水","前两个已经不用'坚持'了"),
    @("第4周","四件事全做到","这时候你已经是一个'有运动习惯的人'了")
)
Set-Cell $ws0 $row 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "只加一件事" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "说明" $true 10 $white $headerBlue $false
MC $ws0 $row 3 5; $row++
foreach($wf in $weeklyFocus){
    Set-Cell $ws0 $row 1 $wf[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $wf[1] $false 10 0 0 $false
    Set-Cell $ws0 $row 3 $wf[2] $false 10 0 0 $true
    MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26
    SB $ws0 $row 1 5; $row++
}
Set-Cell $ws0 $row 1 "  慢就是快。一个月养成一个习惯，比逼自己一个月做到四件事然后放弃——效果好10倍。" $false 10 $softPurple 0 $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

$row+=2
Set-Cell $ws0 $row 1 "  四周总览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

Set-Cell $ws0 $row 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 2 "主题" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 3 "训练" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 4 "有氧" $true 10 $white $headerBlue $false
Set-Cell $ws0 $row 5 "饮食重点" $true 10 $white $headerBlue $false
$row++

$overview4=@(
    @("第1周","唤醒身体","不训练——只建立习惯","每天8000步","戒含糖饮料 每天3L水 睡够7h"),
    @("第2周","走进健身房","3次/周 全身训练(固定器械)","每天8000步+训练后多走10min","三餐规律 蛋白质每餐有 晚餐控碳水"),
    @("第3周","开始适应","4次/周 上下肢分化","每天8000-10000步","每餐一拳主食 一掌蛋白 蔬菜不限"),
    @("第4周","看到变化","4次/周 增加难度","每天10000步","保持 复盘 拍一张全身照对比")
)
foreach($o4 in $overview4){
    Set-Cell $ws0 $row 1 $o4[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $o4[1] $true 10 $white $darkBg $false
    Set-Cell $ws0 $row 3 $o4[2] $false 10 0 0 $false
    Set-Cell $ws0 $row 4 $o4[3] $false 10 0 0 $false
    Set-Cell $ws0 $row 5 $o4[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=30
    SB $ws0 $row 1 5
    $row++
}

$row+=2
Set-Cell $ws0 $row 1 "  其他内容快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

# V2: 照片+外食 索引更新
$idx=@(
    @("方案逻辑+你想改善的部位分析"," Sheet 方案概览"),
    @("第2-4周训练(含怎么选重量)"," Sheet 第2-4周训练"),
    @("热身+步数拆解+拍照标准"," Sheet 热身有氧拉伸"),
    @("饮食+外食/PMS/经期攻略"," Sheet 饮食与规则"),
    @("状态差/来例假/太累了 兜底"," Sheet 备用与复盘"),
    @("4周后怎么看效果 下一步"," Sheet 备用与复盘")
)
foreach($ix in $idx){
    Set-Cell $ws0 $row 1 $ix[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws0 $row 2 $ix[1] $false 10 0 0 $false
    MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24
    SB $ws0 $row 1 5
    $row++
}

# V3: 第一周每日清单 + 兜底话
$row+=2
Set-Cell $ws0 $row 1 "  第一周每日清单——每天照着打勾就行" $true 14 $white $accentGreen $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$checklist=@(
    @("1","早上起床喝一杯水（500ml）"),
    @("2","早餐：2个蛋+牛奶+一片全麦吐司/一小碗燕麦"),
    @("3","白天：水壶放旁边，上午一壶下午一壶，喝够3L"),
    @("4","午餐：米饭一拳+一份肉/鱼+蔬菜多夹"),
    @("5","不喝含糖饮料——奶茶换无糖茶，可乐换气泡水+柠檬"),
    @("6","步数攒到8000步（翻上去看拆解——碎片时间凑）"),
    @("7","晚餐：粗粮一拳+一份肉/豆腐+大量蔬菜"),
    @("8","晚饭后散步20分钟——最重要的一段步数来源"),
    @("9","睡前不吃了，饿了喝温水或吃黄瓜/小番茄"),
    @("10","睡够7小时——定好闹钟，手机放远一点")
)
foreach($cl in $checklist){
    Set-Cell $ws0 $row 1 $cl[0] $true 10 $accentGreen 0 $false
    Set-Cell $ws0 $row 2 $cl[1] $false 10 0 0 $true
    MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=24
    $row++
}
$row++
Set-Cell $ws0 $row 1 "  做到任意6项——今天就是成功的一天。做不到6项也没关系——明天继续。" $false 10 $accentGreen 0 $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=22
$row++

$row+=2
Set-Cell $ws0 $row 1 "  如果第一周完全没做到——一件都没做到" $true 14 $white $softPink $true
MC $ws0 $row 1 5
$ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white
$ws0.Range("A$($row):E$($row)").Interior.Color=$softPink
$row++

$safetyNet=@(
    "没关系。这不是'失败'——是'还没开始'。",
    "明天重新来。不用等周一，不用等下个月。",
    "从'今天喝够3L水'开始，就这一件事。做到了，第一周就算你过。",
    "你不需要完美——你只需要持续。"
)
foreach($sn in $safetyNet){
    Set-Cell $ws0 $row 1 $sn $false 10 0 0 $true
    MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=20
    $row++
}

$ws0.Range("A:A").ColumnWidth=26
$ws0.Range("B:B").ColumnWidth=20
$ws0.Range("C:E").ColumnWidth=22

Write-Host "Cover done"

# ============================================
# SHEET 1: 方案概览
# ============================================
$ws=$wb.Worksheets.Add(); $ws.Name="方案概览"

Set-Cell $ws 1 1 "王丽颖  减脂塑形方案  概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 8
$ws.Range("A1:H1").RowHeight=36

Set-Cell $ws 2 1 "女  25-30岁  175cm / 97kg  BMI 31.7  零基础  目标：减脂+改善体型  商业健身房" $false 11 $gold $darkBg $true
MC $ws 2 1 8
$ws.Range("A2:H2").RowHeight=26
$ws.Range("A2:H2").Font.Color=$white
$ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
Set-Cell $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true
MC $ws $row 1 8

$ov=@(
    @("训练频率","第1周：不训练（只建立习惯）","第2周：3次/周 全身训练","第3-4周：4次/周 上下肢分化"),
    @("每次时长","总计50-60分钟","热身5min + 力量35-40min + 有氧10min + 拉伸5min"),
    @("训练方式","固定器械入门  徒手辅助  第3周起逐步引入自由重量","先建立自信和安全动作模式，再谈重量"),
    @("方案周期","4周为起点","4周后根据体重变化+主观感受+体能水平调整"),
    @("方案特点","零基础渐进式入门  先习惯后训练  固定器械安全感  饮食第一训练第二")
)
$r=5
foreach($o in $ov){
    Set-Cell $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 3 $o[1] $false 10 0 0 $true
    MC $ws $r 3 5
    Set-Cell $ws $r 6 $o[2] $false 10 0 0 $true
    MC $ws $r 6 8
    SB $ws $r 1 8
    $r++
}

$r++
Set-Cell $ws $r 1 "为什么这样设计（很重要，请看完）" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8
$r++

$design=@(
    "你之前办了健身卡基本没来过——不是你的问题，是一上来就要'练'这个门槛太高了。",
    "本方案第一周不让你训练。只做三件事：每天8000步、3L水、睡够7小时。这三件事做到了，体重就会开始动。",
    "为什么？因为对97kg的体重来说，饮食和日常活动量是减脂的绝对主力。训练是加速器，不是发动机。",
    "第二周开始走进健身房——用的都是固定器械。固定器械的好处：轨道固定、不容易做错、不容易受伤、一个人也敢做。",
    "第三周加码到每周4练，上下肢分开——因为身体已经开始适应了，可以承受更多。",
    "整个过程像是'慢慢点火'，不是'上来就冲刺'。你要的不是坚持一个月然后放弃——是养成习惯，一辈子不再胖回去。"
)
foreach($d in $design){
    Set-Cell $ws $r 1 $d $false 10 0 0 $true
    MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=22
    $r++
}

$r++
Set-Cell $ws $r 1 "四周节奏一览" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8
$r++

Set-Cell $ws $r 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws $r 2 "阶段" $true 10 $white $headerBlue $false
Set-Cell $ws $r 3 "训练日" $true 10 $white $headerBlue $false
Set-Cell $ws $r 5 "主题" $true 10 $white $headerBlue $false
Set-Cell $ws $r 7 "本周目标" $true 10 $white $headerBlue $false
$r++

$weeks4=@(
    @("1","唤醒","不训练","","建立习惯","建立三个习惯：每天8000步+3L水+7h睡眠。体重开始动。"),
    @("2","入门","每周3次","全身训练(固定器械)","走进健身房","第一次训练完成就是胜利。固定器械建立安全感。"),
    @("3","适应","每周4次","上下肢分化","感觉'我可以'","固定器械+少量徒手+逐步引入轻哑铃。身体开始紧致。"),
    @("4","看到变化","每周4次","上下肢分化+有氧","拍对比照","体能明显提升。拍一张全身照对比第一天。量体重和腰围。")
)
foreach($w4 in $weeks4){
    Set-Cell $ws $r 1 $w4[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws $r 2 $w4[1] $true 10 $white $darkBg $false
    Set-Cell $ws $r 3 $w4[2] $false 10 0 0 $false
    Set-Cell $ws $r 5 $w4[3] $false 10 0 0 $false
    Set-Cell $ws $r 7 $w4[4] $false 10 0 0 $true
    $ws.Range("A$($r):H$($r)").RowHeight=28
    SB $ws $r 1 8
    $r++
}

$r++
Set-Cell $ws $r 1 "关于你提到的那些想改善的部位" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8
$r++

$bodyGoals=@(
    "你想改善：臀凹陷、大小腿粗、妈妈臀、胯宽、背变薄。这些目标我全部记下来了。",
    "但有一个事实你需要先知道：没有'局部减脂'。你不能只瘦大腿不瘦脸，不能只瘦肚子不瘦胸。脂肪是全身一起掉的。",
    "好消息是——当你全身脂肪率下降（比如从35%降到28%），上面那些部位全都会变好看。",
    "臀凹陷——通过臀大肌训练（臀推、髋外展）增加肌肉量填充。",
    "大小腿粗——脂肪减少后腿自然变细。小腿粗也和水肿有关——每天3L水+控盐有帮助。",
    "背变薄——背阔肌训练让背部紧致，同时全身减脂让背部脂肪减少。",
    "到第4周结束时，这些部位不一定已经'完美'——但你会看到它们正在往好的方向走。"
)
foreach($bg in $bodyGoals){
    Set-Cell $ws $r 1 $bg $false 10 0 0 $true
    MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=22
    $r++
}

$r++
Set-Cell $ws $r 1 "三个核心原则" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8
$r++

$principles=@(
    @("1","先习惯后训练","第一周不进健身房。先建立走路+喝水+睡眠的习惯——这三件事做好了，体重就开始掉。"),
    @("2","完成比完美重要","8000步没走到？走5000步也比不走强。今天没练？明天补。不追求'完美的计划执行'，追求'持续的往前挪'。"),
    @("3","饮食是主力，训练是加速器","你现在97kg——减脂的80%靠吃。训练让你的身材紧致、有线条，但体重下降主要是吃出来的。")
)
foreach($p in $principles){
    Set-Cell $ws $r 1 $p[0] $false 10 0 0 $false
    Set-Cell $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws $r 3 $p[2] $false 10 0 0 $true
    MC $ws $r 3 8
    $ws.Range("A$($r):H$($r)").RowHeight=28
    SB $ws $r 1 8
    $r++
}

$ws.Range("A:A").ColumnWidth=30
$ws.Range("B:H").ColumnWidth=16

Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 第2-4周训练计划
# ============================================
$ws2=$wb.Worksheets.Add(); $ws2.Name="第2-4周训练"

Set-Cell $ws2 1 1 "第2-4周  训练计划" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8
$ws2.Range("A1:H1").RowHeight=34

Set-Cell $ws2 2 1 "第1周不训练（只做封面页三件事）。第2周走进健身房。第3-4周开始有规律训练。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# V2: 怎么选重量指南
$r=4
Set-Cell $ws2 $r 1 "  怎么选重量——每次训练前看这个" $true 14 $white $accentGreen $true
MC $ws2 $r 1 8
$ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white
$ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

$weightGuide=@(
    "不要被'最轻重量'四个字骗了——不同器械最轻片效果不同。坐姿腿举机最轻片对你可能还是太轻，坐姿推胸可能刚刚好。",
    "",
    "选重量三步法：",
    "1. 先挂最轻的片，做一组试试",
    "2. 如果15次做完'完全没感觉'  加一片",
    "3. 如果做到第12次就做不动了  减一片",
    "  理想状态：做完15次，最后3次有点累，但动作全程不变形",
    "",
    "记住：第2周的目标不是'练到累'——是学会每个器械怎么用、找到'哪块肌肉在发力'。重量可以轻，动作不能错。"
)
foreach($wg in $weightGuide){
    Set-Cell $ws2 $r 1 $wg $false 10 0 0 $true
    MC $ws2 $r 1 8
    $ws2.Range("A$($r):H$($r)").RowHeight=20
    if($wg -eq ""){ $ws2.Range("A$($r):H$($r)").RowHeight=10 }
    $r++
}

$r++
Set-Cell $ws2 $r 1 "第2周  全身训练  每周3次（如周一/三/五 或 周二/四/六）" $true 14 $white $darkBg $true
MC $ws2 $r 1 8
$ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white
$ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

Set-Cell $ws2 $r 1 "第2周目标：走进健身房就是胜利。固定器械轨道固定、安全、一个人也敢做。三个训练日做同样的内容——先熟悉动作，不急着变换。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8
$ws2.Range("A$($r):H$($r)").RowHeight=26
$r++

# Week 2 workout (same all 3 days)
$actsW2=@(
    @("坐姿推胸","3组","12-15次","先挂最轻片 按选重三步法","60秒","轻松-中等","【呼吸】推出呼气、收回吸气。【要点】后背贴紧靠背。推到前面不锁手肘。先找到'胸肌在用力'的感觉。"),
    @("坐姿划船","3组","12-15次","先挂最轻片 按选重三步法","60秒","轻松-中等","【呼吸】拉向身体呼气、还原吸气。【要点】身体不要后仰——想象肩胛骨往中间夹。拉到腹部位置。"),
    @("坐姿腿举","3组","12-15次","空机或最轻片","60秒","轻松-中等","【呼吸】推出去呼气、收回来吸气。【要点】脚踩实踏板与肩同宽。膝盖不要完全蹬直——留一点点弯曲保护关节。"),
    @("坐姿髋外展","3组","15-20次","轻重量","45秒","轻松","【呼吸】打开时呼气、收回吸气。【要点】身体贴紧靠背不要晃。感受臀外侧发力。臀凹陷的改善——这个动作长期做有效。"),
    @("高位下拉","3组","12-15次","最轻片","60秒","轻松","【呼吸】拉下来呼气、还原吸气。【要点】先沉肩膀再用手臂拉。杆拉到锁骨高度。背变薄的第一步。"),
    @("平板支撑","2组","20-30秒","自重","30秒","做到力竭前停","【要点】手肘正下方撑地，腰不要塌——想象有人要打你肚子。能做多久做多久，第一次20秒也很棒。")
)
$r=Write-TrainingBlock $ws2 $r "第2周训练（三个训练日都做这个）" "固定器械入门  每个动作找到'哪块肌肉在用力'  |  RPE 4-6（轻松到中等）" "动作质量优先于重量——不要和别人比" "38-42" $actsW2

# ===== 第3周: 4x/week Upper/Lower =====
$r++
Set-Cell $ws2 $r 1 "第3-4周  上下肢分化  每周4次（如周一/二/四/五 或 周一/三/四/六）" $true 14 $white $darkBg $true
MC $ws2 $r 1 8
$ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white
$ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

Set-Cell $ws2 $r 1 "第3周起按上下肢分开训练。下肢日+上肢日交替。重量比第2周重1-2片（仍用选重三步法判断）。第4周再尝试+1片或+1-2kg哑铃。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8
$ws2.Range("A$($r):H$($r)").RowHeight=26
$r++

# Lower A
$actsLA=@(
    @("坐姿腿举","3组","10-12次","比第2周重1-2片","75秒","中等偏强","【呼吸】蹬出呼气、回收吸气。【发力感】脚后跟发力蹬，大腿前侧和臀部一起用力。重量比第2周重一点点。"),
    @("坐姿髋外展","3组","15-20次","比第2周重1片","45秒","中等","【发力感】打开时臀外侧收紧保持1秒。这个动作是改善胯宽和臀凹陷的关键。"),
    @("坐姿腿屈伸","3组","12-15次","轻-中重量","60秒","中等","【呼吸】踢出去呼气、放下吸气。【要点】脚尖回勾，感受大腿前侧发力。膝盖不舒服的话降重量。"),
    @("哑铃臀桥","3组","12-15次","双手持一颗哑铃6-10kg放在髋部","60秒","中等","【呼吸】推起呼气、下放吸气。【要点】下巴微收。顶峰夹紧屁股1-2秒。改善臀凹陷和妈妈臀。"),
    @("坐姿腿弯举","3组","12-15次","轻重量","60秒","中等","【呼吸】勾起呼气、还原吸气。【要点】大腿后侧发力把小腿拉下来。改善大腿后侧线条。"),
    @("平板支撑","3组","30-45秒","自重","30秒","做到力竭前","比第2周多坚持5-10秒。")
)
$r=Write-TrainingBlock $ws2 $r "下肢日 A（第3-4周 每周2次）" "股四头肌 + 臀部 + 腘绳肌  |  改善臀凹陷/妈妈臀/大腿线条" "比第2周重1-2片。先保证动作质量。" "40-45" $actsLA

# Upper A
$actsUA=@(
    @("坐姿推胸","3组","10-12次","比第2周重1-2片","75秒","中等","【发力感】推到前面时胸肌向中间挤压保持1秒。背贴紧靠背。"),
    @("高位下拉","3组","10-12次","比第2周重1片","75秒","中等","【发力感】拉到锁骨时背阔肌收紧。肩胛骨先下沉再拉——不要用手臂硬拽。背薄的关键动作。"),
    @("坐姿推肩","3组","12-15次","最轻片起步","60秒","中等","【呼吸】推起呼气、下放吸气。【要点】下背贴紧靠背。推到上面不锁死手肘。肩部线条感。"),
    @("坐姿划船","3组","10-12次","比第2周重1片","60秒","中等","【发力感】肩胛骨向中间夹。改善圆肩驼背——让上半身体态更直。"),
    @("哑铃侧平举","3组","12-15次","最轻哑铃2-4kg","45秒","中等","【要点】不要耸肩——用肩膀外侧发力。手臂微屈。改善肩部线条。"),
    @("仰卧举腿","2组","10-15次","自重","45秒","中等","【要点】用下腹带动腿上举，不靠惯性摇摆。")
)
$r=Write-TrainingBlock $ws2 $r "上肢日 A（第3-4周 每周2次）" "胸 + 背 + 肩 + 手臂  |  改善背厚/圆肩/上半身体态" "所有固定器械 一个人安全" "40-45" $actsUA

# Training schedule
$r++
Set-Cell $ws2 $r 1 "第3-4周 训练排班建议" $true 12 $headerBlue 0 $true
MC $ws2 $r 1 8
$r++

Set-Cell $ws2 $r 1 "推荐排法" $true 10 $headerBlue $lightGray $false
Set-Cell $ws2 $r 2 "周一=下肢 / 周二=上肢 / 周四=下肢 / 周五=上肢 / 周三六日=休息+走路" $false 10 0 0 $true
MC $ws2 $r 2 8
$ws2.Range("A$($r):H$($r)").RowHeight=24
SB $ws2 $r 1 8
$r++

Set-Cell $ws2 $r 1 "备选排法" $true 10 $headerBlue $lightGray $false
Set-Cell $ws2 $r 2 "周一=下肢 / 周三=上肢 / 周四=下肢 / 周六=上肢——如果周一周二连练太累的话" $false 10 0 0 $true
MC $ws2 $r 2 8
$ws2.Range("A$($r):H$($r)").RowHeight=24
SB $ws2 $r 1 8
$r++

Set-Cell $ws2 $r 1 "如果这周只能练2-3次" $true 10 $headerBlue $lightGray $false
Set-Cell $ws2 $r 2 "第3周退回做第2周的全身体系。第4周优先保下肢。具体见Sheet备用与复盘。" $false 10 0 0 $true
MC $ws2 $r 2 8
$ws2.Range("A$($r):H$($r)").RowHeight=24
SB $ws2 $r 1 8

# Week 3→4 progression
$r+=2
Set-Cell $ws2 $r 1 "第3周→第4周 怎么加码" $true 12 $headerBlue 0 $true
MC $ws2 $r 1 8
$r++

$prog34=@(
    "第3周：先做到每次训练顺利完成。重量保守——感觉对了比数字重要。",
    "第4周：每个动作尝试比第3周重1片（插片器械）或+1-2kg（哑铃）。如果动作开始变形→退回第3周重量。",
    "有氧：第3周训练后走10分钟（跑步机坡度5-8% 速度4-5km/h）；第4周走到12-15分钟。",
    "第4周结束时：你觉得第2周刚进健身房时的自己'太轻松了'——这就是进步。"
)
foreach($p34 in $prog34){
    Set-Cell $ws2 $r 1 $p34 $false 10 0 0 $true
    MC $ws2 $r 1 8
    $ws2.Range("A$($r):H$($r)").RowHeight=22
    $r++
}

$ws2.Range("A:A").ColumnWidth=4
$ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=7
$ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=18
$ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=10
$ws2.Range("H:H").ColumnWidth=56

Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 有氧+热身+拉伸
# ============================================
$ws3=$wb.Worksheets.Add(); $ws3.Name="热身有氧拉伸"

Set-Cell $ws3 1 1 "有氧安排  训练前热身  训练后拉伸" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8
$ws3.Range("A1:H1").RowHeight=34

# Daily walking
$r=3
Set-Cell $ws3 $r 1 "每天步数目标（最重要！比训练还重要）" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white
$ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

Set-Cell $ws3 $r 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "每天步数" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 4 "怎么做到" $true 10 $white $headerBlue $false
$r++

$steps=@(
    @("第1周","8000步","上班多走一站公交/地铁。午休溜达10分钟。晚饭后散步20分钟。手机自带计步就够了。"),
    @("第2周","8000步+训练后走10分钟","保持8000步 训练日在跑步机上多走10分钟(坡度5-8% 速度4-5km/h)"),
    @("第3周","8000-10000步","在第2周基础上 周末多走一会儿 不需要一口气走完 碎片时间凑够"),
    @("第4周","10000步","保持 如果你第1周8000步觉得吃力 第4周做到10000步的时候你会惊讶的")
)
foreach($st in $steps){
    Set-Cell $ws3 $r 1 $st[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $st[1] $true 10 $headerBlue 0 $false
    Set-Cell $ws3 $r 4 $st[2] $false 10 0 0 $true
    MC $ws3 $r 4 8
    $ws3.Range("A$($r):H$($r)").RowHeight=36
    SB $ws3 $r 1 8
    $r++
}

# Warmup (simple for beginner)
$r+=2
Set-Cell $ws3 $r 1 "训练前热身（5分钟  必须做  尤其是零基础）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white
$ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

Set-Cell $ws3 $r 1 "步" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 2 "动作" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 5 "要点" $true 9 $white $headerBlue $false
MC $ws3 $r 5 8
$r++

$warmups=@(
    @("1","原地踏步+摆臂","1分钟","像走路一样 逐渐加快节奏 身体微微发热"),
    @("2","肩关节绕圈","向前10次 向后10次","手臂画大圈——把肩膀活动开 推胸和下拉前必做"),
    @("3","站姿摆腿(扶墙)","前后+左右 各10次/侧","大腿前后的肌肉拉开——腿举和腿弯举前的重要准备"),
    @("4","自重深蹲(扶椅背)","10次","扶着椅子做 不用蹲很深 膝盖能弯多少弯多少 找到'大腿在用力'的感觉"),
    @("5","第一个动作空机/最轻重量试做","1组x10次","用最轻的重量先做一遍——让身体知道'接下来要干这个了'")
)
foreach($w in $warmups){
    Set-Cell $ws3 $r 1 $w[0] $false 9 0 0 $false
    Set-Cell $ws3 $r 2 $w[1] $true 9 $headerBlue 0 $false
    Set-Cell $ws3 $r 4 $w[2] $false 9 0 0 $false
    Set-Cell $ws3 $r 5 $w[3] $false 9 0 0 $true
    MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=28
    SB $ws3 $r 1 8
    $r++
}

# Stretch
$r+=2
Set-Cell $ws3 $r 1 "训练后拉伸（5分钟  帮你第二天不那么酸痛）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white
$ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$stretches=@(
    @("1","胸大肌拉伸(扶墙或门框)","每侧30秒","推胸后的放松——手臂张开扶墙 身体往反方向转"),
    @("2","背阔肌拉伸(扶墙侧屈)","每侧30秒","下拉/划船后的放松——手扶墙 身体向一侧弯"),
    @("3","大腿前侧拉伸(站姿脚跟贴臀)","每侧30秒","扶墙做——腿举/腿屈伸后的放松"),
    @("4","大腿后侧拉伸(坐姿直腿前屈)","30秒","腿弯举后的放松——坐地上 伸直腿 手够脚尖"),
    @("5","臀大肌拉伸(跷二郎腿抱膝)","每侧30秒","髋外展/臀桥后的放松——坐椅子上 脚踝搭对侧膝盖"),
    @("6","深呼吸","10次慢吸慢呼","让心跳慢下来——告诉身体'训练结束了 可以开始恢复了'")
)
Set-Cell $ws3 $r 1 "序号" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 2 "动作" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
Set-Cell $ws3 $r 5 "要点" $true 9 $white $headerBlue $false
MC $ws3 $r 5 8
$r++
foreach($s in $stretches){
    Set-Cell $ws3 $r 1 $s[0] $false 9 0 0 $false
    Set-Cell $ws3 $r 2 $s[1] $true 9 $headerBlue 0 $false
    Set-Cell $ws3 $r 4 $s[2] $false 9 0 0 $false
    Set-Cell $ws3 $r 5 $s[3] $false 9 0 0 $true
    MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=32
    SB $ws3 $r 1 8
    $r++
}

$ws3.Range("A:A").ColumnWidth=8
$ws3.Range("B:B").ColumnWidth=24
$ws3.Range("C:C").ColumnWidth=12
$ws3.Range("D:D").ColumnWidth=12
$ws3.Range("E:H").ColumnWidth=18

Write-Host "Sheet 3 done"

# ============================================
# SHEET 3b: 拍照标准 (V2新增)
# ============================================
$r+=2
Set-Cell $ws3 $r 1 "  拍照标准——今天就拍 4周后同一位置同一光线同一衣服对比" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8
$ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white
$ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

Set-Cell $ws3 $r 1 "今天拍一组 4周后同一条件再拍一组。不要吸肚子！不要找角度！真实的才有对比价值。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $r++

Set-Cell $ws3 $r 1 "拍什么" $true 10 $white $headerBlue $false
Set-Cell $ws3 $r 2 "怎么拍" $true 10 $white $headerBlue $false; MC $ws3 $r 2 5
Set-Cell $ws3 $r 6 "要点" $true 10 $white $headerBlue $false; $r++

$photoGuide=@(
    @("正面照","自然站立 双手放两侧","手机与肚脐同高 全身入镜 不要侧身"),
    @("侧面照","侧身站立 手臂自然下垂","手机与腰同高 能看到肚子和背部的轮廓"),
    @("背面照","背对镜子 双手叉腰","手机与肩胛骨同高 能看到背部线条和臀型"),
    @("穿什么","运动内衣+紧身裤/短裤","露出腰腹才能看到变化。不要穿宽松T恤拍。"),
    @("什么时候拍","今天(第1天)  第4周最后一天","同一时间(建议早上空腹) 同一位置 同一光线")
)
foreach($pg in $photoGuide){
    Set-Cell $ws3 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws3 $r 2 $pg[1] $false 10 0 0 $false; MC $ws3 $r 2 5
    Set-Cell $ws3 $r 6 $pg[2] $false 10 0 0 $true
    $ws3.Range("A$($r):H$($r)").RowHeight=30
    SB $ws3 $r 1 8; $r++
}

# ============================================
# SHEET 4: 饮食+规则
# ============================================
$ws4=$wb.Worksheets.Add(); $ws4.Name="饮食与规则"

Set-Cell $ws4 1 1 "饮食方案  生活习惯调整  怎么判断该加量还是该休息" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8
$ws4.Range("A1:H1").RowHeight=34

# Nutrition
$r=3
Set-Cell $ws4 $r 1 "饮食方案（减脂的主力军——80%的效果靠这个）" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white
$ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

Set-Cell $ws4 $r 1 "你的现状：97kg，想快速减脂。以下数字是方向——不需要每天称克数。理解原则，生活化执行。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=24
$r++

Set-Cell $ws4 $r 1 "营养素" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "目标" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "怎么吃（不称克数版本）" $true 10 $white $headerBlue $false
$r++

$nutriWang=@(
    @("总热量","1600-1800 kcal","比你现在少500-700kcal——刚好够让体重每周掉0.5-1kg，但不至于饿得受不了。"),
    @("蛋白质","130-160g","每餐一份手掌大的蛋白质：早餐2个蛋+牛奶 午餐晚餐各有一份肉/鱼/豆腐。练后喝一杯蛋白粉的话更容易达标。"),
    @("碳水","160-180g","米饭每餐半拳到一拳。优先选杂粮饭/红薯/玉米。晚餐碳水可以比午餐少一半。水果一天1-2份就够了(一根香蕉或一个苹果的量)。"),
    @("脂肪","40-50g","炒菜油控一下。不吃油炸。坚果一天一小把就够了。"),
    @("饮水","3L/天","买一个1.5L大水壶 上午喝完一壶 下午喝完一壶。水里加点柠檬片/黄瓜片增加口感。戒掉所有含糖饮料。"),
    @("蔬菜","不限量","每餐至少占一半盘子。先吃蔬菜再吃肉再吃主食——这个顺序掉秤更快。")
)
foreach($nw in $nutriWang){
    Set-Cell $ws4 $r 1 $nw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $nw[1] $false 10 0 0 $false
    Set-Cell $ws4 $r 4 $nw[2] $false 10 0 0 $true
    MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=36
    SB $ws4 $r 1 8
    $r++
}

# Meal guide
$r++
Set-Cell $ws4 $r 1 "三餐示例（照着这个方向吃，不用精确复制）" $true 12 $headerBlue 0 $true
MC $ws4 $r 1 8
$r++

Set-Cell $ws4 $r 1 "餐次" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "吃什么" $true 10 $white $headerBlue $false
MC $ws4 $r 2 5
Set-Cell $ws4 $r 6 "要点" $true 10 $white $headerBlue $false
$r++

$mealsWang=@(
    @("早餐","2个鸡蛋(煮/蒸)+一杯牛奶(250ml)+一片全麦吐司/一小碗燕麦","蛋白质吃够上午不饿。鸡蛋别只吃蛋白——蛋黄里的营养对女性重要。"),
    @("午餐","米饭半拳+一份肉/鱼(手掌大)+蔬菜(占一半盘子)","食堂或外卖：选清蒸/炖/卤的菜 避开油炸和勾芡。先吃蔬菜 再吃肉 最后吃饭。"),
    @("加餐(下午)","一个小水果/一杯无糖酸奶/一把坚果(10颗以内)","如果午餐到晚餐之间饿了就吃 不饿可以不吃。不要饿到晚餐暴食。"),
    @("晚餐","米饭半拳或不吃主食+一份肉/鱼/豆腐+大量蔬菜","晚餐控碳水是掉秤的关键。晚上饿了吃黄瓜/番茄/喝温水。"),
    @("训练后","一杯蛋白粉+水(或牛奶)","训练完30分钟内喝。加上日常三餐的蛋白质，总量轻松达标。"),
    @("全天","戒含糖饮料。喝水3L。","奶茶换无糖茶/美式咖啡。可乐换气泡水+柠檬。这是你第一周掉秤的最大来源。")
)
foreach($mw in $mealsWang){
    Set-Cell $ws4 $r 1 $mw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $mw[1] $false 10 0 0 $true
    MC $ws4 $r 2 5
    Set-Cell $ws4 $r 6 $mw[2] $false 10 0 0 $true
    $ws4.Range("A$($r):H$($r)").RowHeight=38
    SB $ws4 $r 1 8
    $r++
}

# V2: 外食/外卖指南
$r+=2
Set-Cell $ws4 $r 1 "如果今天吃外卖/外食——这样选  V2新增" $true 14 $white $darkBg $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white
$ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

Set-Cell $ws4 $r 1 "不要求顿顿完美。一天里有一顿吃对了——就比昨天好。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $r++

Set-Cell $ws4 $r 1 "  推荐选这些" $true 12 $accentGreen 0 $true; MC $ws4 $r 1 8; $r++

$goodEats=@(
    @("沙县小吃","鸡腿饭(去皮)+烫青菜","蛋白质够 碳水适中 蔬菜有"),
    @("麻辣烫","清汤底 多蔬菜+豆制品+瘦肉 少丸子","丸子=淀粉+油 少拿"),
    @("轻食沙拉","酱汁另放 只蘸不浇","沙拉酱=热量炸弹 蘸着吃别浇上去"),
    @("便利店","鸡胸肉+茶叶蛋+无糖豆浆+玉米","方便便宜 一顿蛋白质够了"),
    @("潮汕牛肉/粿条","选牛肉+蔬菜 汤底少喝","汤底盐分高 喝几口就好")
)
foreach($ge in $goodEats){
    Set-Cell $ws4 $r 1 $ge[0] $true 10 $accentGreen $lightGray $false
    Set-Cell $ws4 $r 2 $ge[1] $false 10 0 0 $false
    Set-Cell $ws4 $r 4 $ge[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26
    SB $ws4 $r 1 8; $r++
}

$r++
Set-Cell $ws4 $r 1 "  尽量避开这些" $true 12 $accentRed 0 $true; MC $ws4 $r 1 8; $r++

$badEats=@(
    @("炒饭/炒面/盖浇饭","油+碳水双高——减脂期最大的隐形热量来源"),
    @("油炸类(炸鸡排/炸猪排/炸串)","裹粉吸油=热量翻倍"),
    @("含糖饮料(奶茶/果汁/可乐)","液体热量——不占肚子却400-600kcal"),
    @("勾芡/红烧/糖醋类","淀粉+油+糖三高——选清蒸/白切/卤/炖替代")
)
foreach($be in $badEats){
    Set-Cell $ws4 $r 1 $be[0] $true 10 $accentRed $lightGray $false
    Set-Cell $ws4 $r 2 $be[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24
    SB $ws4 $r 1 8; $r++
}

# V2: PMS应对
$r+=2
Set-Cell $ws4 $r 1 "经期前一周（PMS阶段）——提前知道就不慌  V2新增" $true 14 $white $softPink $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white
$ws4.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

Set-Cell $ws4 $r 1 "很多女性在经前一周食欲暴增、情绪波动、水肿——这时候最容易破功。提前告诉你：这是正常的激素波动，不是你意志力差。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

Set-Cell $ws4 $r 1 "你会遇到的" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "为什么" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 4 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$pms=@(
    @("体重涨1-3kg","正常水肿 不是胖了——激素导致水分滞留","别称体重 或少称。经期结束后水自然排掉。"),
    @("食欲暴增 特别想吃甜的","激素波动——不是你意志力差","可以吃 但选相对健康的：黑巧克力(70%+)2小块/水果/无糖酸奶+蜂蜜。"),
    @("情绪低落 不想训练","正常——经前雌激素下降影响情绪","训练强度降一级(重量减/组数减)。但尽量保持去健身房——哪怕只做热身+有氧。经前一周'维持'就是胜利。"),
    @("训练时觉得比平时累","身体在准备来月经 能量储备偏低","降重量。感觉比平时低1-2分。不要求'突破'——能完成就是胜利。")
)
foreach($pm in $pms){
    Set-Cell $ws4 $r 1 $pm[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $pm[1] $false 10 0 0 $false
    Set-Cell $ws4 $r 4 $pm[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=42
    SB $ws4 $r 1 8; $r++
}

$r++
Set-Cell $ws4 $r 1 "来例假时" $true 12 $accentRed 0 $true; MC $ws4 $r 1 8; $r++
Set-Cell $ws4 $r 1 "经期第1-2天休息或只走路。第3天后如果OK恢复轻量训练。经期训练不勉强——身体在流血需要休息，不是偷懒。" $false 10 0 0 $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

# 习惯重点
Set-Cell $ws4 $r 1 "这4周的生活习惯重点" $true 12 $headerBlue 0 $true; MC $ws4 $r 1 8; $r++

$habitsWang=@(
    @("第1周","戒含糖饮料 每天3L水 睡够7h","三件事做到任意两件=成功"),
    @("第2周","三餐规律 每餐有蛋白质 晚餐控碳水","走进健身房同时饮食跟上 不要求完美要求持续"),
    @("第3周","用'先吃菜再吃肉最后吃饭'的顺序","自然少吃饭——不用挨饿"),
    @("第4周","复盘：哪餐最难坚持？哪个习惯已经不需要想了？","庆祝养成的习惯 修补困难的环节")
)
Set-Cell $ws4 $r 1 "周" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 2 "重点" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 5 "说明" $true 10 $white $headerBlue $false
MC $ws4 $r 5 8; $r++
foreach($hw in $habitsWang){
    Set-Cell $ws4 $r 1 $hw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 2 $hw[1] $false 10 0 0 $false
    Set-Cell $ws4 $r 5 $hw[2] $false 10 0 0 $true
    MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=28
    SB $ws4 $r 1 8
    $r++
}

# Rules
$r+=2
Set-Cell $ws4 $r 1 "怎么判断下一步怎么走" $true 14 $white $darkBg $true
MC $ws4 $r 1 8
$ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white
$ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$rulesWang=@(
    @("这周训练太轻松","下周每个动作加1片(器械)或+1-2kg(哑铃)。先保证动作不变形。"),
    @("训练后48小时还酸痛","正常——零基础前两周就是会酸。多喝水+多睡+拉伸。关节疼不是肌肉疼——告诉我。"),
    @("这周体重没掉","先看是不是经期前后——女性经期前水肿1-3kg正常。不是的话 回头检查饮料和步数。"),
    @("经期前一周体重涨了","正常水肿！不是胖了！经期结束水就排掉。"),
    @("来例假了","第1-2天休息或只走路。第3天后如果OK恢复轻量训练。不勉强。"),
    @("太累了不想练","走5000步+拉伸——做了就算。但不要连续跳过2次。"),
    @("这周吃得不好/喝了奶茶","明天继续 别破罐破摔等周一。一顿吃多不会胖——连续一周才会。"),
    @("膝盖/腰/关节疼","立即停止那个动作 换另一个不疼的。疼超过2天告诉我。")
)
Set-Cell $ws4 $r 1 "遇到这种情况" $true 10 $white $headerBlue $false
Set-Cell $ws4 $r 3 "怎么办" $true 10 $white $headerBlue $false
MC $ws4 $r 3 8
$r++
foreach($rw in $rulesWang){
    Set-Cell $ws4 $r 1 $rw[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws4 $r 3 $rw[1] $false 10 0 0 $true
    MC $ws4 $r 3 8
    $ws4.Range("A$($r):H$($r)").RowHeight=30
    SB $ws4 $r 1 8
    $r++
}

$ws4.Range("A:A").ColumnWidth=26
$ws4.Range("B:H").ColumnWidth=18

Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 备用+复盘
# ============================================
$ws5=$wb.Worksheets.Add(); $ws5.Name="备用与复盘"

Set-Cell $ws5 1 1 "精简备用  4周后怎么看效果  下一步怎么走" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8
$ws5.Range("A1:H1").RowHeight=34

# Backup simplified plan
$r=3
Set-Cell $ws5 $r 1 "这周状态差/加班/事情多 怎么办（降级方案）" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white
$ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

Set-Cell $ws5 $r 1 "原则：宁可降级训练，不要完全不练。降级不丢人——连续放弃才是问题。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8
$r++

Set-Cell $ws5 $r 1 "本来计划" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "降级为" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 4 "最少完成这个" $true 10 $white $headerBlue $false
$r++

$fallback=@(
    @("4次/周训练","减到2-3次 保留下肢日","哪怕一周只练了1次 也比不练强"),
    @("今天该练但很累","热身5min+最轻重量做2-3个动作+拉伸","走进健身房待20分钟就是胜利"),
    @("今天完全不想出门","在家/楼下 走20分钟+做10分钟拉伸","也算运动了——走路永远不丢人"),
    @("这周饮食全乱了","明天重新开始","不要等下周一开始——明天早上就可以重来"),
    @("来例假前3天","只走路 不训练","身体在准备月经 需要休息——不是偷懒"),
    @("经前特别想吃甜的","吃！但选黑巧克力/水果/酸奶","不是意志力差——是激素。满足但不过量。")
)
foreach($fb in $fallback){
    Set-Cell $ws5 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $fb[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 4 $fb[2] $false 10 0 0 $true
    MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=28
    SB $ws5 $r 1 8
    $r++
}

# Review
$r+=2
Set-Cell $ws5 $r 1 "4周后 怎么知道自己有没有进步" $true 14 $white $darkBg $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white
$ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

Set-Cell $ws5 $r 1 "不要只看体重秤！以下每一条都是进步：" $false 11 $accentGreen 0 $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=24
$r++

$progress=@(
    @("体重","掉了1-4kg","周三早上空腹称。不要每天称——女性体重每天波动。经期前后不称——水肿骗你。"),
    @("腰围","小了2-5cm","软尺量肚脐一圈。比体重更真实——体重没变腰细了=脂肪换肌肉。"),
    @("照片","不一样了","翻出第一天的三张照片 同一位置同一光线再拍一组。照片不会骗人——你会看到的。"),
    @("衣服","松了","裤腰能塞进一个手指 上衣不绷了 手臂线条出来了——体重秤测不出。"),
    @("体力","变好了","从8000步气喘到10000步不累。从'完全不想动'到'今天没动有点不舒服'。"),
    @("自信","敢走进健身房了","第2周第一次进健身房可能紧张。第4周你已经知道每个器械怎么用了——你不再是一个'不运动的人'了。"),
    @("习惯","不用逼自己了","走路/喝水/不喝甜饮料——已经从'任务'变成'自然'。这个，才是永远不会反弹的保证。")
)
Set-Cell $ws5 $r 1 "看什么" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 2 "可能的变化" $true 10 $white $headerBlue $false
Set-Cell $ws5 $r 4 "怎么看" $true 10 $white $headerBlue $false
MC $ws5 $r 4 8
$r++
foreach($pg in $progress){
    Set-Cell $ws5 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    Set-Cell $ws5 $r 2 $pg[1] $false 10 0 0 $false
    Set-Cell $ws5 $r 4 $pg[2] $false 10 0 0 $true
    MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=36
    SB $ws5 $r 1 8
    $r++
}

# Next steps
$r+=2
Set-Cell $ws5 $r 1 "4周之后 下一步怎么走" $true 14 $white $darkBg $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white
$ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$nextSteps=@(
    "如果减了2-4kg 且体能明显变好：继续保持这套方案再走4周。第5-8周可以加一点重量 再多加一组。",
    "如果体重没动 但腰围小了 体力好了：恭喜你——脂肪正在被肌肉替代。体重没变但你已经不一样了。继续。",
    "如果完全没变化：回头检查饮食和步数——是不是饮料没完全戒掉？是不是步数没走到？先修这两个。",
    "如果你已经爱上运动了：告诉我。下一周期可以增加更多训练变化——但现在 先把这个周期走完。"
)
foreach($ns in $nextSteps){
    Set-Cell $ws5 $r 1 $ns $false 10 0 0 $true
    MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24
    $r++
}

# Self-training starter
$r+=2
Set-Cell $ws5 $r 1 "你学会自己做计划的第一步" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white
$ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

$selfTrain=@(
    "现在不需要你自己做计划——这套方案已经帮你做好了4周。",
    "但4周之后，你已经有感觉了：知道什么动作练哪里、知道什么感觉是对的、知道自己能完成什么。",
    "那时候你只需要记住三句话：",
    "  1. 每周至少练3次——每次选4-6个固定器械动作 每个3组12-15次",
    "  2. 每天走路8000-10000步——走到能一边走一边断续说话的速度",
    "  3. 戒糖饮 多喝水 每餐有蛋白质 晚餐控碳水",
    "这三句话就是你的'终身减脂方案'。不需要复杂的周期化 不需要精密的RPE——够简单才能做一辈子。"
)
foreach($st in $selfTrain){
    Set-Cell $ws5 $r 1 $st $false 10 0 0 $true
    MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=22
    $r++
}

$r++
Set-Cell $ws5 $r 1 "  最后的话：你之前问过私教，办了卡也没来——那不是你的问题。是之前的方案对你来说门槛太高了。现在这个方案，第一周不进健身房。做到了封面三件事，体重就会开始动。你离你想要的体型，只差一个'开始'。而这个开始——简单到不可能失败。" $false 10 $warmPink 0 $true
MC $ws5 $r 1 8
$ws5.Range("A$($r):H$($r)").RowHeight=42

$ws5.Range("A:A").ColumnWidth=26
$ws5.Range("B:H").ColumnWidth=20

Write-Host "Sheet 5 done"

# Save
$savePath="D:\Codex\王丽颖_减脂塑形方案_V2.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()
[GC]::Collect()
Write-Host "Saved: $savePath"
Write-Host "DONE"
