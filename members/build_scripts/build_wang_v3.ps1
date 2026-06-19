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
    SetC $ws $ref 1 $title $true 13 $white $darkBg $true; MC $ws $ref 1 8
    $ws.Range("A$($ref):H$($ref)").RowHeight=26
    $ws.Range("A$($ref):H$($ref)").Font.Color=$white
    $ws.Range("A$($ref):H$($ref)").Interior.Color=$darkBg; $ref++
    SetC $ws $ref 1 "$focus  |  $time 分钟  |  $tips" $false 10 $headerBlue $lightGray $true
    MC $ws $ref 1 8; $ref++
    SetC $ws $ref 1 "序" $true 10 $white $headerBlue $false
    SetC $ws $ref 2 "动作" $true 10 $white $headerBlue $false
    SetC $ws $ref 3 "组数" $true 10 $white $headerBlue $false
    SetC $ws $ref 4 "次数" $true 10 $white $headerBlue $false
    SetC $ws $ref 5 "重量怎么选" $true 10 $white $headerBlue $false
    SetC $ws $ref 6 "休息" $true 10 $white $headerBlue $false
    SetC $ws $ref 7 "感觉" $true 10 $white $headerBlue $false
    SetC $ws $ref 8 "动作要点" $true 10 $white $headerBlue $true
    $ws.Range("A$($ref):H$($ref)").RowHeight=26; $ref++
    $i=1
    foreach($a in $acts){
        SetC $ws $ref 1 $i $false 10 0 $lightGray $false
        SetC $ws $ref 2 $a[0] $true 10 $headerBlue 0 $false
        for($j=1;$j -le 6;$j++){SetC $ws $ref ($j+2) $a[$j] $false 9 0 0 ($j -eq 6)}
        $ws.Range("A$($ref):H$($ref)").RowHeight=58
        SB $ws $ref 1 8; $ref++; $i++
    }
    return $ref+2
}

Write-Host "Setup"

# ============================================
# SHEET 0:  从这里开始（封面）
# ============================================
$ws0=$wb.Worksheets.Item(1);$ws0.Name=" 从这里开始"

SetC $ws0 1 1 "王丽颖  减脂塑形方案" $true 22 $gold $darkBg $true
MC $ws0 1 1 5; $ws0.Range("A1:E1").RowHeight=46
$ws0.Range("A1:E1").Font.Color=$gold; $ws0.Range("A1:E1").Interior.Color=$darkBg

SetC $ws0 2 1 "崔知行教练  2026-06-18  |  175cm / 97kg  |  零基础入门  从第一步开始" $false 11 $white $darkBg $true
MC $ws0 2 1 5; $ws0.Range("A2:E2").Font.Color=$white; $ws0.Range("A2:E2").Interior.Color=$darkBg

$row=4
SetC $ws0 $row 1 "  第一周——只做这三件事" $true 16 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=32
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

SetC $ws0 $row 1 "不是让你立刻泡在健身房。第一周只做这三件事。做到任意两件——就是成功的一周。" $false 11 $lightGray 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=24; $row++

$firstWeek=@(
    @("1","每天8000步","拆开看一点都不多 翻到下面有具体拆解"),
    @("2","每天3L水","买一个1.5L大水壶 上午一壶下午一壶 戒掉含糖饮料"),
    @("3","每天睡够7小时","美容院上班站着累——睡眠是脂肪燃烧的底层")
)
foreach($fw in $firstWeek){
    SetC $ws0 $row 1 $fw[0] $true 18 $accentGreen 0 $false
    SetC $ws0 $row 2 $fw[1] $true 12 $headerBlue 0 $false
    SetC $ws0 $row 3 $fw[2] $false 11 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=42; SB $ws0 $row 1 5; $row++
}

# V3: 每日打卡清单
$row++
SetC $ws0 $row 1 "  第一周每日清单——每天照着打勾就行" $true 14 $white $accentGreen $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$accentGreen
$row++

$checklist=@(
    @("1","早上起床喝一杯水（500ml）"),
    @("2","早餐：2个蛋 + 牛奶 + 一片全麦吐司 / 一小碗燕麦"),
    @("3","白天：水壶放旁边，上午一壶下午一壶，喝够3L"),
    @("4","午餐：米饭一拳 + 一份肉/鱼 + 蔬菜多夹"),
    @("5","不喝含糖饮料——奶茶换无糖茶，可乐换气泡水+柠檬"),
    @("6","步数攒到8000步（翻下去看拆解——碎片时间凑）"),
    @("7","晚餐：粗粮一拳 + 一份肉/豆腐 + 大量蔬菜"),
    @("8","晚饭后散步20分钟——最重要的一段步数来源"),
    @("9","睡前不吃了，饿了喝温水或吃黄瓜/小番茄"),
    @("10","睡够7小时——定好闹钟，手机放远一点")
)
foreach($cl in $checklist){
    SetC $ws0 $row 1 $cl[0] $true 10 $accentGreen 0 $false
    SetC $ws0 $row 2 $cl[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}
$row++
SetC $ws0 $row 1 "  做到任意6项——今天就是成功的一天。做不到也没关系——明天继续。不需要完美，只需要持续。" $false 10 $accentGreen 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

$row+=2
SetC $ws0 $row 1 "  如果第一周完全没做到——一件都没做到" $true 14 $white $softPink $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPink
$row++

SetC $ws0 $row 1 "没关系。这不是'失败'——是'还没开始'。明天重新来。不用等周一，不用等下个月。从'今天喝够3L水'开始，就这一件事。做到了，第一周就算你过。你不需要完美——你只需要持续。" $false 10 0 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=38; $row++

# 8000步拆解
$row+=2
SetC $ws0 $row 1 "  8000步拆解——从哪里来（示例）" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$stepBreakdown=@(
    @("早上通勤多走一站公交/地铁","+1500步","早出门10分钟 提前一站下车走过去"),
    @("午休溜达10分钟","+2000步","吃完午饭别马上坐 绕着美容院/商圈走一圈"),
    @("下午接水/上厕所多走两趟","+1000步","每次绕远路 多走几步攒起来"),
    @("晚饭后散步20分钟","+2500步","吃完饭别马上坐下——出去走 顺便完成一天目标"),
    @("晚上在家走动/收拾","+1000步","做家务也是步数 擦桌子 收拾衣服 都算"),
    @("全天合计","~8000步  如果今天实在走不够——走5000步也比不走强")
)
SetC $ws0 $row 1 "场景" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "步数" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "怎么做" $true 10 $white $headerBlue $false; MC $ws0 $row 3 5; $row++
foreach($sb in $stepBreakdown){
    SetC $ws0 $row 1 $sb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $sb[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $sb[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

# 每周只加一件事
$row++
SetC $ws0 $row 1 "  如果同时做四件事太难——试试这个节奏" $true 14 $white $softPurple $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softPurple
$row++

$weeklyFocus=@(
    @("第1周","只做一件事：戒含糖饮料","其他三件事能做就做 做不到不批评自己"),
    @("第2周","戒含糖饮料 + 每天8000步","第一个习惯稳了 再加一个"),
    @("第3周","戒含糖饮料 + 8000步 + 3L水","前两个已经不用'坚持'了"),
    @("第4周","四件事全做到","这时候你已经是一个'有运动习惯的人'了")
)
SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "只加一件事" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "说明" $true 10 $white $headerBlue $false; MC $ws0 $row 3 5; $row++
foreach($wf in $weeklyFocus){
    SetC $ws0 $row 1 $wf[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $wf[1] $false 10 0 0 $false
    SetC $ws0 $row 3 $wf[2] $false 10 0 0 $true; MC $ws0 $row 3 5
    $ws0.Range("A$($row):E$($row)").RowHeight=26; SB $ws0 $row 1 5; $row++
}
SetC $ws0 $row 1 "  慢就是快。一个月养成一个习惯，比逼自己一个月做到四件事然后放弃——效果好10倍。" $false 10 $softPurple 0 $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++

# 你可能会问
$row++
SetC $ws0 $row 1 "  你可能会问" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$faq=@(
    @("Q: 不用练吗？","A: 第一周先建立饮食和步行的习惯。这两件做好了 体重就会动。第二周再走进健身房。"),
    @("Q: 8000步一下走不完怎么办？","A: 不用一口气走完。碎片时间凑够——上班前+午休+晚饭后+在家走动。手机计步自动帮你加。"),
    @("Q: 饮料一口都不能喝？","A: 先试试换：奶茶换无糖茶 可乐换气泡水+柠檬。先减半 不用一步到位戒掉。"),
    @("Q: 第一周就做不到怎么办？","A: 做到一件也算赢。第一天没做到第二天补。不追求完美——追求持续。也可以用上面的'每周只加一件事'节奏。")
)
foreach($f in $faq){
    SetC $ws0 $row 1 $f[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $f[1] $false 10 0 0 $true; MC $ws0 $row 2 5
    $ws0.Range("A$($row):E$($row)").RowHeight=42; SB $ws0 $row 1 5; $row++
}

# 四周总览
$row+=2
SetC $ws0 $row 1 "  四周总览" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

SetC $ws0 $row 1 "周" $true 10 $white $headerBlue $false
SetC $ws0 $row 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws0 $row 3 "训练" $true 10 $white $headerBlue $false
SetC $ws0 $row 4 "有氧+步数" $true 10 $white $headerBlue $false
SetC $ws0 $row 5 "本周第一要务" $true 10 $white $headerBlue $false; $row++

$o4=@(
    @("1","唤醒","不训练","8000步/天","戒含糖饮料+3L水+睡够7h  三件做到两件=胜利"),
    @("2","入门","3次/周 固定器械","8000步+练后走10min","走进健身房就是胜利  怎么选重量看Sheet2"),
    @("3","适应","4次/周 上下肢分化","8000-10000步","三餐规律 每餐有蛋白质 晚餐换粗粮"),
    @("4","看到变化","4次/周 上下肢+有氧","10000步","拍对比照 量腰围 称体重  看你变了多少")
)
foreach($o4i in $o4){
    SetC $ws0 $row 1 $o4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws0 $row 2 $o4i[1] $true 10 $white $darkBg $false
    SetC $ws0 $row 3 $o4i[2] $false 10 0 0 $false
    SetC $ws0 $row 4 $o4i[3] $false 10 0 0 $false
    SetC $ws0 $row 5 $o4i[4] $false 9 0 0 $true
    $ws0.Range("A$($row):E$($row)").RowHeight=28; SB $ws0 $row 1 5; $row++
}

# 掉秤预期管理 V3新增
$row+=2
SetC $ws0 $row 1 "  关于掉秤速度——提前知道就不焦虑" $true 14 $white $softBlue $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$softBlue
$row++

$expectMgmt=@(
    "第一周可能会掉1-2kg（主要是水分——戒掉含糖饮料后身体开始排水）",
    "第二周开始变慢（0.5-1kg/周）——这才是真正的脂肪在掉",
    "如果某周体重没变：先看是不是经前（水肿1-3kg正常） 再看是不是吃了咸的（钠储水）  最后看腰围——腰围小了就是进步",
    "记住：你用了好几年胖到97kg，不可能用几周瘦到70kg。慢就是快。稳才能不反弹。"
)
foreach($em in $expectMgmt){
    SetC $ws0 $row 1 $em $false 10 0 0 $true; MC $ws0 $row 1 5
    $ws0.Range("A$($row):E$($row)").RowHeight=22; $row++
}

$row+=2
SetC $ws0 $row 1 "  快速跳转" $true 14 $white $darkBg $true
MC $ws0 $row 1 5; $ws0.Range("A$($row):E$($row)").RowHeight=26
$ws0.Range("A$($row):E$($row)").Font.Color=$white; $ws0.Range("A$($row):E$($row)").Interior.Color=$darkBg
$row++

$idx=@(
    @("方案逻辑+你想改善的部位分析"," Sheet 方案概览"),
    @("第2-4周训练 含怎么选重量"," Sheet 第2-4周训练"),
    @("热身怎么热身 拉伸+拍照怎么拍"," Sheet 热身有氧拉伸"),
    @("饮食怎么吃 外食/PMS/经期攻略"," Sheet 饮食与规则"),
    @("状态差/来例假/太累了 兜底"," Sheet 备用与复盘"),
    @("4周后怎么看效果 下一步"," Sheet 备用与复盘")
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

SetC $ws 1 1 "王丽颖  减脂塑形方案  概览" $true 18 $headerBlue 0 $true
MC $ws 1 1 8; $ws.Range("A1:H1").RowHeight=36

SetC $ws 2 1 "女  25-30岁  175cm / 97kg  BMI 31.7  零基础  减脂+改善体型  商业健身房" $false 11 $gold $darkBg $true
MC $ws 2 1 8; $ws.Range("A2:H2").Font.Color=$white; $ws.Range("A2:H2").Interior.Color=$darkBg

$row=4
SetC $ws $row 1 "方案概览" $true 14 $headerBlue $lightBlue $true; MC $ws $row 1 8

$ov=@(
    @("训练频率","第1周：不训练（只建立习惯）","第2周：3次/周 全身训练(固定器械)","第3-4周：4次/周 上下肢分化"),
    @("每次时长","总计50-60分钟","热身5min + 力量35-40min + 有氧10-15min + 拉伸5min"),
    @("训练方式","固定器械入门  徒手辅助  第3周起逐步引入轻哑铃","先建立自信和安全动作模式——再谈重量"),
    @("方案周期","4周为起点","4周后根据体重变化+腰围+体能+主观感受调整"),
    @("方案特点","零基础渐进入门  先习惯后训练  固定器械安全感  饮食第一训练第二")
)
$r=5
foreach($o in $ov){
    SetC $ws $r 1 $o[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 3 $o[1] $false 10 0 0 $true; MC $ws $r 3 5
    SetC $ws $r 6 $o[2] $false 10 0 0 $true; MC $ws $r 6 8
    SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "为什么这样设计（很重要，请看完）" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8; $r++

$design=@(
    "你之前办了健身卡基本没来过——不是你的问题，是一上来就要'训练'这个门槛太高了。",
    "本方案第一周不让你训练。只做三件事：每天8000步、3L水、睡够7小时。这三件做到了 体重就会开始动。",
    "为什么？对97kg来说，饮食和日常活动量是减脂的绝对主力。训练是加速器，不是发动机。",
    "第二周走进健身房——全部固定器械。轨道固定、不容易做错、不容易受伤、一个人也敢做。",
    "第三周加码到每周4练，上下肢分开——因为身体开始适应了，可以承受更多。",
    "整个过程像是'慢慢点火'，不是'上来就冲刺'。你要的不是坚持一个月然后放弃——是养成习惯，再也不胖回去。"
)
foreach($d in $design){
    SetC $ws $r 1 $d $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws $r 1 "四周节奏一览" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8; $r++

SetC $ws $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws $r 2 "阶段" $true 10 $white $headerBlue $false
SetC $ws $r 3 "训练日" $true 10 $white $headerBlue $false
SetC $ws $r 5 "主题" $true 10 $white $headerBlue $false
SetC $ws $r 7 "本周目标" $true 10 $white $headerBlue $false; $r++

$w4=@(
    @("1","唤醒","不训练","","建立习惯","三件事做到两件=胜利。体重开始动。"),
    @("2","入门","3次/周","全身(固定器械)","走进健身房","第一次训练完成就是胜利。固定器械建立安全感。"),
    @("3","适应","4次/周","上下肢分化","感觉'我可以'","固定器械+少量徒手+轻哑铃。身体开始紧致。"),
    @("4","看到变化","4次/周","上下肢+有氧","拍对比照","体能提升。拍标准三件套照片对比第一天。量腰围称体重。")
)
foreach($w4i in $w4){
    SetC $ws $r 1 $w4i[0] $true 10 $headerBlue $lightGray $false
    SetC $ws $r 2 $w4i[1] $true 10 $white $darkBg $false
    SetC $ws $r 3 $w4i[2] $false 10 0 0 $false
    SetC $ws $r 5 $w4i[3] $false 10 0 0 $false
    SetC $ws $r 7 $w4i[4] $false 10 0 0 $true
    $ws.Range("A$($r):H$($r)").RowHeight=26; SB $ws $r 1 8; $r++
}

$r++
SetC $ws $r 1 "关于你提到的那些想改善的部位" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8; $r++

$bg=@(
    "你想改善：臀凹陷、大小腿粗、妈妈臀、胯宽、背变薄。这些目标我全部记下来了。",
    "但有一个事实你先知道：没有'局部减脂'——不能只瘦大腿不瘦脸，脂肪是全身一起掉的。",
    "好消息：当你全身脂肪率下降，上面这些部位全都会变好看。同时：",
    "  臀凹陷/妈妈臀  通过髋外展+臀桥增加臀大肌肌肉量填充",
    "  大小腿粗  脂肪减少后腿自然变细；小腿粗也和水肿有关——多喝水+控盐有帮助",
    "  背变薄  高位下拉+坐姿划船让背部紧致",
    "  胯宽  髋外展改善臀部外侧线条",
    "到第4周结束时，这些部位不一定已经'完美'——但你会看到它们正在往好的方向走。"
)
foreach($b in $bg){
    SetC $ws $r 1 $b $false 10 0 0 $true; MC $ws $r 1 8
    $ws.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws $r 1 "三个核心原则" $true 14 $headerBlue $lightBlue $true
MC $ws $r 1 8; $r++

$pr=@(
    @("1","先习惯后训练","第一周不进健身房。先建立走路+喝水+睡眠的习惯——这三件做好了体重就开始掉。"),
    @("2","完成比完美重要","8000步没走到？走5000步也比不走强。今天没练？明天补。不追求完美——追求持续。"),
    @("3","饮食是主力，训练是加速器","你现在97kg——减脂的80%靠吃。训练让身材紧致有线条，但体重下降主要是吃出来的。")
)
foreach($p in $pr){
    SetC $ws $r 1 $p[0] $false 10 0 0 $false
    SetC $ws $r 2 $p[1] $true 10 $headerBlue 0 $false
    SetC $ws $r 3 $p[2] $false 10 0 0 $true; MC $ws $r 3 8
    $ws.Range("A$($r):H$($r)").RowHeight=26; SB $ws $r 1 8; $r++
}

$ws.Range("A:A").ColumnWidth=30; $ws.Range("B:H").ColumnWidth=16
Write-Host "Sheet 1 done"

# ============================================
# SHEET 2: 第2-4周训练计划
# ============================================
$ws2=$wb.Worksheets.Add();$ws2.Name="第2-4周训练"

SetC $ws2 1 1 "第2-4周  训练计划" $true 16 $headerBlue 0 $true
MC $ws2 1 1 8; $ws2.Range("A1:H1").RowHeight=34
SetC $ws2 2 1 "第1周不训练（只做封面三件事）。第2周走进健身房。训练前先看下方[怎么选重量]。" $false 10 $headerBlue $lightGray $true
MC $ws2 2 1 8

# 选重量指南 V2
$r=4
SetC $ws2 $r 1 "  怎么选重量——每次训练前看这个" $true 14 $white $accentGreen $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$accentGreen
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
    SetC $ws2 $r 1 $wg $false 10 0 0 $true; MC $ws2 $r 1 8
    $ws2.Range("A$($r):H$($r)").RowHeight=20
    if($wg -eq ""){$ws2.Range("A$($r):H$($r)").RowHeight=10}
    $r++
}

# 第2周
$r++
SetC $ws2 $r 1 "第2周  全身训练  每周3次（如周一/三/五）" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第2周目标：走进健身房就是胜利。固定器械轨道固定、安全、一个人也敢做。三个训练日做同样的内容——先熟悉动作，不急着变换。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; $r++

$actsW2=@(
    @("坐姿推胸","3组","12-15次","先挂最轻片 按选重三步法","60秒","轻松-中等","【呼吸】推出呼气、收回吸气。【要点】后背贴紧靠背。推到前面不锁死手肘。先找到'胸肌在用力'的感觉。"),
    @("坐姿划船","3组","12-15次","先挂最轻片 按选重三步法","60秒","轻松-中等","【呼吸】拉向身体呼气、还原吸气。【要点】身体不要后仰——想象肩胛骨往中间夹。拉到腹部位置。"),
    @("坐姿腿举","3组","12-15次","空机或最轻片","60秒","轻松-中等","【呼吸】推出呼气、收回吸气。【要点】脚踩实与肩同宽。膝盖不要完全蹬直——留一点弯曲保护关节。"),
    @("坐姿髋外展","3组","15-20次","轻重量——臀外侧有感就行","45秒","轻松","【呼吸】打开呼气、收回吸气。【要点】身体贴紧靠背不晃。感受臀外侧发力。改善臀凹陷和胯宽。"),
    @("高位下拉","3组","12-15次","最轻片","60秒","轻松","【呼吸】拉下呼气、还原吸气。【要点】先沉肩膀再用手臂拉。杆拉到锁骨高度。背变薄的第一步。"),
    @("平板支撑","2组","20-30秒","自重","30秒","力竭前停","【要点】手肘正下方撑地。腰不要塌——想象有人要打你肚子。能做多久做多久。第一次20秒也很棒。")
)
$r=Write-TrainingBlock $ws2 $r "第2周训练（三个训练日都做这个）" "固定器械入门  先找到发力感  |  35-40分钟" "动作质量优先——不要和别人比" "38-42" $actsW2

# 第3-4周
$r++
SetC $ws2 $r 1 "第3-4周  上下肢分化  每周4次（例：周一/二/四/五）" $true 14 $white $darkBg $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=26
$ws2.Range("A$($r):H$($r)").Font.Color=$white; $ws2.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws2 $r 1 "第3周起上下肢分开。重量比第2周重1-2片（仍用选重三步法判断）。第4周再尝试+1片或+1-2kg哑铃。" $false 10 $headerBlue $lightGray $true
MC $ws2 $r 1 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; $r++

$actsLA=@(
    @("坐姿腿举","3组","10-12次","比第2周重1-2片(三步法)","75秒","中等偏强","【呼吸】推出呼气、回收吸气。【发力感】脚后跟发力蹬，大腿前侧和臀部一起用力。"),
    @("坐姿髋外展","3组","15-20次","比第2周重1片","45秒","中等","【发力感】打开时臀外侧收紧保持1秒。改善胯宽和臀凹陷的关键动作。"),
    @("坐姿腿屈伸","3组","12-15次","轻-中重量","60秒","中等","【呼吸】踢出呼气、放下吸气。【要点】脚尖回勾，大腿前侧发力。膝盖不舒服降重量。"),
    @("哑铃臀桥","3组","12-15次","双手持哑铃6-10kg放髋部","60秒","中等","【呼吸】推起呼气、下放吸气。【要点】下巴微收。顶峰夹紧屁股1-2秒。改善臀凹陷和妈妈臀。"),
    @("坐姿腿弯举","3组","12-15次","轻重量——大腿后侧有感","60秒","中等","【呼吸】勾起呼气、还原吸气。【要点】大腿后侧发力勾下。改善大腿后侧线条。"),
    @("平板支撑","3组","30-45秒","自重","30秒","力竭前","比第2周多5-10秒。")
)
$r=Write-TrainingBlock $ws2 $r "下肢日 A（第3-4周 每周2次）" "股四头+臀+腘绳  改善臀凹陷/妈妈臀/大腿线条  |  40-45分钟" "比第2周重1-2片 先保动作质量" "40-45" $actsLA

$actsUA=@(
    @("坐姿推胸","3组","10-12次","比第2周重1-2片(三步法)","75秒","中等","【发力感】推到前面胸肌向中间挤压保持1秒。背贴紧靠背。"),
    @("高位下拉","3组","10-12次","比第2周重1片","75秒","中等","【发力感】拉到锁骨背阔肌收紧。肩胛先下沉再拉——别用手臂硬拽。背薄的关键。"),
    @("坐姿推肩","3组","12-15次","最轻片起步","60秒","中等","【呼吸】推起呼气、下放吸气。【要点】下背贴紧靠背。推到上面不锁肘。肩部线条。"),
    @("坐姿划船","3组","10-12次","比第2周重1片","60秒","中等","【发力感】肩胛骨向中间夹。改善圆肩驼背——让上半身体态更直。"),
    @("哑铃侧平举","3组","12-15次","最轻哑铃2-4kg","45秒","中等","【要点】不要耸肩——用肩膀外侧发力。手臂微屈。改善肩部线条。"),
    @("仰卧举腿","2组","10-15次","自重","45秒","中等","【要点】用下腹带动腿上举，不靠惯性摆。")
)
$r=Write-TrainingBlock $ws2 $r "上肢日 A（第3-4周 每周2次）" "胸+背+肩+手臂  改善背厚/圆肩/上半身体态  |  40-45分钟" "全部固定器械 一个人安全" "40-45" $actsUA

# 排班
$r++
SetC $ws2 $r 1 "第3-4周训练排班" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "推荐" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周一=下肢 / 周二=上肢 / 周四=下肢 / 周五=上肢 / 周三六日=休息+走路" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "备选" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "周一=下肢 / 周三=上肢 / 周四=下肢 / 周六=上肢——如果连着练两天太累" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8; $r++

SetC $ws2 $r 1 "只能2-3次" $true 10 $headerBlue $lightGray $false
SetC $ws2 $r 2 "第3周退回做第2周全身体系。第4周优先保下肢日。详见备用与复盘。" $false 10 0 0 $true
MC $ws2 $r 2 8; $ws2.Range("A$($r):H$($r)").RowHeight=22; SB $ws2 $r 1 8

$r+=2
SetC $ws2 $r 1 "第3周  第4周怎么加码" $true 12 $headerBlue 0 $true; MC $ws2 $r 1 8; $r++

$p34=@(
    "第3周：先做到每次训练顺利完成。重量保守——感觉对了比数字重要。",
    "第4周：每个动作尝试比第3周重1片或+1-2kg哑铃。动作变形→退回第3周重量。",
    "有氧：第3周训练后走10分钟(跑步机坡度5-8% 速度4-5km/h)；第4周走12-15分钟。",
    "第4周结束时的信号：你回头看第2周的动作，觉得'太轻松了'——就是进步。"
)
foreach($p in $p34){
    SetC $ws2 $r 1 $p $false 10 0 0 $true; MC $ws2 $r 1 8
    $ws2.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$ws2.Range("A:A").ColumnWidth=4; $ws2.Range("B:B").ColumnWidth=20
$ws2.Range("C:C").ColumnWidth=7; $ws2.Range("D:D").ColumnWidth=12
$ws2.Range("E:E").ColumnWidth=20; $ws2.Range("F:F").ColumnWidth=8
$ws2.Range("G:G").ColumnWidth=10; $ws2.Range("H:H").ColumnWidth=56
Write-Host "Sheet 2 done"

# ============================================
# SHEET 3: 有氧+热身+拉伸+拍照
# ============================================
$ws3=$wb.Worksheets.Add();$ws3.Name="热身有氧拉伸"

SetC $ws3 1 1 "有氧安排  热身  拉伸  拍照标准" $true 16 $headerBlue 0 $true
MC $ws3 1 1 8; $ws3.Range("A1:H1").RowHeight=34

$r=3
SetC $ws3 $r 1 "每天步数目标（比训练更重要）" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "步数" $true 10 $white $headerBlue $false
SetC $ws3 $r 4 "怎么做到" $true 10 $white $headerBlue $false; MC $ws3 $r 4 8; $r++

$steps=@(
    @("第1周","8000步","拆开走：上班多走一站+午休10min+晚饭后20min+碎片走动。走不够5000也比不走强。"),
    @("第2周","8000步+练后走10min","保持8000步  训练日跑步机多走10min(坡度5-8% 速度4-5km/h)"),
    @("第3周","8000-10000步","在第2周基础上 周末多走一会儿。不需要一口气——碎片时间凑够。"),
    @("第4周","10000步","保持  第1周8000步觉得吃力的话 第4周做到10000步你会惊讶的")
)
foreach($st in $steps){
    SetC $ws3 $r 1 $st[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $st[1] $true 10 $headerBlue 0 $false
    SetC $ws3 $r 4 $st[2] $false 10 0 0 $true; MC $ws3 $r 4 8
    $ws3.Range("A$($r):H$($r)").RowHeight=36; SB $ws3 $r 1 8; $r++
}

# Warmup
$r+=2
SetC $ws3 $r 1 "训练前热身（5分钟  必须做）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "步" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 5 "要点" $true 9 $white $headerBlue $false; MC $ws3 $r 5 8; $r++

$wus=@(
    @("1","原地踏步+摆臂","1min","逐渐加快 微微发热"),
    @("2","肩关节绕圈","前10 后10","画大圈——推胸下拉前必做"),
    @("3","扶墙摆腿(前后+左右)","各10/侧","拉开大腿前后——腿举和腿弯举前的重要准备"),
    @("4","扶椅背自重深蹲","10次","能蹲多深蹲多深 找'大腿在用力'的感觉"),
    @("5","第一个动作空机/最轻重量试做","1组x10次","让身体知道'接下来要干这个了'")
)
foreach($w in $wus){
    SetC $ws3 $r 1 $w[0] $false 9 0 0 $false
    SetC $ws3 $r 2 $w[1] $true 9 $headerBlue 0 $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $w[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $w[3] $false 9 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=24; SB $ws3 $r 1 8; $r++
}

# Stretch
$r+=2
SetC $ws3 $r 1 "训练后拉伸（5分钟  帮你第二天不那么酸痛）" $true 14 $white $darkBg $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws3 $r 1 "序" $true 9 $white $headerBlue $false
SetC $ws3 $r 2 "动作" $true 9 $white $headerBlue $false; MC $ws3 $r 2 3
SetC $ws3 $r 4 "时间" $true 9 $white $headerBlue $false
SetC $ws3 $r 5 "要点" $true 9 $white $headerBlue $false; MC $ws3 $r 5 8; $r++

$strs=@(
    @("1","胸大肌拉伸(扶墙)","30s/侧","推胸后——手臂张开扶墙 身体反方向转"),
    @("2","背阔肌拉伸(扶墙侧屈)","30s/侧","下拉/划船后——手扶墙 身体侧弯"),
    @("3","大腿前侧拉伸(站姿脚跟贴臀)","30s/侧","扶墙——腿举/腿屈伸后"),
    @("4","大腿后侧拉伸(坐姿直腿前屈)","30s","腿弯举后——坐地上 伸直腿 手够脚尖"),
    @("5","臀拉伸(跷二郎腿抱膝)","30s/侧","髋外展/臀桥后——坐椅子上 脚踝搭对侧膝盖"),
    @("6","深呼吸","10次慢吸慢呼","让心跳慢下来——告诉身体'训练结束 可以恢复了'")
)
foreach($s in $strs){
    SetC $ws3 $r 1 $s[0] $false 9 0 0 $false
    SetC $ws3 $r 2 $s[1] $true 9 $headerBlue 0 $false; MC $ws3 $r 2 3
    SetC $ws3 $r 4 $s[2] $false 9 0 0 $false
    SetC $ws3 $r 5 $s[3] $false 9 0 0 $true; MC $ws3 $r 5 8
    $ws3.Range("A$($r):H$($r)").RowHeight=30; SB $ws3 $r 1 8; $r++
}

# Photo guide V2
$r+=2
SetC $ws3 $r 1 "  拍照标准——今天就拍 4周后同一条件对比" $true 14 $white $accentGreen $true
MC $ws3 $r 1 8; $ws3.Range("A$($r):H$($r)").RowHeight=26
$ws3.Range("A$($r):H$($r)").Font.Color=$white; $ws3.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws3 $r 1 "今天拍一组。4周后同一时间、同一光线、同一套衣服再拍。不要吸肚子！不要找角度！真实的才有对比价值。" $false 10 $headerBlue $lightGray $true
MC $ws3 $r 1 8; $r++

SetC $ws3 $r 1 "拍什么" $true 10 $white $headerBlue $false
SetC $ws3 $r 2 "怎么拍" $true 10 $white $headerBlue $false; MC $ws3 $r 2 5
SetC $ws3 $r 6 "要点" $true 10 $white $headerBlue $false; $r++

$photoGuide=@(
    @("正面照","自然站立 双手放两侧","手机与肚脐同高 全身入镜 不要侧身"),
    @("侧面照","侧身站立 手臂自然下垂","手机与腰同高 能看到肚子和背部的轮廓"),
    @("背面照","背对镜子 双手叉腰","手机与肩胛骨同高 能看到背部线条和臀型"),
    @("穿什么","运动内衣+紧身裤/短裤","露出腰腹 才能看到变化。不要穿宽松T恤拍。"),
    @("什么时候拍","今天(第1天) 第4周最后一天","同一时间(建议早上空腹) 同一位置 同一光线")
)
foreach($pg in $photoGuide){
    SetC $ws3 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    SetC $ws3 $r 2 $pg[1] $false 10 0 0 $false; MC $ws3 $r 2 5
    SetC $ws3 $r 6 $pg[2] $false 10 0 0 $true
    $ws3.Range("A$($r):H$($r)").RowHeight=30; SB $ws3 $r 1 8; $r++
}

$ws3.Range("A:A").ColumnWidth=8; $ws3.Range("B:B").ColumnWidth=22
$ws3.Range("C:C").ColumnWidth=14; $ws3.Range("D:D").ColumnWidth=14; $ws3.Range("E:H").ColumnWidth=18
Write-Host "Sheet 3 done"

# ============================================
# SHEET 4: 饮食+外食+PMS+粗粮+规则
# ============================================
$ws4=$wb.Worksheets.Add();$ws4.Name="饮食与规则"

SetC $ws4 1 1 "饮食方案  外食指南  PMS/经期应对  判断规则" $true 16 $headerBlue 0 $true
MC $ws4 1 1 8; $ws4.Range("A1:H1").RowHeight=34

# Nutrition
$r=3
SetC $ws4 $r 1 "饮食方案（减脂主力——80%靠这个）" $true 14 $white $accentGreen $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

SetC $ws4 $r 1 "你的现状：97kg，想快速减脂。以下数字是方向——不需要称克数。理解原则，生活化执行。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=22; $r++

SetC $ws4 $r 1 "营养素" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "目标" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "怎么吃（不称克数版本）" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$nutri=@(
    @("总热量","1600-1800","比你现在的饮食少500-700kcal——刚好让体重每周掉0.5-1kg，但不至于饿得受不了。"),
    @("蛋白质","130-160g","每餐一份手掌大蛋白质：早2蛋+奶 午晚各一份肉/鱼/豆腐。练后喝蛋白粉更容易达标。"),
    @("碳水","160-180g","三餐每餐一拳。优先杂粮饭/红薯/玉米/南瓜。晚餐换成粗粮——饱腹感更强，睡前不饿。如果今天没有粗粮：白米饭一拳也可以——吃对总量比吃对种类重要。"),
    @("脂肪","40-50g","炒菜控油。不吃油炸。坚果一天一小把够了。"),
    @("饮水","3L/天","买1.5L大水壶 上午一壶下午一壶。水里加柠檬片。戒所有含糖饮料——这是你第一周掉秤的最大来源。"),
    @("蔬菜","不限量","每餐占一半盘子。先吃蔬菜再吃肉最后吃饭——这个顺序血糖更稳，饱腹感更强，掉秤更快。")
)
foreach($n in $nutri){
    SetC $ws4 $r 1 $n[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $n[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $n[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=38; SB $ws4 $r 1 8; $r++
}

# 粗粮替换清单 V3
$r++
SetC $ws4 $r 1 "晚餐粗粮替换清单（不用记 照着买）" $true 12 $headerBlue 0 $true; MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "精制碳水（少吃）" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "换成粗粮（推荐）" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "为什么" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$grains=@(
    @("白米饭","糙米饭、杂粮饭","膳食纤维多 饱腹感强 血糖稳"),
    @("白面包","全麦面包","同样的量 更扛饿"),
    @("普通面条","荞麦面、意面","升糖慢 下午不容易犯困"),
    @("土豆泥","烤红薯、蒸南瓜","天然甜味 满足碳水渴望"),
    @("甜点/饼干","玉米、燕麦","天然食材 没有添加糖")
)
foreach($g in $grains){
    SetC $ws4 $r 1 $g[0] $true 10 $accentRed $lightGray $false
    SetC $ws4 $r 2 $g[1] $true 10 $accentGreen 0 $false
    SetC $ws4 $r 4 $g[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

# 三餐示例
$r++
SetC $ws4 $r 1 "三餐示例（照着这个方向吃 不用精确复制）" $true 12 $headerBlue 0 $true
MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "餐次" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "吃什么" $true 10 $white $headerBlue $false; MC $ws4 $r 2 5
SetC $ws4 $r 6 "要点" $true 10 $white $headerBlue $false; $r++

$meals=@(
    @("早餐","2个鸡蛋+一杯牛奶+一片全麦吐司或一小碗燕麦","蛋白质吃够上午不饿。鸡蛋别只吃蛋白——蛋黄营养对女性重要。"),
    @("午餐","米饭一拳+一份肉/鱼(手掌大)+蔬菜占半盘","食堂/外卖：选清蒸/炖/卤 避开油炸和勾芡。先吃菜再吃肉最后吃饭。"),
    @("下午加餐","小水果/无糖酸奶/坚果10颗以内","饿了吃 不饿可以不吃。不要饿到晚餐暴食。"),
    @("晚餐","粗粮一拳(糙米/红薯/玉米/南瓜)+一份肉/鱼/豆腐+大量蔬菜","晚餐吃粗粮不饿肚子。睡前饿了：喝温水/吃黄瓜/小番茄——不要硬扛。"),
    @("训练后","蛋白粉+水或牛奶","训练完30分钟内喝。不买蛋白粉——训练后半小时内吃含蛋白质的正餐。")
)
foreach($m in $meals){
    SetC $ws4 $r 1 $m[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $m[1] $false 10 0 0 $true; MC $ws4 $r 2 5
    SetC $ws4 $r 6 $m[2] $false 10 0 0 $true
    $ws4.Range("A$($r):H$($r)").RowHeight=36; SB $ws4 $r 1 8; $r++
}

# 外食指南 V2
$r+=2
SetC $ws4 $r 1 "如果今天吃外卖/外食——这样选" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "不要求顿顿完美。一天里有一顿吃对了——就比昨天好。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $r++

SetC $ws4 $r 1 "  推荐选这些" $true 12 $accentGreen 0 $true; MC $ws4 $r 1 8; $r++

$goodEats=@(
    @("沙县小吃","鸡腿饭(去皮)+烫青菜","蛋白质够 碳水适中 蔬菜有"),
    @("麻辣烫","清汤底 多蔬菜+豆制品+瘦肉 少丸子","丸子=淀粉+油 少拿"),
    @("轻食沙拉","酱汁另放 只蘸不浇","沙拉酱=热量炸弹 蘸着吃别浇上去"),
    @("便利店","鸡胸肉+茶叶蛋+无糖豆浆+玉米","方便便宜 一顿蛋白质够了"),
    @("潮汕牛肉/粿条","选牛肉+蔬菜 汤底少喝","汤底盐分高 喝几口就好")
)
foreach($ge in $goodEats){
    SetC $ws4 $r 1 $ge[0] $true 10 $accentGreen $lightGray $false
    SetC $ws4 $r 2 $ge[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $ge[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$r++
SetC $ws4 $r 1 "  尽量避开这些" $true 12 $accentRed 0 $true; MC $ws4 $r 1 8; $r++

$badEats=@(
    @("炒饭/炒面/盖浇饭","油+碳水双高——减脂期最大的隐形热量来源"),
    @("油炸类(炸鸡排/炸猪排/炸串)","裹粉吸油=热量翻倍"),
    @("含糖饮料(奶茶/果汁/可乐)","液体热量——不占肚子却400-600kcal"),
    @("勾芡/红烧/糖醋类","淀粉+油+糖三高——选清蒸/白切/卤/炖替代")
)
foreach($be in $badEats){
    SetC $ws4 $r 1 $be[0] $true 10 $accentRed $lightGray $false
    SetC $ws4 $r 2 $be[1] $false 10 0 0 $true; MC $ws4 $r 2 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

# PMS V2
$r+=2
SetC $ws4 $r 1 "经期前一周（PMS阶段）——提前知道就不慌" $true 14 $white $softPink $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$softPink
$r++

SetC $ws4 $r 1 "很多女性在经前一周食欲暴增、情绪波动、水肿——这时候最容易破功。提前告诉你：这是正常的激素波动，不是你意志力差。" $false 10 $headerBlue $lightGray $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

SetC $ws4 $r 1 "你会遇到的" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "为什么" $true 10 $white $headerBlue $false
SetC $ws4 $r 4 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 4 8; $r++

$pms=@(
    @("体重涨1-3kg","正常水肿 不是胖了——激素导致水分滞留","别称体重 或少称。经期结束后水自然排掉。"),
    @("食欲暴增 特别想吃甜的","激素波动——不是你意志力差","可以吃 但选相对健康的：黑巧克力(70%+)2小块/水果/无糖酸奶+蜂蜜。"),
    @("情绪低落 不想训练","正常——经前雌激素下降影响情绪","训练强度降一级(重量减/组数减)。但尽量保持去健身房——哪怕只做热身+有氧。经前一周'维持'就是胜利。"),
    @("训练时觉得比平时累","身体在准备来月经 能量储备偏低","降重量。感觉比平时低1-2分。不要求'突破'——能完成就是胜利。")
)
foreach($pm in $pms){
    SetC $ws4 $r 1 $pm[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $pm[1] $false 10 0 0 $false
    SetC $ws4 $r 4 $pm[2] $false 10 0 0 $true; MC $ws4 $r 4 8
    $ws4.Range("A$($r):H$($r)").RowHeight=42; SB $ws4 $r 1 8; $r++
}

$r++
SetC $ws4 $r 1 "来例假时" $true 12 $accentRed 0 $true; MC $ws4 $r 1 8; $r++
SetC $ws4 $r 1 "经期第1-2天休息或只走路。第3天后如果OK恢复轻量训练。经期训练不勉强——身体在流血需要休息，不是偷懒。" $false 10 0 0 $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=24; $r++

# 习惯重点
SetC $ws4 $r 1 "这4周的生活习惯重点" $true 12 $headerBlue 0 $true; MC $ws4 $r 1 8; $r++

$habitsWang=@(
    @("第1周","戒含糖饮料 每天3L水 睡够7h","三件事做到任意两件=成功"),
    @("第2周","三餐规律 每餐有蛋白质 晚餐换粗粮","走进健身房同时饮食跟上 不要求完美要求持续"),
    @("第3周","用'先吃菜再吃肉最后吃饭'的顺序","自然少吃饭——不用挨饿"),
    @("第4周","复盘：哪餐最难坚持？哪个习惯已经不需要想了？","庆祝养成的习惯 修补困难的环节")
)
SetC $ws4 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws4 $r 2 "重点" $true 10 $white $headerBlue $false
SetC $ws4 $r 5 "说明" $true 10 $white $headerBlue $false; MC $ws4 $r 5 8; $r++
foreach($hw in $habitsWang){
    SetC $ws4 $r 1 $hw[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 2 $hw[1] $false 10 0 0 $false
    SetC $ws4 $r 5 $hw[2] $false 10 0 0 $true; MC $ws4 $r 5 8
    $ws4.Range("A$($r):H$($r)").RowHeight=24; SB $ws4 $r 1 8; $r++
}

# Rules
$r+=2
SetC $ws4 $r 1 "遇到这些情况怎么判断" $true 14 $white $darkBg $true
MC $ws4 $r 1 8; $ws4.Range("A$($r):H$($r)").RowHeight=26
$ws4.Range("A$($r):H$($r)").Font.Color=$white; $ws4.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws4 $r 1 "这种情况" $true 10 $white $headerBlue $false
SetC $ws4 $r 3 "怎么办" $true 10 $white $headerBlue $false; MC $ws4 $r 3 8; $r++

$rules=@(
    @("这周训练太轻松","下周每个动作加1片或+1-2kg哑铃——先保动作不变形"),
    @("训练后48h还酸痛","零基础前两周就会酸 正常。多喝水+多睡+拉伸。关节疼不是肌肉疼  告诉我。"),
    @("这周体重没掉","先看是不是经期前后——经前水肿1-3kg正常。不是的话 回头检查饮料和步数。"),
    @("经期前一周体重涨了","正常水肿！不是胖了！经期结束水就排掉。"),
    @("来例假了","第1-2天休息或只走路。第3天后如果OK恢复轻量训练。不勉强。"),
    @("太累了不想练","走5000步+拉伸——做了就算。但别连续跳过2次。"),
    @("今天吃多了/喝了奶茶","明天继续 别破罐破摔等周一。一顿吃多不会胖——连续一周才会。"),
    @("膝盖/腰/关节疼","立即停那个动作 换不疼的。疼超2天告诉我。")
)
foreach($r3 in $rules){
    SetC $ws4 $r 1 $r3[0] $true 10 $headerBlue $lightGray $false
    SetC $ws4 $r 3 $r3[1] $false 10 0 0 $true; MC $ws4 $r 3 8
    $ws4.Range("A$($r):H$($r)").RowHeight=26; SB $ws4 $r 1 8; $r++
}

$ws4.Range("A:A").ColumnWidth=28; $ws4.Range("B:H").ColumnWidth=18
Write-Host "Sheet 4 done"

# ============================================
# SHEET 5: 备用+复盘
# ============================================
$ws5=$wb.Worksheets.Add();$ws5.Name="备用与复盘"

SetC $ws5 1 1 "降级方案  每周一件事路径  4周复盘  下一步" $true 16 $headerBlue 0 $true
MC $ws5 1 1 8; $ws5.Range("A1:H1").RowHeight=34

# Fallback
$r=3
SetC $ws5 $r 1 "状态差/加班/事情多——降级方案" $true 14 $white $accentOrange $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentOrange
$r++

SetC $ws5 $r 1 "宁可降级训练 不要完全不练。降级不丢人——连续放弃才是问题。" $false 10 $headerBlue $lightGray $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "本来计划" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "降级为" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "最少完成这个" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$fallback=@(
    @("4次/周训练","减到2-3次 保留下肢日","哪怕一周只练1次 也比不练强"),
    @("今天该练但很累","热身+最轻重量做2-3个动作+拉伸","走进健身房待20分钟=胜利"),
    @("完全不想出门","在家/楼下走20分钟+10分钟拉伸","也算运动了——走路永远不丢人"),
    @("这周饮食全乱了","明天重新开始","不等周一——明天早上就可以重来"),
    @("来例假前3天","只走路 不训练","身体在准备月经 需要休息——不是偷懒"),
    @("经前特别想吃甜的","吃！但选黑巧克力/水果/酸奶","不是意志力差——是激素。满足但不过量。")
)
foreach($fb in $fallback){
    SetC $ws5 $r 1 $fb[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $fb[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $fb[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=24; SB $ws5 $r 1 8; $r++
}

# Slow habits (copy from cover for reference)
$r+=2
SetC $ws5 $r 1 "如果第一周四件事太难——'每周只加一件事'路径" $true 14 $white $softPurple $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$softPurple
$r++

$slowHabits=@(
    @("第1周","只做一件事：戒含糖饮料","其他三件能做就做 做不到不批评自己"),
    @("第2周","戒含糖饮料 + 每天8000步","第一个习惯稳了 再加第二个"),
    @("第3周","戒含糖饮料 + 8000步 + 3L水","前两个已经不用'坚持'了"),
    @("第4周","四件事全做到","这时候你已经是一个'有运动习惯的人'了")
)
SetC $ws5 $r 1 "周" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "只加一件事" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "怎么做" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++
foreach($sh in $slowHabits){
    SetC $ws5 $r 1 $sh[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $sh[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $sh[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=32; SB $ws5 $r 1 8; $r++
}

# Review
$r+=2
SetC $ws5 $r 1 "4周后——怎么知道自己有没有进步" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

SetC $ws5 $r 1 "不要只看体重秤！以下每一条都是进步：" $false 11 $accentGreen 0 $true
MC $ws5 $r 1 8; $r++

SetC $ws5 $r 1 "看什么" $true 10 $white $headerBlue $false
SetC $ws5 $r 2 "可能的变化" $true 10 $white $headerBlue $false
SetC $ws5 $r 4 "怎么看/为什么" $true 10 $white $headerBlue $false; MC $ws5 $r 4 8; $r++

$progress=@(
    @("体重","掉了1-4kg","周三早上空腹称。不要每天称——女性体重每天波动。经期前后不称——水肿骗你。"),
    @("腰围","小了2-5cm","软尺量肚脐一圈。比体重更真实——体重没变腰细了=脂肪换肌肉。"),
    @("照片","不一样了","翻出第一天的三张照片 同一位置同一光线再拍一组。照片不会骗人——你会看到的。"),
    @("衣服","松了","裤腰能塞进一个手指 上衣不绷了 手臂线条出来了——体重秤测不出。"),
    @("体力","变好了","从8000步气喘到10000步不累。从'完全不想动'到'今天没动有点不舒服'。"),
    @("自信","敢进健身房了","第2周第一次进可能紧张。第4周你知道每个器械怎么用了——你不再是一个'不运动的人'了。"),
    @("习惯","不用逼自己了","走路/喝水/不喝甜饮料——已经从'任务'变成'自然'。这个，才是永远不会反弹的保证。")
)
foreach($pg in $progress){
    SetC $ws5 $r 1 $pg[0] $true 10 $headerBlue $lightGray $false
    SetC $ws5 $r 2 $pg[1] $false 10 0 0 $false
    SetC $ws5 $r 4 $pg[2] $false 10 0 0 $true; MC $ws5 $r 4 8
    $ws5.Range("A$($r):H$($r)").RowHeight=36; SB $ws5 $r 1 8; $r++
}

# Next steps
$r+=2
SetC $ws5 $r 1 "4周之后" $true 14 $white $darkBg $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$darkBg
$r++

$next=@(
    "减了2-4kg 且体能变好：继续保持这套方案再走4周。第5-8周加重+加组。",
    "体重没动 但腰细了体力好了：脂肪在换成肌肉——继续！你已经不一样了。",
    "完全没变化：回头检查饮料和步数——这两个先做到。饮食记录三天看看真实摄入。",
    "你已经不害怕健身房了：告诉我。下一周期增加更多训练变化。"
)
foreach($ns in $next){
    SetC $ws5 $r 1 $ns $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=20; $r++
}

$r++
SetC $ws5 $r 1 "你学会自己做计划的第一步" $true 14 $white $accentGreen $true
MC $ws5 $r 1 8; $ws5.Range("A$($r):H$($r)").RowHeight=26
$ws5.Range("A$($r):H$($r)").Font.Color=$white; $ws5.Range("A$($r):H$($r)").Interior.Color=$accentGreen
$r++

$self=@(
    "现在不需要你自己做计划——这套方案已经做好了4周。",
    "4周后你只需要记住三句话：",
    "  1. 每周至少练3次——每次选4-6个固定器械动作 每个3组12-15次",
    "  2. 每天走路8000-10000步——能一边走一边断续说话的速度",
    "  3. 戒糖饮 多喝水 每餐有蛋白质 晚餐吃粗粮",
    "够简单才能做一辈子。不需要复杂的周期化 不需要精密的RPE。",
    "",
    "最后的话：你之前问过私教 办了卡也没来——那不是你的问题。是之前的方案门槛太高了。",
    "现在这份方案 第一周不进健身房。做到封面三件事 体重就开始动。",
    "你离你想要的体型——只差一个'开始'。而这个开始 简单到不可能失败。"
)
foreach($s in $self){
    SetC $ws5 $r 1 $s $false 10 0 0 $true; MC $ws5 $r 1 8
    $ws5.Range("A$($r):H$($r)").RowHeight=20
    if($s -eq ""){$ws5.Range("A$($r):H$($r)").RowHeight=10}
    $r++
}

$ws5.Range("A:A").ColumnWidth=28; $ws5.Range("B:H").ColumnWidth=20
Write-Host "Sheet 5 done"

# Save
$savePath="D:\Codex\王丽颖_减脂塑形方案_V3.xlsx"
$wb.SaveAs($savePath); $wb.Close(); $excel.Quit()
[GC]::Collect()
Write-Host "SAVED: $savePath"
Write-Host "DONE"
