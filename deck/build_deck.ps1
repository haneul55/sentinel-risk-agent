# 기획서 양식(pptx)에 내용을 채워 기획서 pptx / pdf 를 만드는 스크립트 (PowerPoint 필요)
# 테마: 네이비 + 블루 한 가지 강조색, 흰 카드, 넓은 여백
$ErrorActionPreference = "Stop"
# 공모전 공식 양식(pptx)은 저장소에 포함하지 않는다. 대회 페이지에서 받아 template\ 에 두고 실행한다.
$root = (Resolve-Path "$PSScriptRoot\..").Path
$src  = "$root\template\KIC_AI_에이전트_공모전_기획서_양식.pptx"
$out  = "$root\build\Sentinel_기획서.pptx"
$pdf  = "$root\build\Sentinel_기획서.pdf"
$img  = "$root\docs\images"
$prev = "$root\build\preview"
New-Item -ItemType Directory -Force "$root\build" | Out-Null
$FONT = "맑은 고딕"

function C([string]$hex) { [Convert]::ToInt32($hex.Substring(0,2),16) + 256*[Convert]::ToInt32($hex.Substring(2,2),16) + 65536*[Convert]::ToInt32($hex.Substring(4,2),16) }
$NAVY="14295A"; $BLUE="2F6FE4"; $TEXT="1F2937"; $MUTED="6B7280"; $SURF="F4F6FA"; $TINT="EEF3FD"; $BORDER="D9DEE7"; $WHITE="FFFFFF"

function StyleRange($tr, $size, $bold, $color) {
  $tr.Font.Name = $FONT; $tr.Font.NameFarEast = $FONT; $tr.Font.Italic = 0
  $tr.Font.Size = [single]$size
  $tr.Font.Bold = $(if ($bold) { -1 } else { 0 })
  $tr.Font.Color.RGB = [int](C $color)
}
function Box($s, $l, $t, $w, $h, $fill, $line=$null, $shadow=$false, $radius=0.08) {
  $sh = $s.Shapes.AddShape(5, $l, $t, $w, $h); $sh.Adjustments.Item(1) = [single]$radius
  $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = (C $fill)
  if ($line) { $sh.Line.Visible = -1; $sh.Line.ForeColor.RGB = (C $line); $sh.Line.Weight = 0.75 } else { $sh.Line.Visible = 0 }
  if ($shadow) { $sh.Shadow.Visible = -1; $sh.Shadow.OffsetX = 0; $sh.Shadow.OffsetY = 1.5; $sh.Shadow.Blur = 6; $sh.Shadow.ForeColor.RGB = (C $NAVY); $sh.Shadow.Transparency = 0.88 }
  else { $sh.Shadow.Visible = 0 }
  $sh.TextFrame.TextRange.Text = ""
  return $sh
}
# parts: @(text, size, bold, color) 의 배열. "`r" = 새 문단
function Rich($s, $l, $t, $w, $h, $parts, $anchor=1, $align=1, $space=2) {
  $sh = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $tf = $sh.TextFrame; $tf.WordWrap = -1; $tf.AutoSize = 0
  $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0; $tf.VerticalAnchor = $anchor
  if ($parts[0] -is [string]) { $parts = @(,$parts) }
  $text = ($parts | ForEach-Object { $_[0] }) -join ""
  $tr = $tf.TextRange; $tr.Text = $text
  $tr.ParagraphFormat.Alignment = $align; $tr.ParagraphFormat.SpaceAfter = $space; $tr.ParagraphFormat.SpaceBefore = 0
  $pos = 1
  foreach ($p in $parts) { $len = $p[0].Length; if ($len -gt 0) { StyleRange ($tr.Characters($pos, $len)) $p[1] $p[2] $p[3] }; $pos += $len }
  return $sh
}
function Head($s, $text) { Rich $s 76 130 807 22 @(,@($text, 15, $true, $NAVY)) 1 1 0 | Out-Null }
function Label($s, $l, $t, $w, $text) { Rich $s $l $t $w 16 @(,@($text, 11, $true, $NAVY)) 1 1 0 | Out-Null }
function Badge($s, $l, $t, $num, $d=18) {
  $o = $s.Shapes.AddShape(9, $l, $t, $d, $d); $o.Fill.Solid(); $o.Fill.ForeColor.RGB = (C $BLUE); $o.Line.Visible = 0; $o.Shadow.Visible = 0
  $tf = $o.TextFrame; $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0; $tf.WordWrap = 0
  $tr = $tf.TextRange; $tr.Text = [string]$num; StyleRange $tr 9 $true $WHITE; $tr.ParagraphFormat.Alignment = 2; $tf.VerticalAnchor = 3
}
# 흰 카드 + 번호 배지 + 제목/설명
function Card($s, $l, $t, $w, $h, $num, $title, $desc, $tsize=11, $dsize=10) {
  Box $s $l $t $w $h $WHITE $BORDER $true | Out-Null
  $pad = 12
  if ($num) { Badge $s ($l + $pad) ($t + 11) $num; $tx = $l + $pad + 26 } else { $tx = $l + $pad }
  Rich $s $tx ($t + 10) ($l + $w - $pad - $tx) ($h - 18) @(@("$title`r", $tsize, $true, $NAVY), @($desc, $dsize, $false, $TEXT)) 1 1 3 | Out-Null
}
function Table($s, $l, $t, $widths, $data, $size=10) {
  $rows = $data.Count; $cols = $widths.Count; $W = ($widths | Measure-Object -Sum).Sum
  $sh = $s.Shapes.AddTable($rows, $cols, $l, $t, $W, $rows*14)
  $tb = $sh.Table
  try { $tb.ApplyStyle("{2D5ABB26-0587-4C30-8999-92F81FD0307C}", $false) } catch {}
  for ($c=1; $c -le $cols; $c++) { $tb.Columns.Item($c).Width = $widths[$c-1] }
  for ($r=1; $r -le $rows; $r++) {
    for ($c=1; $c -le $cols; $c++) {
      $cell = $tb.Cell($r,$c); $cs = $cell.Shape
      $cs.TextFrame.MarginLeft=7; $cs.TextFrame.MarginRight=5; $cs.TextFrame.MarginTop=3.5; $cs.TextFrame.MarginBottom=3.5
      $cs.TextFrame.VerticalAnchor = 3
      $tr = $cs.TextFrame.TextRange; $tr.Text = [string]$data[$r-1][$c-1]
      $tr.ParagraphFormat.SpaceAfter = 0; $tr.ParagraphFormat.Alignment = 1
      $cs.Fill.Visible = -1; $cs.Fill.Solid()
      if ($r -eq 1) { StyleRange $tr $size $true $WHITE; $cs.Fill.ForeColor.RGB = (C $NAVY) }
      else {
        if ($c -eq 1) { StyleRange $tr $size $true $NAVY } else { StyleRange $tr $size $false $TEXT }
        $cs.Fill.ForeColor.RGB = (C $WHITE)
      }
      foreach ($b in 2,4) { $cell.Borders.Item($b).Visible = 0 }
      foreach ($b in 1,3) { $bd = $cell.Borders.Item($b); $bd.Visible = -1; $bd.ForeColor.RGB = (C $BORDER); $bd.Weight = 0.75 }
    }
    $tb.Rows.Item($r).Height = 10
  }
  return $sh
}
function Tight($sh, $m=2.3) { $tb = $sh.Table; for ($r=1; $r -le $tb.Rows.Count; $r++) { for ($c=1; $c -le $tb.Columns.Count; $c++) { $tf = $tb.Cell($r,$c).Shape.TextFrame; $tf.MarginTop = $m; $tf.MarginBottom = $m } ; $tb.Rows.Item($r).Height = 10 } }
function Arrow($s, $l, $t, $w=14, $h=9) {
  $a = $s.Shapes.AddShape(36, $l, $t, $w, $h); $a.Fill.Solid(); $a.Fill.ForeColor.RGB = (C "B8C2D3"); $a.Line.Visible = 0; $a.Shadow.Visible = 0
}
function Pic($s, $file, $l, $t, $w, $h) { $s.Shapes.AddPicture("$img\$file", 0, -1, $l, $t, $w, $h) | Out-Null }
function Bottom($sh) { $sh.Top + $sh.Height }
function ClearGuide($s) {
  $del = @(); foreach ($sh in $s.Shapes) { if ($sh.HasTextFrame -and $sh.TextFrame.TextRange.Text.StartsWith("[작성방법]")) { $del += $sh } }
  foreach ($d in $del) { $d.Delete() }
}

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($src, 0, 0, 0)
$report = @()
# 추가 본문 2쪽: 양식 4쪽(서비스 제안)·5쪽(Workflow)을 복제해 맨 뒤(11, 12)에 두고, 내용을 채운 뒤 제자리로 옮긴다
$pres.Slides.Item(4).Duplicate().MoveTo($pres.Slides.Count)
$pres.Slides.Item(5).Duplicate().MoveTo($pres.Slides.Count)
function Pill($s, $l, $t, $w, $text, $fill, $fg, $line=$null, $h=14, $size=8) {
  $p = Box $s $l $t $w $h $fill $line $false 0.5
  $tf = $p.TextFrame; $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0; $tf.WordWrap = 0; $tf.VerticalAnchor = 3
  $tr = $tf.TextRange; $tr.Text = $text; StyleRange $tr $size $true $fg; $tr.ParagraphFormat.Alignment = 2
}

# ---------- 1. 표지: 네이비 그라데이션 ----------
$s = $pres.Slides.Item(1)
$bg = $s.Shapes.AddShape(1, 0, 0, 960, 495); $bg.Line.Visible = 0; $bg.Shadow.Visible = 0
$bg.Fill.TwoColorGradient(3, 1); $bg.Fill.ForeColor.RGB = (C "0C1C40"); $bg.Fill.BackColor.RGB = (C "4F7CC4")
$bg.ZOrder(1)
foreach ($sh in $s.Shapes) {
  if ($sh.Name -eq "Google Shape;88;p1" -or $sh.Name -eq "TextBox 4") { $tr = $sh.TextFrame.TextRange; $tr.Font.Color.RGB = (C $WHITE); $tr.Font.Name = $FONT; $tr.Font.NameFarEast = $FONT }
  if ($sh.Name -eq "Google Shape;88;p1") { $sh.TextFrame.TextRange.Font.Bold = -1; $sh.TextFrame.TextRange.Font.Size = 30 }
  if ($sh.Name -eq "Google Shape;88;p1" -or $sh.Name -eq "Google Shape;94;p1" -or $sh.Name -eq "TextBox 4") { $sh.Top = $sh.Top - 48 }
  if ($sh.Name -eq "Google Shape;94;p1") { $sh.Line.ForeColor.RGB = (C "AFC3E6"); $sh.Line.Weight = 1 }
  if ($sh.Name -eq "Google Shape;90;p1" -and $sh.Top -lt 100) { $sh.Visible = 0 }
  if ($sh.HasTable) {
    $tb = $sh.Table
    $vals = @("하느으리야", "김하늘", "개인 참가")
    $sh.Top = 378
    for ($r=1; $r -le 3; $r++) { for ($c=1; $c -le 2; $c++) {
      $cell = $tb.Cell($r,$c); $cs = $cell.Shape; $cs.Fill.Visible = -1; $cs.Fill.Solid(); $cs.Fill.ForeColor.RGB = (C "1B3366"); $cs.Fill.Transparency = 0.35
      foreach ($b in 1,2,3,4) { $bd = $cell.Borders.Item($b); $bd.Visible = -1; $bd.ForeColor.RGB = (C "AFC3E6"); $bd.Weight = 0.75 }
      $tr = $cs.TextFrame.TextRange
      if ($c -eq 2) { $tr.Text = $vals[$r-1] }
      StyleRange $tr 12 ($c -eq 1) $WHITE; $tr.ParagraphFormat.Alignment = 2
      $cs.TextFrame.MarginTop = 2; $cs.TextFrame.MarginBottom = 2
      # (placeholder highlight removed)
    } }
  }
}
Rich $s 130 222 700 96 @(@("제안 AI 에이전트`r", 11, $true, "9FC3FF"), @("센티넬 ", 34, $true, $WHITE), @("Sentinel`r", 34, $true, "9FC3FF"), @("보유 자산 리스크 조기경보 에이전트", 13, $false, "D6E2F5")) 1 2 4 | Out-Null
foreach ($sh in $s.Shapes) { if ($sh.HasTable) { for ($r=1; $r -le 3; $r++) { $sh.Table.Rows.Item($r).Height = 24 } } }
$chips = @(@("99.5%", "주요 공시가 한국 업무시간 밖"), @("38.6억 달러", "하룻밤 연쇄 노출 실제 사례"), @("0원", "공개 데이터만 사용"))
for ($i=0; $i -lt 3; $i++) {
  $cx = 266 + $i*146
  $cb = Box $s $cx 316 136 42 "1B3366"
  $cb.Fill.Transparency = 0.3; $cb.Line.Visible = -1; $cb.Line.ForeColor.RGB = (C "AFC3E6"); $cb.Line.Weight = 0.75
  Rich $s $cx 321 136 36 @(@("$($chips[$i][0])`r", 17, $true, $WHITE), @($chips[$i][1], 9, $false, "D6E2F5")) 1 2 0 | Out-Null
}

# ---------- 2. Summary ----------
$s = $pres.Slides.Item(2)
$sum = @(
  @("센티넬(Sentinel)", ": 위험 조짐을 먼저 잡고, 사건이 KIC 포트폴리오 전체로 번지는 노출까지 계산해 근거와 함께 알려주는 야간 리스크 에이전트"),
  @("주요 공시의 99.5%가 한국 업무시간 밖에 나온다.", " KIC 상위 15종목의 1년 공시 28,688건 중 주요 사건 공시는 184건이고, 그중 183건이 평일 9~18시 밖에 접수됐다. 운용역과 리스크 담당자는 다음 날 아침에야 안다"),
  @("조짐 감지, 사건 조사, 연쇄 노출 계산.", " 위험 요인 문구가 바뀌면 조짐으로 잡고, 사건이 나면 에이전트가 도구를 골라 조사해 출처 두 곳 이상으로 검증한다. 이어 종목 간 관계를 따라 KIC 노출 금액을 합산해 알린다"),
  @("공개 데이터 6종을 쓴다.", " SEC EDGAR(13F, 공시, XBRL), GDELT·Finnhub(뉴스), 주가, 한국은행 ECOS(환율), 美재무부(금리). 구현은 Claude, LangGraph, PostgreSQL+pgvector. 공시에서 뽑은 관계 지도와 피드백을 판정에 쓴다"),
  @("1차로 볼 공시가 전체의 0.6%로 줄어든다.", " 경보는 사건 후 15분 이내를 목표로 하고, 모든 경보에 원문 근거를 붙인다. 놓친 사건과 오경보 비율은 MVP에서 직접 측정한다. 이후 677종목, 국내 자산(DART), 채권으로 넓힐 수 있다"))
foreach ($sh in $s.Shapes) { if ($sh.HasTable) { $tb = $sh.Table
  for ($r=1; $r -le 5; $r++) {
    $lc = $tb.Cell($r,1).Shape; StyleRange $lc.TextFrame.TextRange 13 $true $NAVY; $lc.Fill.Visible = -1; $lc.Fill.Solid(); $lc.Fill.ForeColor.RGB = (C $SURF)
    $tr = $tb.Cell($r,2).Shape.TextFrame.TextRange; $tr.Text = $sum[$r-1][0] + $sum[$r-1][1]
    StyleRange $tr 12 $false $TEXT; $tr.ParagraphFormat.Alignment = 1
    StyleRange ($tr.Characters(1, $sum[$r-1][0].Length)) 12 $true $BLUE
    if ($r -eq 1) {
      $lab = "실데이터 검증  "; $ext = "KIC 보유 15종목의 1년 공시로 확인: 주요 공시 99.5%가 한국 업무시간 밖, Apple CEO 교체 공시를 원문으로 재구성"
      $st = $tr.Text.Length + 2
      $tr.InsertAfter("`r" + $lab + $ext) | Out-Null
      StyleRange ($tr.Characters($st, $lab.Length)) 10.5 $true $NAVY
      StyleRange ($tr.Characters($st + $lab.Length, $ext.Length)) 10.5 $false $NAVY
      $tr.Paragraphs(2).ParagraphFormat.SpaceBefore = 5
    }
    $tb.Cell($r,2).Shape.TextFrame.MarginLeft = 14; $tb.Cell($r,2).Shape.TextFrame.MarginRight = 14
  }
} }

# ---------- 3. 문제 정의 ----------
$s = $pres.Slides.Item(3); ClearGuide $s
Head $s "주요 공시의 99.5%가 한국 업무시간 밖에 나온다"
Box $s 76 160 400 150 $WHITE $BORDER $true | Out-Null
Pic $s "3_공시_신호대잡음.png" 87 162 378 147.4
Box $s 486 160 397 150 $WHITE $BORDER $true | Out-Null
Pic $s "4_공시_시각.png" 498 162 372 147
Box $s 76 318 400 50 $SURF | Out-Null
Rich $s 88 324 380 42 @(@("주요 사용자`r", 10.5, $true, $NAVY), @("주 ", 9.5, $true, $BLUE), @("해외 주식 운용역·리서치 담당자   ", 9.5, $false, $TEXT), @("보조 ", 9.5, $true, $BLUE), @("리스크관리·보고서 담당자", 9.5, $false, $TEXT)) 1 1 2 | Out-Null
Box $s 486 318 397 50 $SURF | Out-Null
Rich $s 498 324 377 42 @(@("KIC 포트폴리오`r", 10.5, $true, $NAVY), @("총운용자산 2,320억 달러(2025년 말)`r13F 공개 미국 주식 677종목 527억 달러, 상위 15종목 비중 37%", 9, $false, $TEXT)) 1 1 2 | Out-Null
$t = Table $s 76 377 @(92, 330, 385) @(
  @("업무 시나리오", "지금", "센티넬 도입 후"),
  @("새벽 5~7시", "실적·경영진 교체 공시가 몰리지만 모두 잠든 시간", "감지·검증 후 연결 종목까지 합친 노출 금액 순으로 경보"),
  @("아침 9시", "사이트를 돌며 보유 종목과 하나씩 대조", "보유 종목 관련 항목만 담긴 브리핑 확인"),
  @("오전", "정리하느라 판단이 늦어짐", "원문 확인 후 대응 판단에 바로 집중")) 10
$report += "S3 bottom: " + (Bottom $t)

# ---------- 4. 서비스 제안 ----------
$s = $pres.Slides.Item(4); ClearGuide $s
Head $s "조짐은 먼저 잡고, 사건은 KIC 포트폴리오 전체 노출로 알려준다"
$cards = @(
  @("조짐 감지 (문구 변화)", "연간·분기 보고서 위험 요인에서 새로 생긴 문장 추출"),
  @("이중 감지", "보유 종목을 입력받아 숫자는 규칙, 글 위험은 AI로 감지"),
  @("자율 조사·검증", "에이전트가 더 찾아보고 출처 2곳 이상 교차 확인"),
  @("연쇄 노출 계산", "관계를 따라 KIC 노출 금액을 합산, 큰 순서로 경보"))
$y = 160
for ($i=0; $i -lt 4; $i++) { Card $s 76 $y 372 41 ($i+1) $cards[$i][0] $cards[$i][1] 10.5 9.5; $y += 45 }
$ex = Box $s 76 341 372 31 $SURF
Rich $s 86 343 356 28 @(
  @("조짐 실제 사례  ", 8.5, $true, $BLUE), @("Apple 10-K(2026.11.3) 위험 요인 387문장 중 새 문장 89개`r", 8.5, $false, $TEXT),
  @("`"expects these trends to intensify for memory and storage`" → 보유 종목 Micron 연결", 8.5, $false, $NAVY)) 3 1 0 | Out-Null
Box $s 458 160 425 212 $WHITE $BORDER $true | Out-Null
Pic $s "1_전체구조.png" 468 166.5 405 199
$t = Table $s 76 378 @(118, 300, 389) @(
  @("비교 대상", "한계", "센티넬의 차별점"),
  @("뉴스 키워드 알림", "종목명만 맞으면 모두 알림, 종목 간 연결은 못 봄", "보유 종목 간 관계를 따라 연쇄 노출 금액까지 계산"),
  @("블룸버그 알림", "사건 알림은 빠르지만 순위 판단·검증은 사람 몫", "KIC 노출 금액 순 정렬, 문구 변화로 조짐부터 포착"),
  @("챗봇·요약 LLM", "물어봐야 답하고, 숫자까지 AI가 판단", "묻기 전에 감시, 숫자는 규칙, 모든 경보에 원문 근거")) 10
$report += "S4 bottom: " + (Bottom $t)

# ---------- 5. Agent Workflow ----------
$s = $pres.Slides.Item(5); ClearGuide $s
Head $s "리스크 조사 에이전트: 위험 신호가 실제인지, 얼마나 심각한지 근거와 함께 판정한다"
Box $s 76 160 592 292 $WHITE $BORDER $true | Out-Null
Pic $s "2_에이전트_Workflow.png" 84 166 576 281
# 오른쪽: 출력 예시(경보 카드) + 심각도 기준
Box $s 680 160 203 134 $WHITE $BLUE $true | Out-Null
$pill = Box $s 692 169 34 15 $BLUE
$pill.TextFrame.MarginLeft=0; $pill.TextFrame.MarginRight=0; $pill.TextFrame.MarginTop=0; $pill.TextFrame.MarginBottom=0
$ptr = $pill.TextFrame.TextRange; $ptr.Text = "긴급"; StyleRange $ptr 8.5 $true $WHITE; $ptr.ParagraphFormat.Alignment = 2; $pill.TextFrame.VerticalAnchor = 3
Rich $s 731 169 145 15 @(,@("실제 공시로 재구성한 예시", 8.5, $false, $MUTED)) 3 1 0 | Out-Null
Rich $s 692 189 182 120 @(
  @("Apple · CEO 교체 공시`r", 10.5, $true, $NAVY),
  @("2026.4.21 06:29 접수 (한국 시간)`r", 8.5, $false, $MUTED),
  @("KIC 보유  ", 8.5, $true, $BLUE), @("30억 달러, 미국 주식 2위(5.7%)`r", 8.5, $false, $TEXT),
  @("내용  ", 8.5, $true, $BLUE), @("8-K Item 5.02, 9/1부로 CEO 교체`r", 8.5, $false, $TEXT),
  @("근거  ", 8.5, $true, $BLUE), @("`"Tim Cook will transition from his role as Chief Executive Officer to Executive Chair`"`r", 8.5, $false, $TEXT),
  @("판정  ", 8.5, $true, $BLUE), @("계획된 승계라 주의, 보유 2위라 긴급", 8.5, $false, $TEXT)) 1 1 4 | Out-Null
Box $s 680 300 203 40 $SURF | Out-Null
Rich $s 690 306 186 30 @(
  @("심각도  ", 8.5, $true, $NAVY), @("긴급 즉시 알림, 주의·참고는 브리핑`r", 8.5, $false, $TEXT),
  @("보유 비중이 크거나 연결 노출이 크면 한 단계 올림", 8, $false, $MUTED)) 1 1 1 | Out-Null
Box $s 680 346 203 106 $WHITE $BLUE $true | Out-Null
Rich $s 692 356 182 90 @(
  @("연쇄 노출 실제 사례`r", 10, $true, $NAVY),
  @("2026.4.3 새벽, 보유 2종목이 한 사람으로 연결`r", 8.5, $false, $MUTED),
  @("05:18  ", 8.5, $true, $BLUE), @("Alphabet 회계책임자 사임 공시`r", 8.5, $false, $TEXT),
  @("06:25  ", 8.5, $true, $BLUE), @("Broadcom, 같은 인물을 CFO로 선임`r", 8.5, $false, $TEXT),
  @("노출 합계  ", 8.5, $true, $BLUE), @("38.6억 달러 (미국 주식의 7.3%)", 8.5, $true, $NAVY)) 1 1 6 | Out-Null

# ---------- 6. 기술·Tool ----------
$s = $pres.Slides.Item(6); ClearGuide $s
Head $s "도구 8종을 직접 골라 쓰고, 틀린 답이 나가지 않도록 확인한다"
Label $s 76 160 480 "도구 (Tool / API)"
$t1 = Table $s 76 178 @(108, 230, 142) @(
  @("도구", "역할 · 입력 → 출력", "호출 조건"),
  @("search_filings", "SEC 공시 · 티커·기간 → 목록·본문", "공시 확인 필요 시"),
  @("get_financials", "XBRL 재무 · 티커·항목 → 분기 값", "실적 쇼크 의심"),
  @("search_news", "뉴스 · 검색어·기간 → 제목·링크", "원인 조사·검증"),
  @("get_price", "주가 · 티커 → 현재가·등락률", "시장 반응 확인"),
  @("get_fx / get_rate", "환율·금리 · 지표 → 시계열", "거시 요인 확인"),
  @("recall_memory", "과거 사건·피드백 → 유사 사례", "판정 직전"),
  @("get_relations", "관계 지도 · 티커 → 연결 종목·관계", "연쇄 노출 계산"),
  @("diff_risk_factors", "보고서 비교 · 티커 → 새 위험 문장", "10-K·10-Q 제출 시")) 9.5
Tight $t1; $y2 = (Bottom $t1) + 10
Label $s 76 $y2 480 "AI 기술 구성"
$t2 = Table $s 76 ($y2 + 18) @(108, 180, 192) @(
  @("구분", "선택", "이유"),
  @("LLM", "Claude Sonnet 5.5 · Haiku 5.5", "조사는 큰 모델, 1차 분류는 작은 모델"),
  @("Framework", "LangGraph", "단계·반복·종료 조건 제어"),
  @("RAG", "PostgreSQL + pgvector", "같은 회사의 비슷한 과거 사건 검색"),
  @("Memory", "피드백·판정 이력 테이블", "'관련 없음' 반복 유형은 심각도↓")) 9.5
Tight $t2; $report += "S6 left bottom: " + (Bottom $t2)
$hR = (Bottom $t2) - 160
Box $s 570 160 313 $hR $SURF | Out-Null
$rel = @(@("환각 차단", "AI가 든 근거 문장을 원문과 대조, 없으면 폐기"), @("출처 표시", "모든 항목에 원문 링크와 근거 문장 필수"), @("규칙·AI 분리", "숫자는 규칙이 판정, AI는 숫자를 만들지 않음"), @("구조화 출력", "정해진 JSON 형식 검사, 어긋나면 재요청"), @("재시도·예외", "API 한도 초과 시 재시도, 실패한 출처는 브리핑에 표시"))
Rich $s 584 170 285 18 @(,@("신뢰성 확보 방안", 11, $true, $NAVY)) 1 1 0 | Out-Null
$yy = 196; $step = [math]::Floor(($hR - 44) / 5)
for ($i=0; $i -lt 5; $i++) {
  Badge $s 584 $yy ($i+1)
  Rich $s 610 ($yy + 1) 262 ($step - 4) @(@("$($rel[$i][0])`r", 10.5, $true, $NAVY), @($rel[$i][1], 10, $false, $TEXT)) 1 1 1 | Out-Null
  $yy += $step
}

# ---------- 7. 데이터 ----------
$s = $pres.Slides.Item(7); ClearGuide $s
Head $s "공개 데이터 6종을 모아 보유 종목 기준으로 정리한다"
Label $s 76 160 480 "활용 데이터와 출처"
$t1 = Table $s 76 178 @(66, 132, 92, 190) @(
  @("데이터", "출처", "갱신", "활용 목적"),
  @("보유 종목", "SEC 13F (KIC)", "분기", "감시 대상, 보유 비중으로 경보 순서"),
  @("공시·재무", "SEC EDGAR API", "15분", "사건 감지, 위험 요인 문구 비교"),
  @("뉴스", "GDELT, Finnhub", "15분", "글 위험 감지, 교차 검증"),
  @("주가", "Finnhub, Alpha Vantage", "매일·장중", "주가 급락"),
  @("환율", "한국은행 ECOS", "매일", "원/달러 급변"),
  @("금리", "미 재무부 (Treasury)", "매일", "금리 급변")) 9.5
Tight $t1; $y2 = (Bottom $t1) + 10
Label $s 76 $y2 480 "리스크 이벤트 6종 ↔ 데이터"
$t2 = Table $s 76 ($y2 + 18) @(92, 118, 270) @(
  @("이벤트", "감지 방식", "근거 데이터"),
  @("주가 급락", "규칙 (하루 −7%↓)", "급락 후 뉴스와 공시로 원인 조사"),
  @("환율 급변", "규칙 (하루 ±1.5%)", "ECOS 원/달러, 미 재무부 금리"),
  @("실적 쇼크", "규칙 + AI", "8-K Item 2.02, XBRL 재무(전년 대비)"),
  @("소송", "AI", "8-K Item 8.01, 10-Q 법률 항목, 뉴스"),
  @("신용등급 하향", "AI", "뉴스·공시 글에서 탐지 (등급 데이터는 유료)"),
  @("위험 요인 변화", "AI (직전 보고서 비교)", "10-K Item 1A, 10-Q 위험 요인의 새 문장")) 9.5
Tight $t2; $report += "S7 left bottom: " + (Bottom $t2)
Label $s 570 160 313 "전처리 파이프라인"
$steps = @(@("원본 보관 (Raw)", "수집한 그대로 저장해 나중에 다시 확인"), @("정제 (Clean)", "한국 시간 통일, 중복 뉴스 제거, 잡음 공시 제외"), @("종목 연결 (Link)", "별칭 사전으로 태그, 공시에서 종목 간 관계 추출"), @("분석용 (Mart)", "종목별 이벤트와 지표를 날짜순으로 정리"))
$y = 180; $gap = 15; $ch = [math]::Floor(((Bottom $t2) - 180 - 3*$gap) / 4)
for ($i=0; $i -lt 4; $i++) {
  Card $s 570 $y 313 $ch ($i+1) $steps[$i][0] $steps[$i][1] 10.5 10
  if ($i -lt 3) { Arrow $s 719.5 ($y + $ch + 3) }
  $y += $ch + $gap
}

# ---------- 8. MVP ----------
$s = $pres.Slides.Item(8); ClearGuide $s
Head $s "19일 동안 15종목, 위험 6종에 대해 수집부터 브리핑까지 완성한다"
$tops = @(
  @("MVP 범위 · 핵심 기능", "KIC 상위 15종목, 위험 6종, 출처 6곳`v조짐·사건 감지, 조사 에이전트, 연쇄 노출"),
  @("개발 환경 · 기술 스택", "Python, PostgreSQL+pgvector, LangGraph,`vClaude API, Streamlit, GitHub Actions(15분 수집)"),
  @("데모 시나리오", "과거 실제 공시(예: 경영진 교체)를 다시 넣어`v감지부터 피드백까지 시연, 영상도 제출"))
$x = 76
foreach ($tp in $tops) {
  Box $s $x 160 262 56 $SURF | Out-Null
  Rich $s ($x + 12) 166 240 48 @(@("$($tp[0])`r", 10.5, $true, $NAVY), @($tp[1], 9.5, $false, $TEXT)) 1 1 2 | Out-Null
  $x += 272.5
}
$t1 = Table $s 76 226 @(140, 345, 322) @(
  @("기간", "할 일", "완료 기준"),
  @("11/20 ~ 12/2 사전 준비", "수집기 시제품, DB 설계", "출처 6곳 수집 성공"),
  @("12/3 ~ 12/9", "수집·정제·종목 연결 파이프라인", "15종목 데이터 매일 자동 적재"),
  @("12/10 ~ 12/16", "규칙 엔진, AI 분류, 조사 에이전트, 관계 지도", "과거 실제 사건 재현 테스트 통과"),
  @("12/17 ~ 12/22", "경보·브리핑 화면, 피드백, 배포, 시연 영상", "데모 URL 작동, GitHub 실행 가이드")) 10
$y2 = (Bottom $t1) + 14
$t2 = Table $s 430 $y2 @(128, 325) @(
  @("주요 리스크", "대응"),
  @("개인 참가, 짧은 기간", "범위 고정(15종목·6종), 주차별 완료 기준"),
  @("AI 오판·환각", "근거 문장 대조, 숫자는 규칙, 라벨 100건 측정"),
  @("무료 API 한도", "수집 주기 조절, 캐시, 재시도"),
  @("LLM 비용", "1차 분류는 작은 모델, 조사는 신호 난 건만")) 10
$hD = (Bottom $t2) - $y2
Box $s 76 $y2 342 $hD $WHITE $BLUE $true | Out-Null
Rich $s 90 ($y2 + 12) 316 ($hD - 16) @(
  @("실현 가능성 근거`r", 11, $true, $NAVY),
  @("데이터 수집 검증  ", 9.5, $true, $BLUE), @("15종목 1년 공시 집계, 13F 파싱 완료`r", 9.5, $false, $TEXT),
  @("무료 한도 안 운영  ", 9.5, $true, $BLUE), @("SEC 15분 15회, Alpha Vantage 하루 15회`r", 9.5, $false, $TEXT),
  @("AI 호출은 소수  ", 9.5, $true, $BLUE), @("주요 공시 하루 1건 미만, 신호 건만 조사", 9.5, $false, $TEXT)) 1 1 10 | Out-Null
$report += "S8 bottom: " + (Bottom $t2)

# ---------- 9. 기대효과 ----------
$s = $pres.Slides.Item(9); ClearGuide $s
Head $s "볼 공시는 줄고, 위험은 더 빨리 알게 된다"
$stats = @(@("15분 이내", "시간", "기존: 다음 날 아침"), @("0.6%", "검토량", "기존: 전체 직접 확인"), @("100%", "신뢰성", "모든 경보에 원문 근거"), @("0원", "데이터 비용", "공개 API만 사용"))
$x = 76
foreach ($st in $stats) {
  Box $s $x 160 118 84 $WHITE $BORDER $true | Out-Null
  Rich $s ($x + 4) 176 110 64 @(@("$($st[0])`r", 19, $true, $BLUE), @("$($st[1])`r", 10, $true, $NAVY), @($st[2], 8.5, $false, $MUTED)) 1 2 1 | Out-Null
  $x += 127.3
}
Label $s 76 260 500 "성능 평가 계획 (MVP에서 직접 측정해 본선에서 보고)"
$t1 = Table $s 76 278 @(82, 268, 150) @(
  @("지표", "측정 방법", "목표"),
  @("놓친 사건 비율", "과거 실제 사건 30건을 다시 흘려 감지 여부 확인", "10% 이하"),
  @("오경보율", "경보 100건에 '관련 있음/없음' 직접 라벨링", "20% 이하"),
  @("근거 일치율", "근거 문장이 원문에 그대로 있는지 자동 대조", "100%"),
  @("경보 지연", "공시 접수 시각과 경보 발송 시각의 차이", "15분 이내"),
  @("관계 정확도", "관계 지도 표본 50건을 원문과 직접 대조", "90% 이상"),
  @("유용성", "사용자 피드백 중 '도움됨' 비율", "70% 이상")) 10
Rich $s 76 ((Bottom $t1) + 8) 500 16 @(,@("부가 효과: 아침 브리핑이 일일 리스크 보고서 초안 역할을 한다", 9, $false, $MUTED)) 1 1 0 | Out-Null
$report += "S9 left bottom: " + ((Bottom $t1) + 24)
Label $s 590 160 293 "확장성 · 운영 적용 경로"
$ext = @(@("범위", "15종목에서 677종목, 이후 사내 비공개 자산까지"), @("지역·자산군", "국내 DART 공시, 채권 신용 스프레드"), @("보안", "공개 13F로 개발, 사내 적용 시 종목 가명 처리"), @("운영", "한 팀 시범 운영 후 기준을 조정해 전사로 확대"))
$endY = (Bottom $t1) + 24
$y = 180; $gap = 15; $ch = [math]::Floor(($endY - 180 - 3*$gap) / 4)
for ($i=0; $i -lt 4; $i++) {
  Card $s 590 $y 293 $ch ($i+1) $ext[$i][0] $ext[$i][1] 10.5 10
  if ($i -lt 3) { Arrow $s 729.5 ($y + $ch + 3) }
  $y += $ch + $gap
}

# ---------- 10. 부록 ----------
$s = $pres.Slides.Item(10); ClearGuide $s
$t = Table $s 76 134 @(95, 225, 377, 110) @(
  @("구분", "이름", "출처 · 이용 조건", "수집일"),
  @("데이터", "SEC EDGAR (13F, 공시, XBRL)", "sec.gov · 무료 재사용, 초당 10회 이하, User-Agent 표시", "2026.11.6~11.11"),
  @("데이터", "GDELT DOC 2.0", "gdeltproject.org · 무료, GDELT 출처 표시·링크", "MVP 기간 수집"),
  @("데이터", "Finnhub", "finnhub.io · 무료 요금제 범위 내 사용", "MVP 기간 수집"),
  @("데이터", "Alpha Vantage", "alphavantage.co · 무료 키 (하루 25회)", "MVP 기간 수집"),
  @("데이터", "한국은행 ECOS", "ecos.bok.or.kr · 무료 키, 출처 표시 (통계표 731Y001)", "MVP 기간 수집"),
  @("데이터", "미 재무부 국채 금리", "home.treasury.gov · 미국 정부 공개 데이터, 출처 표시", "MVP 기간 수집"),
  @("참고 자료", "Lazy Prices (보고서 문구 변화 연구)", "Cohen, Malloy, Nguyen (2020), Journal of Finance 75(3)", "-"),
  @("참고 자료", "KIC 운용자산 보도", "헤럴드경제(2026.2.22), AsiaAsset. 2025년 말 운용자산 2,320억 달러", "-"),
  @("도표", "그림 (구조도, Workflow, 공시 집계·시각)", "본인 작성 · 공시 그래프는 SEC EDGAR 데이터를 직접 집계", "-"),
  @("모델 (계획)", "Claude Sonnet 5.5, Haiku 5.5", "Anthropic API · 유료, 이용약관 준수", "-"),
  @("오픈소스 (계획)", "LangGraph, Streamlit, PostgreSQL, pgvector", "각 오픈소스 라이선스 준수", "-"),
  @("AI 도구", "Claude (Anthropic)", "자료 정리, 집계·문서 작성 보조", "-")) 9.5
Rich $s 76 ((Bottom $t) + 10) 807 50 @(
  @("· 문제 정의의 공시 건수는 SEC EDGAR Submissions API로 직접 집계 (2025.11.6 ~ 2026.11.6, KIC 13F 상위 15종목)`r", 9.5, $false, $MUTED),
  @("· 공시 접수 시각은 SEC 접수 시각(UTC)을 한국 시간으로 변환해 집계`r", 9.5, $false, $MUTED),
  @("· 4쪽 조짐 사례는 Apple 10-K 두 해(2025.10.31, 2026.11.3)의 Item 1A를 문장 단위로 비교`r", 9.5, $false, $MUTED),
  @("· 5~7쪽 화면·경보·사례는 Apple(2026.4.20), Alphabet·Broadcom·Tesla(2026.4.2) 8-K 원문과 13F 보유 금액으로 재구성`r", 9.5, $false, $MUTED),
  @("· 수집일은 원본을 내려받은 날짜. 13F는 2026.6.30 기준 보고서(2026.8.12 제출), 8-K·10-K는 2026.11.11 수집`r", 9.5, $false, $MUTED),
  @("· 본 기획서에 민감정보·개인정보·API Key가 포함되지 않았음을 확인함", 9.5, $false, $MUTED)) 1 1 2 | Out-Null
$report += "S10 bottom: " + ((Bottom $t) + 44)

# ---------- 추가 A (5쪽). 서비스 제안: 사용자 화면 ----------
$s = $pres.Slides.Item(11); ClearGuide $s
Head $s "아침 9시, 담당자는 이 화면 하나로 밤사이 위험을 확인한다"
Box $s 76 158 807 258 $WHITE $BORDER $true 0 | Out-Null
Box $s 76 158 807 24 $NAVY $null $false 0 | Out-Null
foreach ($k in 0,1,2) { $d = $s.Shapes.AddShape(9, (88 + $k*11), 166.5, 7, 7); $d.Fill.Solid(); $d.Fill.ForeColor.RGB = (C "AFC3E6"); $d.Line.Visible = 0; $d.Shadow.Visible = 0 }
Rich $s 128 158 400 24 @(@("센티넬 · 아침 브리핑   ", 9.5, $true, $WHITE), @("2026.4.3 (금) 09:00, 한국 시간", 8.5, $false, "D6E2F5")) 3 1 0 | Out-Null
Rich $s 560 158 313 24 @(,@("실제 공시(2026.4.2~3 밤)로 재구성한 화면 예시", 8.5, $false, "D6E2F5")) 3 3 0 | Out-Null
Box $s 76 182 807 26 $SURF $null $false 0 | Out-Null
Rich $s 88 182 785 26 @(
  @("밤사이 보유 종목 주요 공시 ", 9.5, $false, $TEXT), @("3건", 9.5, $true, $NAVY), @("  →  경보 ", 9.5, $false, $TEXT), @("2건", 9.5, $true, $BLUE),
  @("   (긴급 0 · 주의 1 · 참고 1)", 9, $false, $MUTED), @("                                        정렬: KIC 노출 금액 순", 9, $false, $MUTED)) 3 1 0 | Out-Null
# 왼쪽: 경보 목록
Box $s 86 216 334 66 $TINT $BLUE $false 0.06 | Out-Null
Box $s 86 216 4 66 $BLUE $null $false 0 | Out-Null
Pill $s 98 225 30 "주의" "FFF1D6" "B45309"
Pill $s 132 225 30 "연쇄" $WHITE $BLUE $BLUE
Rich $s 98 245 220 32 @(@("Alphabet → Broadcom 회계책임자 이동`r", 10, $true, $NAVY), @("공시 2건 묶음 · 05:18, 06:25 접수", 8.5, $false, $MUTED)) 1 1 2 | Out-Null
Rich $s 310 224 100 30 @(@("38.6억 달러`r", 11, $true, $NAVY), @("미국 주식의 7.3%", 8.5, $false, $MUTED)) 1 3 1 | Out-Null
Box $s 86 290 334 66 $WHITE $BORDER $false 0.06 | Out-Null
Pill $s 98 299 30 "참고" $SURF $MUTED $BORDER
Rich $s 98 319 220 32 @(@("Tesla 1분기 인도 35.8만 대 발표`r", 10, $true, $NAVY), @("8-K Item 2.02 · 4/2 22:07 접수", 8.5, $false, $MUTED)) 1 1 2 | Out-Null
Rich $s 310 298 100 30 @(@("8.4억 달러`r", 11, $true, $NAVY), @("미국 주식의 1.6%", 8.5, $false, $MUTED)) 1 3 1 | Out-Null
Rich $s 88 366 330 42 @(
  @("· 공시 3건을 사건 기준으로 묶어 경보 2건으로 표시`r", 8.5, $false, $MUTED),
  @("· 긴급은 즉시 메신저 알림, 주의·참고는 이 브리핑에 정리`r", 8.5, $false, $MUTED),
  @("· 경보가 없는 종목은 표시하지 않음", 8.5, $false, $MUTED)) 1 1 2 | Out-Null
# 오른쪽: 선택한 경보의 상세
Box $s 432 216 441 192 $WHITE $BORDER $false 0.03 | Out-Null
Rich $s 446 221 320 24 @(,@("Alphabet 회계책임자, Broadcom CFO로 이동", 11, $true, $NAVY)) 3 1 0 | Out-Null
Pill $s 789 226 72 "근거 2건 확인" $TINT $BLUE
Rich $s 446 250 415 80 @(
  @("무슨 일  ", 9, $true, $BLUE), @("Alphabet 회계책임자가 사임(4/9부), Broadcom이 같은 인물을 CFO로 선임(6/12부)`r", 9, $false, $TEXT),
  @("근거 ①  ", 9, $true, $BLUE), @("Alphabet 8-K: `"did not result from any disagreement with the Company`"`r", 9, $false, $TEXT),
  @("근거 ②  ", 9, $true, $BLUE), @("Broadcom 8-K: `"has served as ... Chief Accounting Officer of Alphabet Inc.`"`r", 9, $false, $TEXT),
  @("판정 이유  ", 9, $true, $BLUE), @("불화 없는 이직이라 단독으론 참고. 보유 2종목이 한 사람으로 이어져 한 단계 올려 주의", 9, $false, $TEXT)) 1 1 3 | Out-Null
Rich $s 446 321 60 22 @(,@("KIC 노출", 9, $true, $NAVY)) 3 1 0 | Out-Null
$ba = Box $s 506 325 246 14 $BLUE $null $false 0
$ba.TextFrame.MarginLeft = 6; $ba.TextFrame.MarginTop = 0; $ba.TextFrame.MarginBottom = 0; $ba.TextFrame.TextRange.Text = "Alphabet 26.5억 (5.0%)"; StyleRange $ba.TextFrame.TextRange 8 $true $WHITE; $ba.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $ba.TextFrame.VerticalAnchor = 3
$bb = Box $s 752 325 113 14 "A9C4F5" $null $false 0
$bb.TextFrame.MarginLeft = 5; $bb.TextFrame.MarginRight = 0; $bb.TextFrame.MarginTop = 0; $bb.TextFrame.MarginBottom = 0; $bb.TextFrame.WordWrap = 0; $bb.TextFrame.TextRange.Text = "Broadcom 12.1억 (2.3%)"; StyleRange $bb.TextFrame.TextRange 8 $true $NAVY; $bb.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $bb.TextFrame.VerticalAnchor = 3
Pill $s 446 377 58 "도움됨" $BLUE $WHITE $null 20 8.5
Pill $s 510 377 64 "관련 없음" $WHITE $TEXT $BORDER 20 8.5
Pill $s 580 377 72 "원문 보기 ↗" $WHITE $TEXT $BORDER 20 8.5
Rich $s 662 377 200 20 @(,@("피드백은 다음 판정에 반영된다", 8.5, $false, $MUTED)) 3 3 0 | Out-Null
# 아래: 화면이 주는 것 3가지
$feat = @(@("노출 금액 순 정렬", "비중이 큰 위험부터 본다"), @("사건 단위로 묶음", "여러 공시를 한 경보로"), @("클릭 한 번 피드백", "오경보 유형은 다음부터 낮춤"))
for ($i=0; $i -lt 3; $i++) {
  $fx = 76 + $i*272.5
  Box $s $fx 424 262 30 $SURF | Out-Null
  Badge $s ($fx + 10) 430 ($i+1)
  Rich $s ($fx + 36) 424 220 30 @(@("$($feat[$i][0])  ", 10, $true, $NAVY), @($feat[$i][1], 9, $false, $TEXT)) 3 1 0 | Out-Null
}

# ---------- 추가 B (7쪽). Agent Workflow: 실제 사례 재현 ----------
$s = $pres.Slides.Item(12); ClearGuide $s
Head $s "실제 사례 재현: 두 회사의 공시를 한 사람으로 이어 하나의 경보로 만든다"
$t = Table $s 76 160 @(46, 112, 382) @(
  @("시각", "단계 · 도구", "에이전트가 한 일과 판단"),
  @("05:18", "트리거", "Alphabet 8-K Item 5.02 접수. 경영진 변동 항목이라 규칙이 조사를 시작"),
  @("", "① 조사 계획", "확인할 것: 사임 사유, 회사와의 불화 여부, 후임과 이직처"),
  @("", "② search_filings", "원문에서 불화 없음 확인. 이직처는 공시에 없어 '참고·확인 필요'로 대기"),
  @("06:25", "트리거", "Broadcom 8-K Item 5.02 접수. CFO 은퇴, 후임으로 Amie Thuener 선임(6/12부)"),
  @("", "get_relations", "'Amie Thuener O'Toole'과 'Amie Thuener'를 직책·이력으로 같은 인물로 판정`v→ 관계 지도에 Alphabet–Broadcom 연결"),
  @("", "recall_memory", "05:18에 대기시킨 건을 불러와 하나의 사건으로 묶음"),
  @("", "③ 근거 검증", "두 회사의 8-K 원문 두 곳에서 근거 문장 인용 확인 → 충분"),
  @("", "④ 연쇄 노출 · 판정", "13F 기준 Alphabet 26.5억 + Broadcom 12.1억 = 38.6억 달러(미국 주식의 7.3%). 단독이면 참고, 연결 노출이 커 한 단계 올려 주의"),
  @("09:00", "아침 브리핑", "경보 목록 맨 위에 표시 (5쪽 화면)")) 10
Tight $t 6.0
$tb = $t.Table
for ($r=2; $r -le $tb.Rows.Count; $r++) {
  $tb.Cell($r,1).Shape.TextFrame.TextRange.Font.Color.RGB = (C $BLUE)
  $tb.Cell($r,2).Shape.TextFrame.TextRange.Font.Bold = -1; $tb.Cell($r,2).Shape.TextFrame.TextRange.Font.Color.RGB = (C $NAVY)
  if ($r -eq 2 -or $r -eq 5) { for ($c=1; $c -le 3; $c++) { $tb.Cell($r,$c).Shape.Fill.ForeColor.RGB = (C $SURF) } }
  if ($r -eq 6 -or $r -eq 9) { for ($c=1; $c -le 3; $c++) { $tb.Cell($r,$c).Shape.Fill.ForeColor.RGB = (C $TINT) } }
}
$tbot = Bottom $t
Rich $s 76 ($tbot + 6) 540 14 @(,@("공시 접수 시각·내용·보유 금액은 실제 자료, 분 단위 처리 순서는 설계대로 재현한 것", 8.5, $false, $MUTED)) 1 1 0 | Out-Null
$report += "SB table bottom: " + $tbot
$endB = $tbot + 20
Box $s 628 160 255 136 $WHITE $BLUE $true | Out-Null
Rich $s 642 171 228 120 @(
  @("에이전트가 스스로 내린 판단`r", 11, $true, $NAVY),
  @("1  ", 10, $true, $BLUE), @("이름이 달라도 직책·이력으로 같은 사람인지 판단`r", 9.5, $false, $TEXT),
  @("2  ", 10, $true, $BLUE), @("한 시간 간격의 두 공시를 하나의 사건으로 묶음`r", 9.5, $false, $TEXT),
  @("3  ", 10, $true, $BLUE), @("연결 노출이 크면 심각도를 한 단계 올림", 9.5, $false, $TEXT)) 1 1 12 | Out-Null
$hb = [math]::Floor(($endB - 304 - 8) / 2)
Box $s 628 304 255 $hb $SURF | Out-Null
Rich $s 642 314 228 ($hb - 16) @(@("키워드 알림이라면`r", 10, $true, $NAVY), @("Alphabet, Broadcom 알림이 따로 2건 온다.`v둘의 연결은 사람이 아침에 알아채야 한다", 9, $false, $TEXT)) 1 1 3 | Out-Null
Box $s 628 (312 + $hb) 255 $hb $SURF | Out-Null
Rich $s 642 (322 + $hb) 228 ($hb - 16) @(@("숫자는 규칙으로`r", 10, $true, $NAVY), @("노출 합산은 13F 보유 금액으로 계산하고,`vAI는 숫자를 만들지 않는다", 9, $false, $TEXT)) 1 1 3 | Out-Null

# 제자리로: A → 5쪽, B → 7쪽 (Workflow 다음)
$pres.Slides.Item(11).MoveTo(5)
$pres.Slides.Item(12).MoveTo(7)

# ---------- 정렬 마무리: 박스 안 글상자 하나만 있는 경우, 박스 안쪽에 꽉 맞추고 가로·세로 가운데 정렬 ----------
$centered = 0
foreach ($sl in $pres.Slides) {
  if ($sl.SlideIndex -eq 1) { continue }
  $all = @(); foreach ($sh in $sl.Shapes) { $all += $sh }
  foreach ($bx in $all) {
    if ($bx.Type -ne 1 -or $bx.AutoShapeType -ne 5) { continue }
    if ($bx.Width -lt 60 -or $bx.Height -lt 24) { continue }
    $inside = @($all | Where-Object { $_.Name -ne $bx.Name -and $_.Left -ge $bx.Left - 1 -and $_.Top -ge $bx.Top - 1 -and ($_.Left + $_.Width) -le ($bx.Left + $bx.Width + 1) -and ($_.Top + $_.Height) -le ($bx.Top + $bx.Height + 1) })
    if ($inside.Count -ne 1 -or $inside[0].Type -ne 17) { continue }
    $tx = $inside[0]; $padX = $(if ($bx.Width -lt 260) { 10 } else { 14 }); $padY = $(if ($bx.Height -lt 50) { 3 } else { 6 })
    $tx.Left = $bx.Left + $padX; $tx.Width = $bx.Width - 2*$padX
    $tx.Top = $bx.Top + $padY; $tx.Height = $bx.Height - 2*$padY
    $tx.TextFrame.VerticalAnchor = 3; $tx.TextFrame.WordWrap = -1
    $tx.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    if ($tx.TextFrame.TextRange.BoundHeight -gt $tx.Height + 2) { $report += "넘침: 쪽 $($sl.SlideIndex) '$($tx.TextFrame.TextRange.Text.Substring(0, [math]::Min(15, $tx.TextFrame.TextRange.Text.Length)))'" }
    $centered++
  }
}
$report += "가운데 정렬한 박스: $centered"

# ---------- 저장 · PDF · 미리보기 ----------
$pres.SaveAs($out)
$pres.SaveAs($pdf, 32)
New-Item -ItemType Directory -Force $prev | Out-Null
foreach ($sl in $pres.Slides) { $sl.Export("$prev\slide$($sl.SlideIndex).png", "PNG", 1920, 1080) }
$pres.Close()
if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
$report
