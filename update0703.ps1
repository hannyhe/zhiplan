$ErrorActionPreference="Stop"
$c1=[char]0x81F4+[char]0x77E5+[char]0x8BA1+[char]0x5212+[char]0x5185+[char]0x5BB9+[char]0x6536+[char]0x76CA
$c2=[char]0x81F4+[char]0x77E5+[char]0x6536+[char]0x76CA
$srcPath="G:\me\"+$c2+"\2025\"+$c1+".xls"
$dstPath="G:\me\"+$c2+"\2025\test\2607.xlsx"

function Levenshtein($s,$t){
 $m=$s.Length;$n=$t.Length;$w=$n+1
 $d=[int[]]::new(($m+1)*$w)
 for($i=0;$i-le$m;$i++){$d[$i*$w+0]=$i}
 for($j=0;$j-le$n;$j++){$d[0*$w+$j]=$j}
 for($j=1;$j-le$n;$j++){for($i=1;$i-le$m;$i++){
  $cc=if($s[$i-1]-eq$t[$j-1]){0}else{1}
  $a=$d[($i-1)*$w+$j]+1;$b=$d[$i*$w+($j-1)]+1;$c=$d[($i-1)*$w+($j-1)]+$cc
  $d[$i*$w+$j]=if($a-lt$b){if($a-lt$c){$a}else{$c}}else{if($b-lt$c){$b}else{$c}}
 }}
 return $d[$m*$w+$n]
}
function FuzzyRatio($s,$t){
 if([string]::IsNullOrEmpty($s)-or[string]::IsNullOrEmpty($t)){return 0}
 $d=Levenshtein $s $t;$max=[Math]::Max($s.Length,$t.Length)
 if($max-eq0){return 1}
 return 1-$d/$max
}

$ex=New-Object -ComObject Excel.Application
$ex.Visible=$false;$ex.DisplayAlerts=$false;$ex.ScreenUpdating=$false;$ex.EnableEvents=$false

# Read source .xls
$wb=$ex.Workbooks.Open($srcPath)
$s=$wb.Sheets(1);$sr=$s.UsedRange.Rows.Count;$raw=$s.UsedRange.Value2
$h1=$raw[1,1];$h2=$raw[1,2];$h3=$raw[1,3];$h4=$raw[1,4];$h5=$raw[1,5];$h6=$raw[1,6];$h7=$raw[1,7]
$data=@()
for($ri=2;$ri-le$sr;$ri++){
 $q=[string]$raw[$ri,1];if([string]::IsNullOrWhiteSpace($q)){continue}
 $tp=[string]$raw[$ri,2];$dv=$raw[$ri,3]
 if($dv-ne$null){if($dv.GetType().Name-eq"Double"){$ds=[datetime]::FromOADate([double]$dv).ToString("yyyy/M/d")}else{$ds=$dv.ToString()}}else{$ds=""}
 $va1=[double]($raw[$ri,4]);$va2=[double]($raw[$ri,5]);$va3=0;try{$va3=[double]($raw[$ri,6])}catch{};$va4=0;try{$va4=[double]($raw[$ri,7])}catch{}
 $data+=@{Q=$q;T=$tp;Dt=$ds;A=$va1;B=$va2;C=$va3;Gv=$va4}
}
$wb.Saved=$true;$wb.Close()
Write-Output ($data.Count.ToString()+" source rows")

# Open destination
$w=$ex.Workbooks.Open($dstPath)

# Find prev sheet (numeric name, highest)
$maxS=0;$prev=""
for($i=1;$i-le$w.Sheets.Count;$i++){
 $sn=$w.Sheets($i).Name;if($sn-match"^\d{4}$"-and[int]$sn-gt$maxS){$maxS=[int]$sn;$prev=$sn}
}
$today="{0:D4}"-f($maxS+1)
Write-Output ("prev=$prev today=$today")

# Delete if exists
try{$d=$w.Sheets($today);$d.Delete()}catch{}

# Create new sheet
$n=$w.Worksheets.Add()
$n.Name=$today;$n.Move([Type]::Missing,$w.Sheets($w.Sheets.Count))

# Headers (fixed labels matching 2505 structure)
$qd=[string][char]0x524D+[string][char]0x65E5
$h10=$qd+[string][char]0x5DEE
$h11=[string][char]0x6536+[string][char]0x76CA+[string][char]0x5DEE
$h12="L"+$today
$h13=$qd+"M";$h14=$qd+"N";$h15=$qd+"O"
$hh=@($h1,$h2,$h3,$h4,$h5,$h6,$h7,"E/D","G/F",$h10,$h11,$h12,$h13,$h14,$h15)
for($xi=0;$xi-lt$hh.Count;$xi++){$n.Cells.Item(1,$xi+1)=$hh[$xi]}

# Read prev sheet for matching + row references
$p=$w.Sheets($prev);$pr=$p.UsedRange.Rows.Count;$pRaw=$p.UsedRange.Value2
$exactA=@{};$byB=@{};$byC=@{};$allP=@()
for($pi=2;$pi-le$pr;$pi++){
 $q=[string]$pRaw[$pi,1];if([string]::IsNullOrWhiteSpace($q)){continue}
 $bt=[string]$pRaw[$pi,2];$ct=[string]$pRaw[$pi,3]
 $exactA[$q]=$pi
 if(!$byB.ContainsKey($bt)){$byB[$bt]=@()};$byB[$bt]+=@{R=$pi;C=$ct;Q=$q}
 if(!$byC.ContainsKey($ct)){$byC[$ct]=@()};$byC[$ct]+=@{R=$pi;B=$bt;Q=$q}
 $allP+=@{R=$pi;Q=$q;B=$bt;C=$ct}
}
$msg="prev sheet ${prev}: $($allP.Count) questions, $pr rows";Write-Output $msg

# Match each source row to prev sheet (3-layer)
$eCnt=0;$fCnt=0;$bCnt=0;$cCnt=0;$nCnt=0;$matchRow=@()
foreach($item in $data){
 $q=$item.Q;$b=$item.T;$c=$item.Dt
 if($exactA.ContainsKey($q)){$matchRow+=$exactA[$q];$eCnt++;continue}
 $bestR=$null;$bestS=0
 foreach($r in $allP){$s=FuzzyRatio $q $r.Q;if($s-gt$bestS){$bestS=$s;$bestR=$r.R}}
 if($bestR-and$bestS-ge0.65){$matchRow+=$bestR;$fCnt++;continue}
 if($byB.ContainsKey($b)){
  $cands=$byB[$b];$byC2=$cands|Where-Object{$_.C-eq$c}
  $r=if($byC2.Count-eq1){$byC2[0].R}elseif($cands.Count-eq1){$cands[0].R}else{$null}
  if($r-ne$null){$matchRow+=$r;$bCnt++;continue}
 }
 if($byC.ContainsKey($c)){
  $cands=$byC[$c];$byB2=$cands|Where-Object{$_.B-eq$b}
  $pool=if($byB2.Count-ge1){$byB2}else{$cands}
  $bestR2=$null;$bestS2=0
  foreach($cr in $pool){$s2=FuzzyRatio $q $cr.Q;if($s2-gt$bestS2){$bestS2=$s2;$bestR2=$cr.R}}
  if($bestR2-and$bestS2-ge0.3){$matchRow+=$bestR2;$cCnt++;continue}
 }
 $matchRow+=$null;$nCnt++
}
Write-Output ("match: ${eCnt}exact ${fCnt}fuzzy ${bCnt}B ${cCnt}C ${nCnt}new")

# Pass 1: write A-G values + MNO text
for($di=0;$di-lt$data.Count;$di++){
 $rr=$di+2;$item=$data[$di];$mr=$matchRow[$di]
 $n.Cells.Item($rr,1)=$item.Q;$n.Cells.Item($rr,2)=$item.T;$n.Cells.Item($rr,3)=$item.Dt
 $n.Cells.Item($rr,4)=$item.A;$n.Cells.Item($rr,5)=$item.B;$n.Cells.Item($rr,6)=$item.C;$n.Cells.Item($rr,7)=$item.Gv
  if($mr-ne$null){
   $n.Cells.Item($rr,13)=$pRaw[$mr,13]
   $n.Cells.Item($rr,14)=$pRaw[$mr,14]
   $n.Cells.Item($rr,15)=$pRaw[$mr,15]
  }else{$n.Cells.Item($rr,13)="";$n.Cells.Item($rr,14)="";$n.Cells.Item($rr,15)=""}
}

# Pass 2: write H-L formulas (matching 2505 VLOOKUP pattern) via bulk range
$lastRow=$data.Count+1
$hArr=New-Object object[] $data.Count
$iArr=New-Object object[] $data.Count
$jArr=New-Object object[] $data.Count
$kArr=New-Object object[] $data.Count
$lArr=New-Object object[] $data.Count
for($di=0;$di-lt$data.Count;$di++){
 $rr=$di+2
 $hArr[$di]="=ROUND(E${rr}/D${rr},2)"
 $iArr[$di]="=ROUND(G${rr}/F${rr},2)"
 $jArr[$di]="=F${rr}-IFERROR(VLOOKUP(`$A${rr},'${prev}'!`$A`$1:G${pr},6,FALSE),IFERROR(VLOOKUP(`$C${rr},'${prev}'!`$C`$1:G${pr},4,FALSE),F${rr}-D${rr}))"
 $kArr[$di]="=G${rr}-IFERROR(VLOOKUP(`$A${rr},'${prev}'!`$A`$1:H${pr},7,FALSE),IFERROR(VLOOKUP(`$C${rr},'${prev}'!`$C`$1:H${pr},5,FALSE),G${rr}-E${rr}))"
 $lArr[$di]="=IFERROR(ROUND(K${rr}/J${rr},2),)"
}
$n.Range("H2:H${lastRow}").Formula=$hArr
$n.Range("I2:I${lastRow}").Formula=$iArr
$n.Range("J2:J${lastRow}").Formula=$jArr
$n.Range("K2:K${lastRow}").Formula=$kArr
$n.Range("L2:L${lastRow}").Formula=$lArr

# Sum row
$sumRow=$data.Count+2
$n.Cells.Item($sumRow,10)=[string][char]0x5408+[string][char]0x8BA1
$n.Cells.Item($sumRow,11).Formula="=ROUND(SUM(K2:K$($sumRow-1)),2)"

$ex.ScreenUpdating=$true;$ex.ScreenUpdating=$false
$w.Save()
Write-Output "Save OK"
$w.Close();$ex.Quit()
Write-Output "=== Done ==="
