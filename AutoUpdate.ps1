$ErrorActionPreference="Stop"
$c1=[char]0x81F4+[char]0x77E5+[char]0x8BA1+[char]0x5212+[char]0x5185+[char]0x5BB9+[char]0x6536+[char]0x76CA
$c2=[char]0x81F4+[char]0x77E5+[char]0x6536+[char]0x76CA
$srcPath="G:\me\"+$c2+"\2025\"+$c1+".xls"
$dstPath="G:\me\"+$c2+"\2025\2607.xlsx"

$ex=New-Object -ComObject Excel.Application
$ex.Visible=$false;$ex.DisplayAlerts=$false;$ex.ScreenUpdating=$false;$ex.EnableEvents=$false

$wb=$ex.Workbooks.Open($srcPath)
$s=$wb.Sheets(1);$sr=$s.UsedRange.Rows.Count
$raw=$s.UsedRange.Value2
$h1=$raw[1,1];$h2=$raw[1,2];$h3=$raw[1,3];$h4=$raw[1,4];$h5=$raw[1,5];$h6=$raw[1,6];$h7=$raw[1,7]
$data=@()
for($ri=2;$ri-le$sr;$ri++){
 $q=[string]$raw[$ri,1]
 if([string]::IsNullOrWhiteSpace($q)){continue}
 $tp=[string]$raw[$ri,2]
 $dv=$raw[$ri,3]
 if($dv-ne$null){
  if($dv.GetType().Name -eq "Double"){$ds=[datetime]::FromOADate([double]$dv).ToString("yyyy/M/d")}
  else{$ds=$dv.ToString()}
 }else{$ds=""}
 $va1=[double]($raw[$ri,4]);$va2=[double]($raw[$ri,5])
 $va3=0;try{$va3=[double]($raw[$ri,6])}catch{}
 $va4=0;try{$va4=[double]($raw[$ri,7])}catch{}
 $data+=@{Q=$q;T=$tp;Dt=$ds;A=$va1;B=$va2;C=$va3;Gv=$va4}
}
Write-Output ($data.Count.ToString()+" source rows")
$wb.Saved=$true;$wb.Close()

$w=$ex.Workbooks.Open($dstPath)

$maxS=0;$prev=""
for($i=1;$i-le$w.Sheets.Count;$i++){
 $sn=$w.Sheets($i).Name
 if($sn -match "^\d{4}$" -and [int]$sn -gt $maxS){$maxS=[int]$sn;$prev=$sn}
}
$today="{0:D4}"-f($maxS+1)
Write-Output ($prev+" -> "+$today)

try{$d=$w.Sheets($today);$d.Delete();[Runtime.InteropServices.Marshal]::ReleaseComObject($d)|Out-Null}catch{}

$n=$w.Worksheets.Add()
$n.Name=$today
$n.Move([Type]::Missing,$w.Sheets($w.Sheets.Count))

$aa=[string][char]0x5DEE;$bb=[string][char]0x503C;$cz="$aa$bb"
$cc=[string][char]0x524D;$dd=[string][char]0x65E5;$qd="$cc$dd"
$hD="D"+$cz;$hE="E"+$cz;$hL="L"+$today;$hqD=$qd+"D";$hqE=$qd+"E";$hqF=$qd+"F"
$hh=@($h1,$h2,$h3,$h4,$h5,$h6,$h7,"E/D","G/F",$hD,$hE,$hL,$hqD,$hqE,$hqF)
for($xi=0;$xi-lt$hh.Count;$xi++){$n.Cells.Item(1,$xi+1)=$hh[$xi]}

$p=$w.Sheets($prev);$pr=$p.UsedRange.Rows.Count
$pData=$p.UsedRange.Value2
$pl=@{}
for($pi=2;$pi-le$pr;$pi++){
 $q=[string]$pData[$pi,1]
 if([string]::IsNullOrWhiteSpace($q)){continue}
 $pm=[double]($pData[$pi,13]);$pn=[double]($pData[$pi,14]);$po=[double]($pData[$pi,15])
 $pl[$q]=@{M=$pm;N=$pn;O=$po}
}
[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null

$mc=0;$nc=0
for($di=0;$di-lt$data.Count;$di++){
 $rr=$di+2;$item=$data[$di]
 $n.Cells.Item($rr,1)=$item.Q;$n.Cells.Item($rr,2)=$item.T;$n.Cells.Item($rr,3)=$item.Dt
 $n.Cells.Item($rr,4)=$item.A;$n.Cells.Item($rr,5)=$item.B;$n.Cells.Item($rr,6)=$item.C;$n.Cells.Item($rr,7)=$item.Gv
 $n.Cells.Item($rr,8).Formula="=ROUND(E$rr/D$rr,2)"
 $n.Cells.Item($rr,9).Formula="=ROUND(G$rr/F$rr,2)"
 if($pl.ContainsKey($item.Q)){$v=$pl[$item.Q];$aa=$v.M;$bb=$v.N;$xc=$v.O;$mc++}else{$aa=0;$bb=0;$xc=0;$nc++}
 $n.Cells.Item($rr,13)=$aa;$n.Cells.Item($rr,14)=$bb;$n.Cells.Item($rr,15)=$xc
 $n.Cells.Item($rr,10).Formula="=IF(M$rr=0,D$rr,D$rr-M$rr)"
 $n.Cells.Item($rr,11).Formula="=IF(N$rr=0,E$rr,E$rr-N$rr)"
 $n.Cells.Item($rr,12).Formula="=IF(J$rr=0,0,ROUND(K$rr/J$rr,2))"
}
Write-Output ($mc.ToString()+" matched, "+$nc.ToString()+" new")

$sumRow=$data.Count+2
$n.Cells.Item($sumRow,10)=[string][char]0x5408+[string][char]0x8BA1
$n.Cells.Item($sumRow,11).Formula="=SUM(K2:K$($sumRow-1))"

[Runtime.InteropServices.Marshal]::ReleaseComObject($n)|Out-Null
$w.Save()
$w.Close()
$ex.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($w)|Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($ex)|Out-Null
Write-Output "=== Done ==="
