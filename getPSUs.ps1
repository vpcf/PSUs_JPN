Set-StrictMode -Version Latest

function generate_kakaku_link($num){
    "https://kakaku.com/pc/power-supply/itemlist.aspx?pdf_so=p1&pdf_pg=" + $num + "&pdf_pr=-100000"
}

function parse_80plus_cert($ckitemSpecInnr){
    switch($ckitemSpecInnr.length){
        {$_ -gt 0} {
            if($ckitemSpecInnr | ForEach-Object{$_.textContent -match "80PLUS認証：(\w+)"}){
                $Matches[1]
            }else{
                "Not Certified"
            }
        }
        default {"nodata"}
    }
}

function parse_depth($td_size){
    $captures = [regex]::Match($td_size.outerText, "(?:(\d+\.?\d*)x*){3}").groups[1].captures
    if($captures.count){
        $dimension = [double[]]$captures.Value
        if(86 -in $dimension){
            if($dimension -ne 150 -ne 86){$dimension -ne 150 -ne 86}
            else{150} #ATX of 150mm depth
        }
        elseif(65 -in $dimension){$dimension -ne 85 -ne 65}
        elseif(125 -in $dimension){
            if($dimension -ne 125 -gt 65){$dimension -ne 125 -gt 65}
            else{125} #SFX-L of 125mm depth
        }
        else{
            $dimension | Measure-Object -Maximum | ForEach-Object Maximum #Other Platform
        }
    }else{
        0
    }
}

function parse_numof_eps($ckitemSpecInnr){
    switch($ckitemSpecInnr.length){
        {$_ -gt 0} {
            if($ckitemSpecInnr | ForEach-Object{$_.textContent -match "(?=CPU用コネクタ：|EPS).+?(x\d)"}){
                $Matches[1]
            }else{
                "Unknown"
            }
        }
        default {"nodata"}
    } 
}
$page_count = 0

$PSUs = while(++$page_count){
    generate_kakaku_link $page_count | 
    ForEach-Object {
        try{
            $request = Invoke-WebRequest $_
            Write-Host "Page" $page_count
        }catch{
            break
        }

        $table = $request.ParsedHtml.getElementById("compTblList")

        #リスト1製品
        $1stRow = $table.getElementsByClassName("sel alignC ckbtn") | ForEach-Object {$_.parentNode}
        $2ndRow = $table.getElementsByClassName("td-price") | ForEach-Object {$_.parentNode}
        $3rdRow = $2ndRow | ForEach-Object {$_.nextSibling}

        #メーカー名・製品名
        $col0_1 = $1stRow.getElementsByClassName("ckitemLink") | ForEach-Object {$_.outerText}

        #比較ページリンク
        $href = $1stRow.getElementsByClassName("ckitanker") | ForEach-Object {$_.href}

        #最安価格
        $price = $2ndRow.getElementsByClassName("pryen") | ForEach-Object {$_.outerText}

        #80PLUS
        $80plus = $3rdRow | ForEach-Object {
            parse_80plus_cert $_.getElementsByClassName("ckitemSpecInnr")
        }

        #奥行(mm)
        $depth = $2ndRow.getElementsByClassName("end").previousSibling | ForEach-Object {
            parse_depth $_
        }

        #CPU8ピン数
        $8pin = $3rdRow | ForEach-Object {
            parse_numof_eps $_.getElementsByClassName("ckitemSpecInnr")
        }

        #製品ごとにPSCustomObject化
        0..($col0_1.Count - 1) | ForEach-Object {
            #末尾の[カラー]を除去しメーカー名と製品名に分割
            $brand, $name = $col0_1[$_] -replace "\s\[.*\]", "" -split "　"

            #メーカー名表記置き換え
            $brand = $brand -replace "クーラーマスター", "Cooler Master"
            $brand = $brand -replace "サイズ", "Scythe"

            [PSCustomObject]@{
                "brand"=$brand;
                "name"=$name;
                "link"=$href[$_];
                "price"=$price[$_];
                "80PLUS"=$80plus[$_];
                "depth"=$depth[$_];
                "8pin"=$8pin[$_];
            }
        }
    }
}

$outname = Get-Date -Format "'.\\PSUs_'yyyyMMdd_HHmm'.csv'"
$PSUs | Export-Csv -Path $outname -NoTypeInformation -Append -Encoding "UTF8"