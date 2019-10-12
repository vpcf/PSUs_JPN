$rawtext = Get-Content ".\under_10000yen_PSUs.csv"

$rawtext[0] = $rawtext[0] | ForEach-Object {
    $_ -replace "form factor", "規格" `
       -replace "depth", "奥行" `
       -replace "modular(?!\s)", "ｹｰﾌﾞﾙ" `
       -replace "EPS8pin", "EPS 8pin" `
       -replace "platform", "ベースモデル" `
       -replace "pcb", "PCB" `
       -replace "primary topology", "1次側方式" `
       -replace "primary cap(s) brand", "1次側Cap(s)" `
       -replace "secondary side topology", "2次側方式" `
       -replace "secondary electrolytic caps brand", "2次側液体Caps" `
       -replace "secondary solid caps brand", "2次側固体Caps" `
       -replace "modular board caps", "プラグイン基板Caps" `
       -replace "fan", "ファン" `
       -replace "fanless mode", "ファンレス" `
       -replace "note", "備考" `
       -replace "review/image_", "画像等"
}

$content = $rawtext | ForEach-Object {
    $_ -replace "Not Certified", "未取得" `
       -replace "single sided", "片面"`
       -replace "double sided", "両面"`
       -replace "double forward", "ﾀﾞﾌﾞﾙﾌｫﾜｰﾄﾞ" `
       -replace "Active-Clamp", "ｱｸﾃｨﾌﾞｸﾗﾝﾌﾟ" `
       -replace "half-bridge", "ﾊｰﾌﾌﾞﾘｯｼﾞ" `
       -replace "full-bridge", "ﾌﾙﾌﾞﾘｯｼﾞ" `
       -replace "Synchronous Rectification", "同期整流" `
       -replace "Passive Rectification", "ﾀﾞｲｵｰﾄﾞ整流" `
       -replace "Group Regulation", "12V/5V一括制御" `
       -replace "Group(5V/3.3V) Regulation", "5V/3.3V一括制御" `
       -replace "鑫", "(金金金)" `
       -replace "九州阳光电源", "九州陽光電源" `
       -replace "japanese", "日本メーカー" `
       -replace "unknown", "不明" `
} | ConvertFrom-Csv

#リンクテキスト
$content | ForEach-Object {
    $_.name = "[[$($_.name):$($_.link)]]"
    $_."2次側液体Caps" = $_."2次側液体Caps" -replace "/(?!a)", " / "
    $_."2次側固体Caps" = $_."2次側固体Caps" -replace "/(?!a)", " / "
    $_."プラグイン基板Caps" = $_."プラグイン基板Caps" -replace "/(?!a)", " / "
    $_."画像等1" = if($_."画像等1" -ne "-"){"[[link:$($_."画像等1")]]"}
    $_."画像等2" = if($_."画像等2" -ne "-"){"[[link:$($_."画像等2")]]"}
    $_."画像等3" = if($_."画像等3" -ne "-"){"[[link:$($_."画像等3")]]"}
    $_."画像等4" = if($_."画像等4" -ne "-"){"[[link:$($_."画像等4")]]"}
}

$content = $content | Select-Object * -ExcludeProperty "price", "link"

$csv = $content | 
    ConvertTo-Csv -Delimiter "|" -NoTypeInformation |
    ForEach-Object {$_ -replace '"', ""}

$csv = $csv -replace "^(.+)$", '|$1|'
$csv[0] = $csv[0] + "h"

Set-Content -Path ".\under_10000yen_PSUs_pukiwiki.txt" -Value $csv -Encoding "Default"