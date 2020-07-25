$rawtext = Get-Content ".\under_10000yen_PSUs.csv"

$rawtext[0] = $rawtext[0] | ForEach-Object {
    $_ -replace "status", "※" `
       -replace "form factor", "規格" `
       -replace "depth", "奥行" `
       -replace "modular(?!\s)", "ｹｰﾌﾞﾙ" `
       -replace "EPS8pin", "EPS 8pin" `
       -replace "platform", "ベースモデル" `
       -replace "pcb", "PCB" `
       -replace "primary topology", "1次側方式" `
       -replace "primary cap\(s\) brand", "1次側Cap(s)" `
       -replace "secondary side topology", "2次側方式" `
       -replace "secondary electrolytic caps brand", "2次側液体Caps" `
       -replace "secondary solid caps brand", "2次側固体Caps" `
       -replace "modular board caps", "プラグイン基板Caps" `
       -replace "fan(?!less)", "ファン" `
       -replace "fanless mode", "ﾌｧﾝﾚｽ" `
       -replace "note", "備考" `
       -replace "review/image_", "画像等"
}

$content = $rawtext | ForEach-Object {
    $_ -replace "discontinued", "終売" `
       -replace "Not Certified", "未取得" `
       -replace "single sided", "片面"`
       -replace "double sided", "両面"`
       -replace "double forward", "ﾀﾞﾌﾞﾙﾌｫﾜｰﾄﾞ" `
       -replace "Active-Clamp", "ｱｸﾃｨﾌﾞｸﾗﾝﾌﾟ" `
       -replace "half-bridge", "ﾊｰﾌﾌﾞﾘｯｼﾞ" `
       -replace "full-bridge", "ﾌﾙﾌﾞﾘｯｼﾞ" `
       -replace "Synchronous Rectification", "同期整流" `
       -replace "Passive Rectification", "ﾀﾞｲｵｰﾄﾞ整流" `
       -replace "Group Regulation", "12V/5V一括制御" `
       -replace "Group\(5V/3\.3V\) Regulation", "5V/3.3V一括制御" `
       -replace "鑫", "(金金金)" `
       -replace "九州阳光电源", "九州陽光電源" `
       -replace "japanese", "日本メーカー" `
       -replace "unknown", "不明" `
       -replace "JunFu", "&#x4a;unFu" `
       -replace "HongHua", "&#x48;ongHua" `
       -replace "MasterWatt", "&#x4d;asterWatt" `
       -replace "CapXon", "&#x43;apXon" `
       -replace "XinHuiYuan", "&#x58;inHuiYuan" `
       -replace "KuangJin", "&#x4b;uangJin" `
} | ConvertFrom-Csv

$content | ForEach-Object {
    #製品リンク生成
    $_.name = "[[$($_.name):$($_.link)]]"

    #値の置き換えと整形
    $_."2次側液体Caps" = $_."2次側液体Caps" -creplace "/(?!a)", " / "
    $_."2次側固体Caps" = $_."2次側固体Caps" -creplace "/(?!a)", " / "
    $_."プラグイン基板Caps" = $_."プラグイン基板Caps" -replace "/(?!a)", " / "

    #画像リンク生成 ※URLがある場合リンク化し、「-」の場合空欄にする
    $_."画像等1" = if($_."画像等1" -ne "-"){"[[link:$($_."画像等1")]]"}else{""}
    $_."画像等2" = if($_."画像等2" -ne "-"){"[[link:$($_."画像等2")]]"}else{""}
    $_."画像等3" = if($_."画像等3" -ne "-"){"[[link:$($_."画像等3")]]"}else{""}
    $_."画像等4" = if($_."画像等4" -ne "-"){"[[link:$($_."画像等4")]]"}else{""}
}

$content = $content | Select-Object * -ExcludeProperty "price", "link"

#pukiwiki記法テーブルへ変換
$csv = $content | 
    ConvertTo-Csv -Delimiter "|" -NoTypeInformation |
    ForEach-Object {$_ -replace '"', ""}

$csv = $csv -replace "^(.+)$", '|$1|'
$csv[0] = $csv[0] + "h"

Set-Content -Path ".\under_10000yen_PSUs_pukiwiki.txt" -Value $csv -Encoding "Default"