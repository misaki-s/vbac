Excelマクロ（VBA）をVSCodeで編集したい
https://kanegolabo.com/vba-edit#toc6

Excel2016

VBAのエクスポート(bin -> src)
```
cscript vbac.wsf decombine
```

VBAのインポート(src -> bin)
```
cscript vbac.wsf combine
```

ソースを反映したらExcelマクロを開きデバッグしてみましょう。
Alt+F11キーでExcelのVBAエディタが開くのでソースが反映されていることを確認してからデバッグ作業を行ってください。
デバッグ作業が終わったらExcelは必ず閉じるようにしましょう。
開いたままでvbacによるソース取込は動かない。(意味ないじゃん...)
