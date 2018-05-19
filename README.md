# xcut

## 実行例
引数なしで実行するとシート名のリストを出力します。  
```
$ ./xcut hoge.xlsx
Sheet1
Sheet2
```

-s でシート名を指定すると、当該シートの内容をタブ区切りで出力します。
```
$ ./xcut -s Sheet2 hoge.xlsx
A1	B1	C1
～省略～
```

-c で範囲を指定すると、矩形切り出しになります。コロンの前後は省略可能です。( 例. B1: )
```
$ ./xcut -s Sheet2 -c A1:C1 hoge.xlsx
A1	B1	C1
```

-k でキーワード検索できます。-a が指定されていない場合は、最初の1件のみ出力します。
```
$ ./xcut -k 攻撃 -a hoge.xlsx
Sheet1!D1       Text=[武器攻撃力]
Sheet1!AA1      Text=[①攻撃力]
Sheet1!I2       Text=[攻撃力]
～省略～
```

-k では正規表現(regexpパッケージ)を使用可能です。
```
$ ./xcut -k ^攻撃 -a hoge.xlsx
Sheet1!I2       Text=[攻撃力]
Sheet1!S2       Text=[攻撃速度]
～省略～
```

```
$ ./xcut --help
Usage of ./xcut:
  -a    searching all
  -c string
        cut off data
  -f string
        field separator (default "\t")
  -k string
        keyword for search
  -s string
        specified sheet name
```

