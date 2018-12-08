# 生体情報学実験のExcelグラフ生成

## ダウンロード

## 起動

### Windows

`ExcelChartGeneratorForSeitaiJouhogakuZikken.exe` をダブルクリック

### Mac / Linux

ターミナルを開いてプログラムのディレクトリに移動し

`$ mono ExcelChartGeneratorForSeitaiJouhogakuZikken.exe`

## 起動できない時

### Windows

[.NET Framework4.7をダウンロード](https://dotnet.microsoft.com/download/thank-you/net472)してインストール


### Mac / Linux

[mono](https://www.mono-project.com/download/stable/#download-mac)をインストール

## 使いかた

`実験結果csvファイルの親ディレクトリのパスを入力` でcsvファイルが入ったディレクトリをドラックアンドドロップするなどして、指定してください

`[Enter]` を押すと読み込むファイルのリストが表示されるので、正しく読めているか確認してください

`散布図の値の範囲[mm]を設定(推奨4~8)` では生成する散布図のx,y両方の軸の最小値が、{x,y各データの中央値}-{入力した値/2}になり、最大値が{x,y各データの中央値}+{入力した値/2}になります

`散布図の辺の大きさ[px]を設定(推奨400~1000)` では、生成する散布図のx,y両辺の大きさを指定します。両辺を一つの値で指定するので生成される散布図は正方形になります

## 生成されるもの

指定したディレクトリの中に`Generated`という名前のあたらしいディレクトリが生成されて、結果が保存されます

- 散布図のグラフ
- 平均と分散の計算結果(グラフなし)
- フーリエ変換の計算結果(グラフなし)

## ライセンス

ソフトのライセンスはMITライセンス(改変、二次配付の自由、ユーザーに生じた損害からの免責)です。同意の上ご利用ください
