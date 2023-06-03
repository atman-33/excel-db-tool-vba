# ライセンス

このプロジェクトは、[XlsWxg MITライセンス](https://github.com/lqwangxg/XlsWxg)のもとで提供されています。

MITライセンスの全文は以下の通りです。

[ライセンスの全文を表示](https://github.com/lqwangxg/XlsWxg/blob/master/LICENSE)


# 概要
XlsWxgから提供されているOracleのデータ取得Excelを改造しました。  

改造内容は下記です。
- 改造前：BatchQueryのSQL実行結果が1つのシートへ格納
- 改造後：BatchQueryのSQL実行結果をSQL毎に各シートへ格納

SQL実行結果をSQL毎のシートに分けたのは、取得したデータをPowerQueryで扱い易かったためです。

# 利用方法

対象:Windows 64bit

### 1. excel-db-tool.xlsm を開く
<br>

### 2. コンテンツを有効化する
![img](/img/1.png)  
<br>

### 3. Excelアドインを追加
![img](/img/2.png)  

XlsWxg-AddIn64.xll を選択
![img](/img/3.png)  

![img](/img/4.png)  

### 4. 