# ライセンス

このプロジェクトは、[XlsWxg MITライセンス](https://github.com/lqwangxg/XlsWxg)のもとで提供されています。

MITライセンスの全文は以下の通りです。

[ライセンスの全文を表示](https://github.com/lqwangxg/XlsWxg/blob/master/LICENSE)  
<br>

# VSCode XVBA で改造する際の注意点
- 対象のExcelを開いた状態でExportを実行すること（Excelを閉じたままExportするとExcelが壊れて開けなくなる可能性有り）  
<br>

# 概要
XlsWxgから提供されているOracleのデータ取得Excelを改造しました。  

改造内容は下記です。
- 改造前：BatchQueryのSQL実行結果が1つのシートへ格納
- 改造後：BatchQueryのSQL実行結果をSQL毎に各シートへ格納

SQL実行結果をSQL毎のシートに分けたのは、取得したデータをPowerQueryで扱い易かったためです。  
<br>

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
<br>

### 4. BatchQueryシートにSQLを記載
![img](/img/5.png)  
<br>

### 5. SQL実行
![img](/img/6.png)  
<br>

### 6. SQL結果を確認  
BatchQueryソートのA列に記載した文字列（No./ExecFLag）のシートを生成し、そのシートにデータを格納します。  
![img](/img/7.png)  
