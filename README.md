# DataEntryTool
Excelに入力したデータをデータベースに投入するツールです。

### 機能
* データ投入用のテーブルシートの作成
* データベースへのデータ登録、更新、削除

### 準備
使用するデータベースに対応したODBCドライバをインストールしてください。
* [Oracle](http://www.oracle.com/technetwork/jp/topics/utilsoft-100274-ja.html)
* [PostgreSQL](http://www.postgresql.org/ftp/odbc/versions/msi/)

### 使い方
 * テーブルシートの作成
 * データ投入  
データを投入するためにはまず、データ投入用のテーブルシートを作成する必要があります。

1. テーブルシート作成  
はじめにデータ投入用のテーブルシートを作成してください。
 * データベース設定の入力  
接続先のデータベース設定を入力してください。  
※PostgreSQLの場合はポートとデータベース名も入力すること。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14609352/6569f0fc-05c4-11e6-9872-66c027423006.png" width="50%">

 * テーブル一覧の入力  
テーブルシートを作成したいテーブルを入力してください。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14609355/6791cc7e-05c4-11e6-8960-11e65a092a22.png" width="50%">

 * テーブルシートの作成  
テーブルシート作成ボタンを押下してテーブルシートを作成します。
<img src="https://cloud.githubusercontent.com/assets/14181039/14609591/74abceb8-05c5-11e6-8edd-996ec2f3d363.png" width="50%">  
↓  
<img src="https://cloud.githubusercontent.com/assets/14181039/14609560/4a328e24-05c5-11e6-9b33-a62cfeb7ee15.png" width="50%">

### ライセンス

* [The MIT License (MIT)](LICENSE)

### 使用ライブラリ

以下のモジュールを使用して開発を行っています。

#### [Ariawase](https://github.com/vbaidiot/Ariawase)

> The MIT License (MIT)
>
> Copyright (c) 2011-2015 igeta

* **用途 :** インポート/エクスポート処理
* **ライセンス :** The MIT License (MIT)
* **ライセンス全文 :** [licenses/Ariawase.txt](licenses/Ariawase.txt)
