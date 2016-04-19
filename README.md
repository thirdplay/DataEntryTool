# DataEntryTool
Excelに入力したデータをデータベースに投入するツールです。

### 機能
* データ投入用のテーブルシートの作成
* データベースへのデータ登録、更新、削除

### 準備
データベースにアクセスするため、対応したODBCドライバをインストールしてください。
* [Oracle](http://www.oracle.com/technetwork/jp/topics/utilsoft-100274-ja.html)
* [PostgreSQL](http://www.postgresql.org/ftp/odbc/versions/msi/)

### 使い方
データ登録、更新、削除には、テーブルシートの作成が必要です。  
はじめにデータ投入用のテーブルシートを作成してください。

1. テーブルシート作成  
 * データベース設定の入力  
接続先のデータベース設定を入力します。  
※PostgreSQLの場合はポートとデータベース名も入力すること。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14646971/96c124f4-0696-11e6-933d-8e0ea053a9fe.png" width="50%">

 * テーブル一覧の入力  
テーブルシートを作成したいテーブルを入力します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14646972/96c16b80-0696-11e6-93b0-3ee9ca9ae576.png" width="50%">

 * テーブルシートの作成  
「テーブルシート作成」ボタンを押下してテーブルシートを作成します。
<img src="https://cloud.githubusercontent.com/assets/14181039/14646973/96e03f74-0696-11e6-8ec0-13e22470f6b1.png" width="50%">  
＜作成結果＞  
<img src="https://cloud.githubusercontent.com/assets/14181039/14648490/8742a19a-069d-11e6-9125-606f3ce75b36.png" width="50%">
2. データ登録  
データベースにデータを登録します。  
 * データ投入対象の設定  
データを登録するテーブルの「データ投入対象」列に空文字以外の値を設定します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14649025/29cf4a9c-06a0-11e6-9d9b-3f85683ec996.png" width="50%">

 * 登録データの入力  
テーブルシートに投入データを入力します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14649024/29cdc28a-06a0-11e6-8887-a98e5404f2dd.png" width="50%">

 * データ登録  
「データ登録」ボタンを押下してデータを登録します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/14649023/29cc2dbc-06a0-11e6-97c5-c81d5f0c2e99.png" width="50%">  
＜登録結果＞  
<img src="https://cloud.githubusercontent.com/assets/14181039/14649026/29cf783c-06a0-11e6-9c87-2d0e9c9f6801.png" width="50%">

3. データ更新  
データ登録と同様に、テーブルシートに入力された内容でデータを更新します。  
※条件には主キーが指定されます。

4. データ削除  
データ登録と同様に、テーブルシートに入力された内容でデータを削除します。  
※条件には主キーが指定されます。

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
