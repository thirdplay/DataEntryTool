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

#### 1. テーブルシート作成  
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

#### 2. データ登録  
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

#### 3. データ更新、削除  
データ登録と同様に、テーブルシートに入力された内容でデータを更新、削除します。  
※条件には主キーが指定されます。

### データ投入設定
データ投入時、設定値に応じてデータを加工して投入します。

#### 1. 改行コード  
投入するカラムのデータ型が文字列の場合、入力値の改行コードを変換して投入します。

|改行コード|入力値|結果|
|:---------|:-----|:---|
|LF|サン<br>プル１|`'サンプル' || CHR(10) || '１'`|
|CRLF|サン<br>プル２|`'サンプル' || CHR(13) || CHR(10) || '２'`|

#### 2. 日付書式  
投入するカラムのデータ型が日付の場合、TO_DATE関数で文字列から日付に変換して投入します。  
日付書式はTO_DATE関数に指定するフォーマット文字列になるため  
テーブルシートに入力する日付は日付書式に従った形式で入力する必要があります。  

|日付書式|入力値|結果|
|:---------|:-----|:---|
|YYYYMMDD|20160101|`TO_DATE('20160101','YYYYMMDD')`|
|YYYY/MM/DD|2016/02/02|`TO_DATE('2016/02/02','YYYY/MM/DD')`|
|YYYY-MM-DD|2016-03-03|`TO_DATE('2016-03-03','YYYY-MM-DD')`|

#### 3. タイムスタンプ書式  
投入するカラムのデータ型がタイムスタンプの場合、TO_TIMESTAMP関数で文字列からタイムスタンプに変換して投入します。  
タイムスタンプ書式はTO_TIMESTAMP関数に指定するフォーマット文字列になるため  
テーブルシートに入力するタイムスタンプはタイムスタンプ書式に従った形式で入力する必要があります。  

|タイムスタンプ書式|入力値|結果|
|:---------|:-----|:---|
|YYYYMMDDHH24MISSFF|20160101104019123456789|`TO_TIMESTAMP('20160101104019123456789',`<br>`'YYYYMMDDHH24MISSFF')`|
|YYYY-MM-DD HH24:MI:SS.FF|2016-02-02 10:41:20.123456789|`TO_TIMESTAMP('2016-02-02 10:41:20.123456789',`<br>`'YYYY-MM-DD HH24:MI:SS.FF')`|

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
