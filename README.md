# DataEntryTool
Excelに入力したデータをデータベースに投入するツールです。

### 機能
* データ投入用のテーブルシートの作成
* データベースへのデータ登録、更新、削除

### 必須条件
データベースにアクセスするため、対応したODBCドライバをインストールしてください。
* [Oracle](http://www.oracle.com/technetwork/jp/topics/utilsoft-100274-ja.html)
* [PostgreSQL](http://www.postgresql.org/ftp/odbc/versions/msi/)

### 使い方
データ登録、更新、削除、DML出力には、テーブルシートの作成が必要です。  
はじめにデータ投入用のテーブルシートを作成してください。

#### 1. テーブルシート作成  
* データベース設定の入力  
接続先のデータベース設定を入力します。  
※PostgreSQLの場合はポートとデータベース名も入力すること。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897672/5fd39556-2dd0-11e6-8275-5ae985dad876.png" width="50%">

* テーブル一覧の入力  
テーブルシートを作成したいテーブルを入力します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897679/627a2630-2dd0-11e6-9915-0162773cdf2e.png" width="50%">

* テーブルシートの作成  
「テーブルシート作成」ボタンを押下してテーブルシートを作成します。
<img src="https://cloud.githubusercontent.com/assets/14181039/15897682/63806fb2-2dd0-11e6-89f3-3023f331ec22.png" width="50%">  
＜作成結果＞  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897685/658d7ce6-2dd0-11e6-9a7e-755418a2fe86.png" width="50%">

#### 2. データ登録  
データベースにデータを登録します。  
 * データ投入対象の設定  
データを登録するテーブルの「データ投入対象」列に空文字以外の値を設定します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897692/6db54d5e-2dd0-11e6-8de6-128d26bd89c2.png" width="50%">

* 登録データの入力  
テーブルシートに登録データを入力します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897694/6f6786da-2dd0-11e6-904a-97b704b7b521.png" width="50%">

* データ登録  
「データ登録」ボタンを押下してデータを登録します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897698/73940d82-2dd0-11e6-99b1-e74348ef7133.png" width="50%">  
＜登録結果＞  
<img src="https://cloud.githubusercontent.com/assets/14181039/15897699/755b727c-2dd0-11e6-9308-4f23e091d131.png" width="50%">

* データ更新、削除  
データ登録と同様に、テーブルシートに入力された内容でデータを更新、削除します。  
※条件には主キーが指定されます。

#### 3. DML出力  
データ投入時のクエリをDMLとして出力します。

* 出力先の入力  
DMLを出力するディレクトリを設定します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15900189/873d462e-2dd9-11e6-897c-2ae4d7a1cc9b.png" width="50%">

* データ投入対象の設定  
データを登録するテーブルの「データ投入対象」列に空文字以外の値を設定します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15900190/8763300a-2dd9-11e6-83d8-19562487f8fe.png" width="50%">

* 登録データの入力  
テーブルシートに登録データを入力します。  
<img src="https://cloud.githubusercontent.com/assets/14181039/15900193/877767aa-2dd9-11e6-9195-b3c21910eeb0.png" width="50%">  

* 登録DML出力  
「登録DML出力」ボタンを押下して登録DMLを出力します。
<img src="https://cloud.githubusercontent.com/assets/14181039/15900192/877740cc-2dd9-11e6-8087-f867c8a13453.png" width="50%">  
＜出力結果＞  
<img src="https://cloud.githubusercontent.com/assets/14181039/15900191/87763e0c-2dd9-11e6-8e2b-7c44914d8349.png" width="50%">

* 更新、削除DML出力  
登録DML出力と同様に、更新、削除クエリをDMLとして出力します。  

### データ投入設定
データ投入時、設定値に応じてデータを加工して投入します。

#### 1. 改行コード  
投入するカラムのデータ型が文字列の場合、入力値の改行コードを変換して投入します。

|改行コード|入力値|結果|
|:---------|:-----|:---|
|LF|サン<br>プル１|`'サン' || CHR(10) || 'プル１'`|
|CRLF|サン<br>プル２|`'サン' || CHR(13) || CHR(10) || 'プル２'`|

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

### 出力設定
DML出力時、設定値に応じて出力します。

#### 1. 出力先  
DML出力時に出力するディレクトリを設定します。  
未設定の場合、出力時に表示されるダイアログで指定します。

### ライセンス

* [The MIT License (MIT)](LICENSE)

### 使用モジュール

以下のモジュールを使用して開発を行っています。

#### [Ariawase](https://github.com/vbaidiot/Ariawase)

> The MIT License (MIT)
>
> Copyright (c) 2011-2015 igeta

* **用途 :** インポート/エクスポート処理
* **ライセンス :** The MIT License (MIT)
* **ライセンス全文 :** [licenses/Ariawase.txt](licenses/Ariawase.txt)
