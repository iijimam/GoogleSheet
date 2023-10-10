# GoogleAPIを利用してGoogleSheetを操作する例

以下の記事👇を参考に、IRISのEmbedded Pythonを利用してGoogleSheetを操作した例をご紹介します

参照記事：[PythonとSheets API v4でGoogleスプレッドシートを読み書きする](https://www.kumilog.net/entry/2018/03/22/090000)

- Google Sheet APIを利用するための準備については上記記事をご参照ください。

- OAuth2.0クライアントIDが作成できたら、client_secret.jsonの名称でREADMEがあるディレクトリに配置してください。

## 変更した箇所
- OAuth クライアントID作成時、記事では「アプリケーションの種類」に「その他」を選択していますが「ウェブアプリケーション」を選択して作成しました。

- OAuth同意画面でテストユーザを追加しました。
![](/assets/testuser.png)

## Pythonパッケージのインストール

仮想環境を作成し、インストールしています。仮想環境用ディレクトリ（dev01）を作り作成したディレクリに対して以下実行。
```
python -m venv dev01 
```

アクティベートするために以下実行
```
.\dev01\Scripts\activate
```

仮想環境に、[requirements.txt](/requirements.txt) に記載されたパッケージをインポート
```
pip install -r .\requirements.txt
```

## GoogleSheet側の事前準備

シートIDを指定してアクセスするため、GoogleDriveにGoogleSheetを用意します。

複数シートに一括でデータアップロードもできるので、複数のシートを事前に用意しておきます。

`https://docs.google.com/spreadsheets/d/～～/edit#gid=0` のようなURLが作成されます。`～～/d/`以降にある文字列がシートIDとしてアクセスする際必要になるので、メモしておきます。

![](/assets/GoogleSheet.png)

## IRIS側の準備

IRISに登録のある情報を利用してGoogleSheetを作成する場合のIRISへの接続方法は以下3種類あります。

- PyODBC経由：ご参考[「【はじめての InterSystems IRIS】セルフラーニングビデオ：アクセス編：Python から PyODBC を使って IRIS に接続してみよう」](https://jp.community.intersystems.com/node/478616)
- NativeAPI（IRISのPython用SDK）：ご参考[「【はじめての InterSystems IRIS】セルフラーニングビデオ：アクセス編：Python の NativeAPI に挑戦」](https://jp.community.intersystems.com/node/478611)
- Embedded Python：ご参考[「【はじめてのInterSystems IRIS】Embedded Python セルフラーニングビデオシリーズ公開！」](https://jp.community.intersystems.com/node/520751)

**★サンプルでは、「Embedded Python」を利用しています。**

データ入手後、GoogleSheetの操作に必要なJSONを生成する方法としては、
- Pythonで作成
- IRISで作成
のどちらかの方法を選択できます。

**★サンプルでは、IRIS側で作成する方法として、JSONテンプレートエンジンを利用する例をご紹介します。**

### JSONテンプレートエンジンを利用するための準備

必要なクラス定義をダウンロード＆インポートします。
https://github.com/Intersystems-jp/JSONTemplate/tree/main/src/JSONTemplate 以下にある[JSONTemplate.Baseクラス](https://github.com/Intersystems-jp/JSONTemplate/blob/main/src/JSONTemplate/Base.CLS) と [JSONTemplate.Generatorクラス](https://github.com/Intersystems-jp/JSONTemplate/blob/main/src/JSONTemplate/Generator.CLS)

> 利用方法詳細については[「複雑なJSONの生成に便利な「JSONテンプレートエンジン」の使い方ご紹介」](https://jp.community.intersystems.com/node/551396)をご参照ください。

### JSONの用意
今回用意するJSONは以下のような形式のJSONです。

ご参考：https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values

- 指定シートの情報のみを更新する場合のJSON
```
{
    "range": "sales!A1:C3",
    "majorDimension": "ROWS",
    "values": [
        [
            "年月日",
            "ぶどう",
            "りんご"
        ],
        [
            "2023年10月1日",
            "199",
            "29"
        ],
        [
            "2023年10月2日",
            "2",
            "399"
        ]
    ]
}
```

- 複数シートを一括で更新する場合のJSON例

```
{
  "valueInputOption": "RAW",
  "data": [
      {
        "range": "sales!A1:C3",
        "majorDimension": "ROWS",
        "values": [
            [
                "年月日",
                "ぶどう",
                "りんご"
            ],
            [
                "2023年10月1日",
                "199",
                "29"
            ],
            [
                "2023年10月2日",
                "2",
                "399"
            ]
        ]
    },
      {
        "range": "sales2!A1:C3",
        "majorDimension": "ROWS",
        "values": [
            [
                "年月日",
                "ぶどう",
                "りんご"
            ],
            [
                "2023年10月1日",
                "199",
                "29"
            ],
            [
                "2023年10月2日",
                "2",
                "399"
            ]
        ]
    }
  ],
  "includeValuesInResponse": true,
  "responseValueRenderOption": "FORMATTED_VALUE"
}
```

### 用意したテンプレートクラス

- [各シートに登録する項目：GoogleSheet.SheetTesmplate](/GoogleSheet/SheetTemplate.cls)

- [POST要求時のBODY全体：GoogleSheet.DataTemplate](/GoogleSheet/DataTemplate.cls)


### JSONデータの作成：テストメソッド内で直接文言指定

- IRISにログインしたときの実行例
    ```
    set jsonobj=##class(GoogleSheet.SheetTemplate).CreateJSON()
    //JSONデータの確認
    set f=##class(%JSON.Formatter).%New()
    do f.Format(jsonobj)
    ```

- irispythonでPythonシェルを立ち上げた場合の実行例

    irispythonコマンドの実行
    ```
    cd <IRISインストールディレクトリ>/bin
    irispython
    ```
    [GoogleSheet.SheetTemplateクラス](/GoogleSheet/SheetTemplate.cls)のCreateJSON()メソッドの実行は以下の通りです。
    ```
    import iris
    jobj=iris.cls("GoogleSheet.SheetTemplate").CreateJSON()
    f=iris.cls("%JSON.Formatter")._New()
    f.Format(jobj)
    ```

テストメソッドではサンプルデータを直接設定していますが、テーブルやグローバル変数から取得して当てはめることもできます。

### JSONデータの作成：データベース（グローバル変数／テーブルデータ）から値を取得して設定

以下メソッドを実行すると、グローバルとテーブルのサンプルデータを作成します。

- IRISにログインしたときの実行例
    ```
    do ##class(GoogleSheet.SheetTemplate).CreateDummyData()
    ```

- irispythonでPythonシェルを立ち上げた場合の実行例

    irispythonコマンドの実行
    ```
    cd <IRISインストールディレクトリ>/bin
    irispython
    ```
    [GoogleSheet.SheetTemplateクラス](/GoogleSheet/SheetTemplate.cls)のCreateDummyData()メソッドの実行は以下の通りです。
    ```
    import iris
    iris.cls("GoogleSheet.SheetTemplate").CreateDummyData()
    ```


作成したサンプルデータからJSONを生成するサンプルコードは [CreateJSON2()](/GoogleSheet/SheetTemplate.cls#L53)をご参照ください。

- IRISのターミナルでの実行例
    ```
    set jsonobj=##class(GoogleSheet.SheetTemplate).CreateJSON2()
    //JSONデータの確認
    set f=##class(%JSON.Formatter).%New()
    do f.Format(jsonobj)
    ```

- irispythonでPythonシェルを立ち上げた場合の実行例

    irispythonコマンドの実行
    ```
    cd <IRISインストールディレクトリ>/bin
    irispython
    ```
    [GoogleSheet.SheetTemplateクラス](/GoogleSheet/SheetTemplate.cls)のCreateJSON2()メソッドの実行は以下の通りです。
    ```
    import iris
    jobj=iris.cls("GoogleSheet.SheetTemplate").CreateJSON2()
    f=iris.cls("%JSON.Formatter")._New()
    f.Format(jobj)
    ``````

### PythonのサンプルコードからGoogleSheetの複数シートを更新する例

Pythonシェルで以下実行します。（sys.pathに追加するディレクトリは実行環境用に修正してお試しください）

PythonスクリプトからEmbedded PythonでIRISのメソッドを呼び出します。

PythonスクリプトからIRISのメソッドを呼び出すために、sys.pathにIRISの以下専用ディレクトリの追加を行います。

専用ディレクトリは、[config.ini](/config.ini)に初期値を設定しています。IRISの実行環境に合わせて修正してください。

[config.ini](/config.ini)の中身は以下の通りです
```
[section1]
mgr = C:\\InterSystems\\IRISHealth1\\mgt\\python
lib = C:\\InterSystems\\IRISHealth1\\lib\\python
```
- mgr には、IRISのインストールディレクトリ以下\mgrディレクトリをフルパスで指定してください。
- lib には、IRISのインストールディレクトリ以下\libディレクトリをフルパスで指定してください。

設定が完了したら、Pythonシェルで以下実行すると、GoogleSheetにサンプルデータが更新されます。

```
import sys
sys.path+=['C:\WorkSpace\GoogleSheet']
import GSTest
spreadsheetId="1YbPs8yJRiNCrMi7NlHDenAPqXATZYFAWM0jBMDB-8Js"
GSTest.updateData(spreadsheetId)
```
更新後のイメージは以下の通りです。
![](/assets/afterUpdate.png)


シートをクリアする場合は、clearSheet()関数を実行します。

> 第2引数は「シート名!A1:C4」のようにシート名と範囲を指定します。シート全体をクリアする場合は「シート名」を指定します。

```
GSTest.clearSheet(spreadsheetId,"sales!A1:C4")
GSTest.clearSheet(spreadsheetId,"sales2!A1:C3")
```