# スプレッドシートバリデータ＆CSV出力

これはスプレッドシートからCSVを出力したり、バリデーションを簡単にするgoogle app scriptです。

# スプレッドシートへのデプロイ

下記で該当のシートのあるアカウントにOAuthでログインします

> clasp login

依存モジュールなどの取得
> npm install

.clasp.json

↑このファイルがトップに必要です。
```
{
    "scriptId":"ここにapp scriptで作ったスクリプトIDを指定してください",
    "rootDir": "./src"
}
```

※注意下記コマンドは指定したプロジェクトを書き換えます！既に存在している場合は十分注意して実行ください。
> npm run push

.envにscriptIdを準備しておけば上記の.clasp.jsonを作成してデプロイをするスクリプトを置きました。

> sh deploy.sh

上記で失敗する場合は

https://script.google.com/home/usersettings

で有効化が必要

# 使い方

App Scriptで「onOpen」でスプレッドシートにメニューを追加します（CSV出力用）

ただし↑上記の実行時に下記のプライベート情報へのアクセスを許可する必要があります。

![アプリが確認されていない](./resources/承認されていない.png)

settingsシートに設定を記載します。

設定シートの内容は下記


| # | 内容          | Assign       | 値                         | 備考 |
|---|-------------|--------------|-----------------------------------|----|
| 1 | ヘッダ行最終（0～）  | rowOffset    | 3                                 |    |
| 2 | ヘッダ列列最終(0～) | columnOffset | 2                                 |    |
| 3 | 格納フォルダID    | saveFolder   | xxxxxxxxxxxxxxxxx |    |
| 4 | カラム定義シート    | columnsSheet | columns                           |
| 5 | 追加情報 | omake | \<iframe src="something">\</iframe> |omake押下時にhtmlをレンダリング|
| 6 | ファイルPREFIX    | prefix | file_                           |


今は位置固定（D2、D3、D4、D5、D6、D7）です

出力＆型＆バリデーション定義はシート「カラム定義シート（上記の場合columns）」に記載

# 出力＆型＆バリデーション定義シート

| # | 名称            | 型      | validation                           |
|---|---------------|--------|--------------------------------------|
| 0 | 例１         | string | requireNotNull:                      |
| 1 | 例２ | string | requireNotNull:requireStringSize(50) |
| 2 | 例３           | string | requireNotNull:                      |
| 3 | 例４           | string | requireNotNull:                      |
| 4 | 例５           | string | requireNotNull:                      |

名称はバリデーションのみ使用しています。

型は出力時のフォーマット等に利用しています、現在下記。

| 名称 |意味|
|--------|---------------|
| string | ダブルクォートのエスケープ    |
| number | カンマを除去します     |
| date   | シリアル値を日付変換します |
| bool   | 特に何もしません      |

バリデーションは現在下記（複数ある場合は:でつなげます。）

=record_check(レンジ)

のような形式でバリデーション結果を返します

| 名称                            | 意味    |
|-------------------------------|-------|
| requireNotNull                | 必須    |
| requireStringSize(50)         | 文字列上限 |
| requiredNumericRange(0,100) | 数値範囲  |
| requiredRegexp(^[a-zA-Z]+$) | 正規表現合致  |


# CSV出力

extraメニューのCSVユーティリティで起動します

その後右のサイドバーとしてダイアログが出ますので、レンジ（行の範囲を指します、2-5のような指定）を指定してダウンロードCSVでCSVが出力できます。
