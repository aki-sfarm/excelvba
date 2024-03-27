## Azure OpeAIを使います。
このVBAはAPIコール先をAzureOpenAIにしています。
OpenAI社など他のAPIを使う際は、変数endpointのURIを変更してください。


## JSONConverterのインポート

VBA内でJSONConverter.basを利用するため、JSONConverterのインポートが必要です。

### ステップ 1: JSONパーサーの取得

JSONを解析するために必要な`JSONConverter.bas`をGitHubからダウンロードします。

1. [VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)にアクセスします。
2. `JSONConverter.bas`ファイルを探してダウンロードします。

### ステップ 2: ExcelにJSONパーサーを追加

ExcelにダウンロードしたJSONパーサーを組み込む手順です。

1. Excelを開いて、必要なブックを開くか新しいブックを作成します。
2. `Alt` + `F11` キーを押してVBAエディタを開きます。
3. メニューから「ファイル」→「インポートファイル...」を選択し、ダウンロードした`JSONConverter.bas`を選んでインポートします。
4. プロジェクトに`JSONConverter`モジュールが追加されたことを確認します。

### ステップ 3: 参照設定の追加

JSONパーサーを使用するために必要な参照設定を追加します。

1. VBAエディタのメニューから「ツール」→「参照設定...」を選びます。
2. 「Microsoft Scripting Runtime」を探してチェックを入れ、OKボタンをクリックします。

## OpenAI接続用変数の設定
VBAソースの説明の際に、APIキーが表示されないように、
Module3を作成し、その中で
Public Const OpenAIapiKey As String = ""

Module4を作成し、その中で
Public Const OpenAIresourceName As String = ""

を設定してください。
それぞれの変数にAPIキーとAzureOpenAIのリソース名を入力してください。
