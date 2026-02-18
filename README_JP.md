# あなたを笑顔にする、フレンドリーなヘルパーDLL
**インストール不要**、**ActiveX不要**、**管理者権限不要**。
このDLLをVBAプロジェクトのフォルダに追加するだけで、クールなUI機能を利用できます。MS Accessでのみテストされていますが、すべてのVBA環境で動作するはずです。ACCDEでも動作します。

主な目的は、いくつかの .NET 機能をVBAプロジェクトにもたらすことです。あなたのプロジェクトを視覚的にも機能的にも際立たせましょう！
このプラグインは管理者権限もインストールも必要としないため、クライアントの管理者ポリシーを気にすることなくアプリを配布できます。

もちろん、最小限のコードで！

```diff
- 注意:
- DLLの読み込みに関するエラーが発生した場合、プロジェクトがある場所にbinフォルダがあることを確認し、vba_tools.dllを右クリック => プロパティ => ブロック解除（Unblock）を行ってください。binフォルダ内のすべてのDLLに対しても同様の操作を行ってください。
- これは進化中のプロジェクトです。バージョンによって関数名が変わることがあるため、最新版に更新する前にラッパーのテストを行ってください。
```

# 安全のために
オンラインからダウンロードしたファイルについては、以下のサイトを使用してマルウェアのチェックを行ってください。
https://virusdesk.kaspersky.com/
https://www.virustotal.com/

![OnlineScanner](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/vbatoolsIsSafe.png)

# 必要に応じてDLLのブロックを解除してください
![UnblockADllPicture](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/unblockADll1.png)

## 現在の進捗 / バグ修正
```VBA
	+ 27/01/2022: 64bit OfficeでのOutlookからの添付ファイル追加を修正しました。
				: MySQLのNuGetパッケージを更新しました。
				: WebApiはTLS 1.2を優先します。
				: ドラッグ＆ドロップのいくつかのバグを修正 + ドロップされたファイル数を表示（closeAfterSelection=falseが設定されている場合）。
				: x86とx64は個別にビルドされ、それぞれのzipファイルに圧縮されています。

	以前の更新:
	+ Toast（通知）がJSONオプションを受け取るようになりました。ダイアログフォームの軽微なバグを修正。
	+ すべての関数は適切なクラス名の下にグループ化されました。例：Dll.String.AllStringRelatedFunctions
	- 以前のコーディングをすべて更新したことを確認してください。

	+ 任意のウィンドウを透明にする: ウィンドウ全体を透明にするか、キーカラーを設定してそのピクセルのみを透明にします。
	+ IsKeyDown(keyCode): 指定されたキーが押されているかどうか、trueまたはfalseを返します。これは任意のイベントで使用でき、CTRLキーやShiftキーが押されているかを確認するのに便利です。

	+ object ExecuteScalar : (DatabaseType dbType, string connectionString, string sql) : SQLを実行し、結果セットの最初の行の最初の列、または空の文字列を返します。
	+ bool ExecuteNonQuery(DatabaseType dbType, string connectionString, string sql)	: 非クエリSQLコマンドが成功した場合にtrueまたはfalseを返します。
	+ string MySqlGetAvailableServerFromList(string[] connectionString)					: 接続文字列の配列を受け取り、最初に到達可能な接続文字列、または""を返します。
	+ bool MySqlServerIsReachable(string connectionString)								: 接続文字列を使用してMySQLサーバーに接続し、trueまたはfalseを返します。
```


## 何をするものですか？
いくつかの .NET コンポーネントと関数を提供することで、アプリケーションをよりユーザーフレンドリーにするのに役立ちます。VBAよりも視覚的かつ機能的にクールです！

## 使い方は？
基本的なVBAスキルが必要です！
サンプルフォルダから <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> Dll + bin フォルダ </a> をダウンロードし、プロジェクトフォルダに追加してください。ACCDBにはサンプルが含まれており、コピーしてVBAアプリケーションに貼り付けることができます。

**すべての機能を楽しむために、VBA_TOOLS.Dll と Bin フォルダをVBAプロジェクトと同じ場所に置いてください。**



# 興味深い機能リスト
(まだリリースされていない新機能も含まれます)
<ul>
	<li>コンテキストメニュー (ContextMenu)
		<ul>
			<li>クールなコンテキストメニューを表示</li>
		</ul>
	</li>
	<li>バーコードコントロール (Barcode Control)
		<ul>
			<li>現在 Code39, Code128, QrCode をサポート</li>
			<li>現在テスト中</li>
		</ul>
	</li>
	<li>色 (Colour)
		<ul>
			<li>HexToAccess (16進数からAccessカラーへ)</li>
			<li>AccessToHex (Accessカラーから16進数へ)</li>
		</ul>
	</li>
	<li>DBクラス: (MySQLのみサポート)
		<ul>
			<li>ExecuteScalar	: MySQLのSELECTクエリを実行し、最初の行の最初の列のオブジェクトを返します</li>
			<li>ExecuteNonQuery	: MySQLのUPDATE/INSERTクエリを実行し、成功した場合にtrueまたはfalseを返します。</li>
			<li>string MySqlGetAvailableServerFromList(string[] connectionString)</li>
				<ul>接続文字列の配列を受け取り、最初に到達可能な接続文字列、または""を返します。<br/>複数のバックエンドサーバーがあり、どれが到達可能かを知りたい場合に便利です。スレッドを使用しているため高速です！</ul>
			<li>bool MySqlServerIsReachable(string connectionString)</li>
				<ul>: 接続文字列を使用してMySQLサーバーに接続し、trueまたはfalseを返します</ul>
		</ul>
	</li>
	<li>ダイアログボックス (Dialog Boxes)
		<ul>
			<li>クールなダイアログボックス</li>
			<li>拡張クールダイアログボックス</li>
			<li>クールでシンプルなダイアログボックス</li>
			<li>シンプルな確認ダイアログ (Are you sure?)</li>
			<li>ドラッグ＆ドロップ OpenFileDialog
			<ul>
				<li>ドラッグをサポートするシンプルなファイルを開くダイアログ</li>
			</ul>
			</li>
		</ul>
	</li>
	<li>ディスプレイ (Display)
	<ul>
		<li>GetNumberOfMonitors
		<ul>
			<li>モニター数を返します</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>
	<ul>
		<li>GetPrimaryMonitorHandle
		<ul>
			<li>プライマリモニターのハンドルを返します</li>
		</ul>
		</li>
		<li>GetPrimaryMonitorBounds
		<ul>
			<li>プライマリモニターに関する情報を取得します</li>
		</ul>
		</li>
		<li>GetMonitorBoundsByHandle
		<ul>
			<li>ウィンドウハンドルからモニター情報を返します</li>
		</ul>
		</li>
		<li>GetCursorPosition (カーソル位置を取得)</li>
	</ul>
	</li>
	<li>フォーム (Form)
	<ul>
		<li>透明ウィンドウの作成</li>
		<li>Access背景色の変更</li>
		<li>Drage Me
		<ul>
			<li>枠のないフォームを移動できるようにします</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>FTP
	<ul>
		<li>FTPS_UPLOAD
		<ul>
			<li>ファイルを指定されたFTPサーバーにアップロードします</li>
		</ul>
		</li>
		<li>FTPDeleteFile
		<ul>
			<li>指定されたFTPロケーションからファイルを削除します</li>
		</ul>
		</li>
		<li>FTPFileExists
		<ul>
			<li>指定されたファイルが指定されたFTPロケーションに存在するかチェックします</li>
		</ul>
		</li>
	</ul>
	</li>
	<li></li>
	<li></li>
	<li>グラフィックス (Graphics)
	<ul>
		<li>URLから画像コントロールへ保存せずに画像を読み込む</li>
		<li>Web URLを読み込む</li>
		<li>ローカル画像を読み込む</li>
		<li>BLOBフィールドを画像に変換する</li>
		<li>MeasureText
		<ul>
			<li>指定されたテキストの高さと幅をピクセル単位で計算します</li>
		</ul>
		</li>
		<li>ByteToImage
		<ul>
			<li>バイト配列を画像に変換し、指定された場所に保存します</li>
		</ul>
		</li>
		<li>ByteToBitmap
		<ul>
			<li>バイト配列をビットマップに変換します</li>
		</ul>
		</li>
		<li>TakeScreenShotFromHwnd
		<ul>
			<li>指定されたハンドルによってウィンドウのスクリーンショットを撮ります。バイト配列を返します</li>
		</ul>
		</li>
		<li>TakeScreenShot
		<ul>
			<li>デスクトップ全体のスクリーンショットを撮ります。バイト配列を返します</li>
		</ul>
		</li>
		<li>TakeScreenShot1
		<ul>
			<li>スクリーンショットを撮り、指定されたパスに保存します</li>
		</ul>
		</li>
		<li>PictureFromUrl
		<ul>
			<li>URLから画像を読み込み、バイト配列を返します</li>
		</ul>
		</li>
		<li>SaveClipboardToImage
		<ul>
			<li>クリップボードの画像を指定されたパスと形式で保存します</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>インプットボックス (InputBoxes)
	<ul>
		<li>クールなインプットボックスの表示</li>
		<li>検証付きメールアドレス</li>
		<li>パスワード</li>
		<li>複数行 / 単一行</li>
		<li>数字のみ</li>
		<li>検証付き日付</li>
		<li>ドロップダウンボックスの表示</li>
	</ul>
	</li>
	<li>JSON
	<ul>
		<li>Newtonsoft.Jsonを使用</li>
		<li>ExportToJSON
		<ul>
			<li>MS Accessユーザーがクエリ、テーブルのSQL結果をJSON文字列としてエクスポートできるようにします</li>
		</ul>
		</li>
		<li>ImportJSON
		<ul>
			<li>MS AccessユーザーがJSON文字列配列を使用してテーブルにレコードをインポートできるようにします</li>
		</ul>
		</li>
		<li>JSONString
		<ul>
			<li>Newtonsoft.Jsonオブジェクトの文字列表現</li>
		</ul>
		</li>
		<li>JSONGetObject
		<ul>
			<li>プロパティ名でJsonオブジェクトを取得</li>
		</ul>
		</li>
		<li>JSONSetObject
		<ul>
			<li>Jsonオブジェクトにプロパティを追加</li>
		</ul>
		</li>
		<li>JSONGetValue
		<ul>
			<li>jsonオブジェクトから値を取得</li>
		</ul>
		</li>
		<li>JSONToObject
		<ul>
			<li>文字列jsonを動的オブジェクトに変換</li>
		</ul>
		</li>
		<li>JSONSerialize
		<ul>
			<li>動的jsonオブジェクトの文字列表現を返します</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>NetClass
	<ul>
		<li>UrlIsReachable
		<ul>
			<li>URLが到達可能かどうか、trueまたはfalseを返します</li>
		</ul>
		</li>
		<li>UrlIsValid
		<ul>
			<li>URLの形式が正しいかどうか、trueまたはfalseを返します</li>
		</ul>
		</li>
		<li>UrlIsLocalPath
		<ul>
			<li>指定されたURLはローカルファイルパスか？</li>
		</ul>
		</li>
		<li>GetExternalIP</li>
	</ul>
	</li>
	<li>通知 (Notification)
	<ul>
		<li>ノンブロッキング通知の表示</li>
		<li>成功の表示</li>
		<li>警告の表示</li>
		<li>エラーの表示</li>
	</ul>
	</li>
	<li>プログレスバー (ProgressBar)
	<ul>
		<li>クールなプログレスバーの表示</li>
	</ul>
	</li>
	<li>正規表現 (RegEx)
	<ul>
		<li>IsMatch</li>
		<li>GetFirstMatch</li>
		<li>Replace</li>
	</ul>
	</li>
	<li>文字列 (String)
	<ul>
		<li>PadLeft</li>
		<li>PadRight</li>
		<li>TrimEnd</li>
		<li>TrimStart</li>
		<li>StartsWith</li>
		<li>EndsWith</li>
		<li>.NET string.format</li>
	</ul>
	</li>
	<li></li>
</ul>


### [ノンブロッキング通知の表示]
Toastr (https://github.com/CodeSeven/toastr) に触発されました。
VBAユーザーが待機したり、VBAアプリケーションに負荷をかけたりすることなく、シンプルな通知を表示できるようにします。
簡単なコマンドで、フォーカスを奪ったりユーザーを邪魔したりすることなく、メッセージ付きの小さなカラフルな通知がポップアップします。
私は主に、メールの到着やタスクの完了など、アクションを必要としないメッセージを表示するためにこれを使用しています。

![just a notification](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/information.png)

## 通知を好きなようにカスタマイズ:
現在、以下のカスタマイズが可能です。
```
1.Message   : ハイパーリンク用の <a href="">テキスト</a> を含めることができます（その他のHTMLタグは無視されます。ハイパーリンクは www または http または https で始まる必要があります（基本的には整形式のリンクのみ？）
```
```diff
+ ローカルファイルを開くことができるようになりました。ハイパーリンクはローカルファイル形式である必要があります。例: <a href="F:\folderName\picture.png">
+ テスト中: コールバックコマンド（Docmd.OpenFormのみ）をハイパーリンクに埋め込むことができます
```
```
2.Duration (ミリ秒) (デフォルト 2000。0にすると長時間通知を維持します。int.max)
3.Background colour (HTMLカラーコード)
4.Font colour (HTMLカラーコード)
5.X,Y position (デスクトップ上の位置)
```



![picture of 3 notifications](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-collections.png)
```VBA
'使用されるコマンド
Toastr.Toast "おっと、何か問題が発生しました！",vberror,0
Toastr.Toast "黄色の気象警報！",vbexclamation,0
Toastr.Toast "通知を受信しました",vbinformation,0
```

動作中の様子
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction.gif)
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction1.gif)

## ユーザーとの対話やハイパーリンクの表示はどうですか？
メッセージ内にHTMLの ```<a href="">text</a>``` タグを含めることができ、これらはハイパーリンクに変換されます。
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/Hyperlink.png)

ローカルファイルをハイパーリンクとして: ```<a href="F:\folderName\picture.png">この画像を表示</a>```

コールバックコマンド:
1. フォームを開くハイパーリンク: ```<a href="DoCmd.OpenForm frmImageView,acNormal,,wherecondition:=id=2">フォームを開く</a>```
```VBA
	注: docmdコマンドには " や ' を含めないでください
	Filter, WhereCondition, DataMode, WindowMode は名前付きパラメータである必要があります。例: Filter:=FilterCondition または WhereCondition:=id=2

	'同様に、ホストアプリケーションで実行される関数名を渡すこともできます
	<a href="ExecuteMe()"> ホストアプリケーション内の関数を実行 </a>



	[トースト通知]
	1. トースト通知のハイパーリンク解析機能を修正しました
	2. トーストで Docmd.OpenForm を開けるようになりました
	3. トーストでローカル関数を実行できるようになりました。例: <a href="ExecuteMe()"> ホストアプリケーション内の関数を実行 </a> は、リンクをクリックしたときに "ExecuteMe()" を実行します。
	4. トースト / シンプルダイアログボックスでパラメータ付きのローカル関数を実行できるようになりました 例: <a href="ExecuteMe('ParameterA','ParameterB')"> ExecuteMe </a>
	4. トースト / シンプルダイアログボックスで、ハイパーリンクをクリックした後に自身を閉じることができるようになりました。closeme="true" 属性を使用してください。例: <a href="ExecuteMe('ParameterA','ParameterB')" closeme="true"> 実行して閉じる </a>


```



## ダウンロード
サンプルをダウンロードして、あなたのプロジェクトでテストしてください。感想をコメントに残していただけると嬉しいです。
<a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> サンプル</a>


<hr>
<hr>

# クールなダイアログボックスの表示
標準のメッセージボックスは素晴らしいですが、標準機能以上のものが欲しくなることもあります。
例えば：
<ul>
  <li>色を付けたい</li>
  <li>3つ以上のボタンを持ちたい</li>
  <li>自動で閉じるようにしたい</li>
  <li>HTMLタグを使いたい</li>
  <li>ループでVBAアプリに負荷をかけたくない</li>
</ul>
VBAユーザー向けの新しい簡素化されたダイアログボックスをご紹介します。このダイアログボックスは上記の機能を可能にし、アプリケーションをカラフルに保つのに役立ちます。 :) この機能はまだ開発中であり、テスターからのフィードバックをお待ちしています。











<HR>

![Cool DialogBox](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-DialogboxGreen.png)
![Cool DialogBox1](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-DIALOG-BOX.png)

サンプルのaccdbにはVBAラッパーがあり、必要に応じて拡張できます。これはサードパーティのJSON Converterプラグインを使用しており、私の方でいくつか修正を加えています。

```
  'ラッパーを使用すると、これくらいシンプルになります
  Debug.Print gDll.DialogRich("これはタイトルです", "いくつかのコンテンツ", (vbExclamation + vbYesNo))
```

簡素化されたバージョン（HTMLレンダリングなし）も利用可能です。
# クールでシンプルなメッセージボックス
![Cool DialogBox](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-CoolSimpleMessageBox.png)

シンプルなメッセージボックスを表示できます


# クールなプログレスバーの表示
プログレスバー

進捗状況をユーザーに知らせる際の重要な要素です。このクールなプログレスバーを使えば、以下のようなシンプルなコードでいつでもアプリケーションの上にポップアップ表示できます。

```
  Dim ProgressBarID As Long
  ProgressBarID = gDll.ShowProgressBar(100, "クエリを実行中", "お待ちください。プリンタードライバーを準備しています")

  ProgressBarID = gDll.SetProgressBar(ProgressBarID, 10, "ドライバーを待機中..")

  gdll.CloseProgressbar ProgressbarId 'プログレスバーを閉じます
```
![Cool ProgressbarGreen](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-ProgressBar.png)

いつものように、好みに応じてテーマカラーを変更できます。
![Cool ProgressbarRed](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-ProgressBarRed.png)

### 注意:
```ShowProgressBar と SetProgressBar``` はIDを返し、これを使ってプログレスバーを参照できます。これにより、VBAユーザーは同時に複数のプログレスバーを持つことも可能です。

# クールなインプットボックスの表示
InputBoxも頻繁に使用されるコンポーネントです。システムの素朴なInputBoxを好む人もいますが、私たちはモダンなUIカラーが大好きです :)
これらのテーブルからどれを選びますか？

![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/InputBoxDefault.png)  ![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-InputBoxMultiline.png)

## いい色ですね！でも何の意味があるの？
新しいInputBoxにはいくつかの組み込み関数があり、それに応じて設定できます。
現在、以下のタイプがサポートされています。
```
'        Password        = 1, : システムパスワードマスクを使用
'        Text            = 2, : 単一行テキスト
'        MultilineText   = 32, : 複数行テキストボックス
'        Number          = 4, : 数字のみ
'        ShortDate       = 8, : dd/mm/yyyy でマスク。日付は終了時に検証されます
'        LongDate        = 16,  : dd/Month/yyyy でマスク
'        DateTime        = 48,  : dd/mm/yyyy hh:mm:ss でマスク

そして以下のパラメータを受け入れます:
  Type以外はすべてオプションです

  InputBoxType Type,    : 数値
  string Title,         : インプットボックスのタイトル
  string Message,       : インプットボックスのオプションテキスト
  int PosX,             : このボックスを配置する画面に対する相対的なX座標
  int PosY,             : このボックスを配置する画面に対する相対的なY座標
  string ThemeBg,       : HTMLカラーコード
  string ThemeForeColour: HTMLカラーコード

' DLLがあれば、以下のように使用します

  result = gDll.DLL.showinputbox(Type:=32, Title:="", Message:="その日に何が起きたか教えてください！", ThemeBg:="", ThemeForeColour:="")
```
#### カーソルのX,Y位置を返す getCursorPosition 関数もチェックしてみてください！


動作中の様子:

![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/InputBox.png)

いつものように、テーマカラーを変更できます :)

![purple input box](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-InputBoxPurple.png)

ダウンロード <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> サンプル</a>

# [ドロップダウンボックスの表示]
VBA_TOOLSユーザーからのリクエストです。他のユーザー入力ボックスと同様に、クールなドロップダウンボックスから選択させることができるようになりました。
![purple input box](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/DropDownBox.png)

クールなドロップダウンボックスを作ろうと決めたとき、既存のMS Accessのドロップダウンボックスを使わない理由について考えました。
個人的には、ドロップダウンボックスにアイコンを表示するのは素晴らしいアイデアだと思います！ :) さらに、標準のDropDownBoxでは、コンテンツ内の部分一致検索、つまりドロップダウン選択肢の任意の部分を検索することができません。
従来のDropDownのように、既存のクエリやテーブルのリストを表示したいです。
そこで、今回はそれらのポイントをカバーすることにしました。もちろん、将来的には以下の機能を追加する予定です。

1. グループ化されたエントリ: ドロップダウンのエントリがグループヘッダーでグループ化されます。
2. 複数の画像列を持つ？
3. ドロップダウンスタイルを完全に変更: サブメニュー付きのメニューのように...
4. エラーの修正


使い方は以下の通りです:
```VBA
 ?gDll.ShowDropDown("アイテムを選択", "何か内部メッセージ？", "qryDropDown", 2, Array(50, 50))


 パラメータリスト:
 /// <summary>
  /// 選択用のドロップダウン付きダイアログボックスを表示します。文字列値を返します
  /// </summary>
  /// <param name="title">インプットボックスのタイトル</param>
  /// <param name="message">インプットボックスの内部メッセージ</param>
  /// <param name="dbSource">データベースパス</param>
  /// <param name="tableSource">テーブル名またはSQL。SQLを使用する場合は isRawSql=true を使用</param>
  /// <param name="boundColumn">値を取得する列インデックス</param>
  /// <param name="columnWidths">整数の配列</param>
  /// <param name="isRawSql">テーブルソースがプレーンなSQLコマンドかどうかを指定します</param>
  /// <param name="posX"></param>
  /// <param name="posY"></param>
  /// <param name="themeColour"></param>
  /// <param name="themeForeColour"></param>
  /// <returns>文字列値</returns>
```


### 注意:
データソースの最初の列に "icon" が含まれており、それがハイパーリンク（Webまたはローカルファイル）である場合、デフォルトでそれらのリンクは画像に変換されます。
Array(column0_width, column1_width ...) を使用して列幅を設定します


<hr>

# ドラッグ＆ドロップ OpenFileDialog
何だって！！VBAでドラッグ＆ドロップ機能？？はい、正しく読みましたよ。でもあまり興奮しないでくださいね :) これは単なるファイルドロップ機能です。ドラッグ＆ドロップ方式でファイルを選択/開く/取得できるようにします。既存の FileOpenDialog メソッドの直接的な代替手段です。
<hr>

### 選択されたすべてのファイルを含むJSON配列の文字列を返します。（文字列配列をご希望の場合は下記参照）
それらのファイルパスをどうするかはあなた次第です。おそらく将来的には、これを既存のFTPコンポーネントとリンクさせるかもしれません。


現在、以下のパラメータが受け入れられています:

```c#
  以下はすべてオプションです。

  string Message,         : ダイアログボックスへのメッセージ。
  bool AllowMulti,        : 複数ファイルを許可しますか？
  string[] Filters,       : 文字列の配列 => (説明 |*.png)。ファイル拡張子フィルタに使用されます
  int PosX,               : このボックスが表示されるべきモニターに対する相対的なX位置
  int PosY,               : このボックスが表示されるべきモニターに対する相対的なY位置
  string ThemeBg,         : HtmlColourCode
  string ThemeForeColour  : HtmlColourCode
```
(DLL部分は完了していると仮定:) VBAでは以下のように使用します:
または、サンプルファイルをダウンロードして、どの関数をアプリケーションにコピーするかを確認してください。

```
    Dim FilePaths As String
    FilePaths = gDll.DLL.ShowDialogForFile("複数ファイルは許可されていません", False)
```

またはカスタマイズしたもの:
```VBA
    Dim Filters(2) As String

    Filters(0) = "Png画像のみ |*.png"
    Filters(1) = "すべてのファイル |*.*"

    Dim FilePaths As String
    FilePaths = gDll.DLL.ShowDialogForFile(Message:="たくさんのファイルを自由にドロップしてください", allowmulti:=False, Filters:=Filters, PosX:=0, PosY:=0, ThemeBg:="", ThemeForeColour:="")

```
文字列配列の結果が欲しい場合
```
    dim Files() as string
    Files = gDll.ShowDialogForFileArray(Message:="たくさんのファイルを自由にドロップしてください", allowmulti:=False, Filters:=Filters, PosX:=0, PosY:=0, ThemeBg:="", ThemeForeColour:="")
    '選択されたファイルの文字列配列を返します
```

動作確認:
![File drag and drop gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/FileDropInAction.gif)

エラー
![File drag and drop error gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-DRAG-DROP.gif)
<hr>






# URLから画像コントロールへ保存せずに画像を読み込む
おお！これがそのまま可能であることをどれだけの人が望んだでしょうか？私たちの多くは、良いチュートリアルを探すのにかなりの時間を費やし、その結果のほとんどは解決策というよりは単なる回避策でした。APIやクラスを使った何ページものコード、あるいはWebブラウザコントロールの使用、サードパーティの画像コントロールの購入、あるいは画像をダウンロードしてから再度読み込むなど。

Webブラウザコントロールを悪く言うつもりはありません。それはそれで素晴らしいですが、画像の表示用に設計されていないのは確かです（個人的意見）。ズームやストレッチなどの機能は、Webブラウザコントロール経由では利用できません。もちろんHTMLタグを使うこともできますが、それは「回避策」の問題に対する別の「回避策」ですよね？

インストールが必要なのでサードパーティのコントロールは買いたくない！（多くの人にとってNG）
ダウンロードして読み込むのも嫌だ。後始末が面倒すぎる。

画像コントロールに画像を読み込むことができる、私たちのシンプルな1行コードをご紹介しましょう。ダウンロードなし、長すぎるコードなし、ナンセンスなし。

```VBA
  'Dll関数
  'PictureFromUrl(
    string URL,             :  画像URL。Web URLまたはローカルパス
    bool ShowError = false, : URLが読み込めない場合にエラー通知を表示する
    long sender = 0         : 送信者HWND、現在は使用されていません。
    )

  'VBAラッパー (簡単にするために使用)
  'ImageControlGetImage(ImagePath as string, optional ShowError=true)


'Web URLの読み込み
Private Sub Command147_Click()
    Dim WebPicture As String
    WebPicture = "https://avatars2.githubusercontent.com/u/1001697?s=460&v=4"

    Me.Image113.PictureData = gDll.ImageControlGetImage(WebPicture, ShowError:=True)
End Sub

'ローカルファイルパスを読み込むために同じ関数を使用
Private Sub Command149_Click()
    Dim WebPicture As String
    WebPicture = "F:\Projects\VBA_DLL\dialogboxgreen.png"

    Me.Image113.PictureData = gDll.ImageControlGetImage(WebPicture, ShowError:=True)

End Sub

```
動作確認:
![Image from web url](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/ImageControlInAction.gif)

### テーブルからURLを読み込みたい場合
`コントロールソース` プロパティを使用する代わりに、フォームの `レコード移動時 (On Current)` イベントを使用して画像を読み込みます。
```VBA
Private Sub Form_Current()
  '画像を読み込む
    Me.Image8.PictureData = gDll.ImageControlGetImage([url], True)
End Sub
```
楽しんでください、そして感想を聞かせてください！


# VBA用バーコードコントロール
バーコードを表示できるようにしてほしいというVba_toolsユーザーからのもう一つのリクエストです。私はバーコードについて全く知りませんでしたが、Googleで素晴らしいソースを見つけました (https://sourceforge.net/projects/zintnet/)。zintnetのオーナーに感謝します。
いくつかのクラスを適応させ、私たちのVBA_TOOLSプラグインに追加しました。

他のコンポーネントとは異なり、バーコードコントロールはフォームやレポートに埋め込まれるため、コントロールはスタンドアロンフォームにはなれません。レポートや請求書を印刷するときにバーコードが表示されるようにするためです。これを実現するために、.NET環境でバーコードを作成し、そのバーコードを画像としてAccessに戻します。これにより、フォームやレポート上の画像コントロールでバーコードを表示できます。

これもベータ版です。ご覧になって、感想をお知らせください。

使い方は？

```VBA
  Me.imgBarcode.PictureData = gDll.CreateBarcode(Val(Me.BrcodeType.value), Me.txtBarcodeData.value, Val(BarcodeSizeMultiplier.value))

  パラメータリスト:
  '    /// <summary>
'    ///
'    /// </summary>
'    /// <param name="symbology">バーコードの種類</param>
'    /// <param name="barcodeData">バーコードのデータ値</param>
'    /// <param name="width">グラフィックス / 画像の幅</param>
'    /// <param name="height">グラフィックス / 画像の高さ</param>
'    /// <param name="multiplier">サイズをこの値で乗算します。</param>
'    /// <returns>画像データ</returns>
'    CreateBarcode(Symbology symbology, string barcodeData, int width, int height, float multiplier )

```
![qrBarcode.png](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/qrBarcode.png)
![Code39Barcode.png](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/Code39barcode.png)

# VBA用クールなコンテキストメニュー
[テスト中]: テスター募集中。この機能は現在ラップされていません。つまり、公開の準備ができたらさらにパラメータが追加される予定です。

何と言えばいいでしょうか？おそらくほとんどのVBAユーザーは、このコントロールが最初から存在していればいいのにと願っていたことでしょう。右クリックのコンテキストメニューと同様に、左クリック可能なコンテキストメニューを開発しました。
もちろん、スタイル変更可能で、移動可能で、アイコン付きのメニュー項目です。

説明は十分でしょう。動作を見てみましょう。
![ContextMenuPicture](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-CoolContextMenu.gif)

注目していただいたところで、このコントロールがどのように機能するか見てみましょう。2つのバージョンがあります。
1>シンプル: メニュー項目として文字列の配列を受け取り、選択された値としてメニュー項目のテキストを返します。
2> 拡張: 拡張メニューでは、アイコンを表示したり、各メニュー項目のデータ値を受け取ったりできます。つまり、各メニュー項目は配列(string:アイコンパス, string:データ戻り値, string:メニュー項目)として構築されます。例: array("c:\email.png","1","メール送信")
拡張メニューを使用するには、各メニューエントリを別の配列内に追加します。例: Array(array("c:\email.png","1","メール送信"), array("c:\door.png","2","終了"), array("http://someweblink.png","3","Webアイコン付きメニュー項目") )

コードでは以下のようになります:
```VBA
	'シンプルメニュー
	'メニュー項目の配列を作成
    Dim MenuItems() As String
	Dim result As String

    MenuItems = VBA.Split("Do Something,I'm so cool, Send Email, Print, Settings, Save, Save As, Make Pdf", ",")
    FnArrayAddItem MenuItems, "Exit"
    FnArrayAddItem MenuItems, "Exit Application"

    result = gDll.DLL.ShowContextMenu(MenuItems)
	gDll.Toast result, , , Me.hwnd
    If (result = "Exit") Then
        DoCmd.Close acForm, Me.Name, acSaveYes
    ElseIf (result = "Exit Application") Then
        Application.Quit
    End If

```

```VBA
	'アイコン付き拡張メニュー
	Dim result As String
    result = (gDll.DLL.ShowContextMenuA(Array(Array("", "0", "Web loading takes time"), Array("F:\PROJECT_SUPPORT\Images\csharp.png", "1", ".NET is cool"), Array("https://static.thinkster.io/topics/node_icon.png", "2", "Loading icon from web"), Array("glyphicons-389-exit", "3", "Exit"))))


    gDll.Toast result, vbInformation

    If (Val(result)) = 3 Then
        DoCmd.Close acForm, Me.Name, acSaveYes

```






<hr>
<hr>



# その他の興味深い機能

# DragMe
枠のないフォームをドラッグできるようにするシンプルな機能です。
こちらをご覧ください。
![DragME](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-DragMe.gif)

使い方は？
```VBA
	'シンプルにmouseDownイベントを使用します

	Private Sub Label251_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
		Call gDll.DLL.DragMe(Me.hwnd)
	End Sub

```

### AreYouSure? (よろしいですか？)
シンプルなYes/Noポップアップで、trueまたはfalseを返します。ユーザーにYesかNoのアクションを確認したいだけの時があります。
これは単純ですがクールなYes/Noボックスかもしれません :)
```VBA
? gDll.AreYouSure

```
Hexカラーコードを提供するか、ブートストラップクラスを使用している場合は、AreYouSureBoxのテーマカラーを設定することも可能です。
```VBA
? gDll.AreYouSureE(Me, "#aa66cc", "#000000", "#aa66cc", "#F65656") or
?gDll.AreYouSureE(, gBootstrap.default_color_dark, gBootstrap.WHITE, gBootstrap.AMBER, gBootstrap.TEAL_LIGHTEN_3)
```
![AreYouSureCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-AreYouSureCollection.png)

### ファイルをダウンロードし、VBA用プログレスバーを表示
もう一つのクールな機能です。この機能を使用すると、インターネットからファイルをダウンロードし、上記のクールなプログレスバーを使用してダウンロードの進捗状況を表示できます。

```DownloadedFile = DLL.DownloadAFile(Url, [Destination], [OverWrite = true], [ShowProgress = true])```
Url以外、他のすべてのパラメータはオプションです。Destination（保存先）が提供されない場合、ファイルは application.path に保存されます。

### クリップボードの画像をローカルファイルに保存
時々、単純なことがVBAでは非常に難しいことがあります。クリップボードの画像をローカルパスに保存したい場合は、この機能をチェックしてください。

``` SaveClipboardToImage(string PathToSave, string FileName, string ImageType) ``` すべてのパラメータはオプションで、デフォルトではJpeg画像タイプが使用されます。クリップボードオブジェクトに画像が含まれている場合、好きな場所に保存され、ファイルパスが返されます。
サンプルのaccdbに ```SaveClipboardToImage``` というラッパーがありますので、チェックしてみてください。

### PadLeft と PadRight
.NETの padleft と padRight 関数を使用します。
``` gdll.DLL.PadLeft("1",10,"0") => 0000000001
    ?gdll.DLL.PadRight("1",10,"0") = > 1000000000
```
### カーソルのX,Y位置を返す getCursorPosition 関数もチェックしてみてください！

### [SignalRメッセージの受信]
これは独自のSignalRサーバーを持っている私のために機能しますが、一般的には開発中、あるいはまだ準備ができていないと言えます！
Googleプッシュメッセージやその他のプッシュメッセージサービスのようなものです。ログインしているすべてのユーザーに一箇所から通知を送信できます。
これを拡張すれば、ログインしているすべての参加者がメッセージを送受信できるチャットサーバーとして使用することもできます。
繰り返しになりますが、VBAアプリに負荷をかけることはありません。


### ByteToImage
ByteToImage(byte[] byteArraym string TemporaryPath, bool useCache) はMS Accessユーザー向けの関数です。基本的には、データベースから受け取ったバイト配列を画像に変換できます。
画像ファイルのパスを返します。このパスを画像プロパティの場所として使用してください。
Me.Image32.Picture = gDll.ByteToImage(ByteArray, "SaveLocationPath") のようになります。

### FTP(S) アップロード
WinScpを使用してホストにファイルを安全にアップロードするシンプルなツールです。VBAを使いすぎたりActiveXコンポーネントを持ったりせずにファイルをアップロードしたい場合に便利です。
```VBA
	'シンプルに以下のように使用します
	Debug.Print gDll.FTPUploadFile(ServerName, port, Username, Password, "F:\Projects\VBA_DLL\Modern Inputbox for vba purple.png", "/screenshots/", SSHFingerprintOfTheRemoteServer, Ftp, Explicit, False)

	パラメータリスト
```
```C#
		/// <summary>
        /// 指定されたftpサーバーにファイルをアップロードします
        /// </summary>
        /// <param name="host">ホストサーバー</param>
        /// <param name="port">ポート番号</param>
        /// <param name="username">FTPユーザー名</param>
        /// <param name="password">FTPパスワード</param>
        /// <param name="localFileName">ローカルファイルへのパス</param>
        /// <param name="remoteLocation">リモートサーバー内の場所</param>
        /// <param name="hostCertificateFingerprint">リモートサーバーのフィンガープリント</param>
        /// <param name="protocol">Ftpプロトコル、ftp、sftp...</param>
        /// <param name="ftpSecure">接続の種類、implicit、explicit</param>
        /// <param name="giveUpSecurityAndAcceptAnyTlsHostCertificate">デバッグ使用のみ</param>
        /// <returns>trueまたはfalseの文字列表現、またはエラーメッセージ</returns>
````

### FTP リモートファイル削除
リモートサーバーからファイルを単純に削除します。成功した場合はtrueまたはfalse + エラーメッセージを文字列として返します。
```VBA
  'Server as string
  'Port as number
  'Username as string
  'Password as string
  'RemoteFile as string
  'SSHFingerprint as string
? DLL.FTPDeleteFile(ServerName, Port, Username, Password, RemoteFile, TLSHostFingerprint)
```

### ImportJSON (JSONインポート)
Application.ImportXMLに少し似ていますが、JSON配列文字列からAccessテーブルを作成できます。
大規模なテストは行っていませんが、私のニーズには合っています。

シンプルに呼び出します
```VBA
  gdll.ImportJSON(YourJSonArrayString, "ターゲットテーブル名", ImportOptions[append,structureOnly,structureAndData], recreate)
 'Recreateはテーブルを削除して再作成します。AppendOnlyが要求された場合、recreateは無効です

 'これが動作するサンプルコマンドです。tblJsonTestという新しいテーブルを作成し、配列からすべてのコンテンツをインポートします。
 gdll.ImportJson("[{""id"":10,""name"":""User"",""add"":false,""edit"":true,""authorize"":true,""view"":true},    {""id"":11,""name"":""Group"",""add"":true,""edit"":false,""authorize"":false,""view"":true},    {""id"":12,""name"":""Permission"",""add"":true,""edit"":true,""authorize"":true,""view"":true}]","tblJsonTest",acStructureAndData,True)
  '
```

### ExportToJson (JSONエクスポート)
テーブルの内容をJSON文字列としてエクスポートできるようになりました。
方法1:
```VBA
  'SQL SELECTコマンドを実行し、結果セットをJSON形式の文字列として保存します。
  gdll.ExportToJSON("select * from tbljsontest where authorize = true;","MyJson.Txt",overwrite:=false,isRawSql:=true)
```

方法2:
```VBA
  'テーブル/クエリからすべてをエクスポート
  gdll.ExportToJSON("tbljsontest ",SaveAs:= "MyJson.Txt",overwrite:=false,isRawSql:=false)
```
この方法では、テーブル名/クエリ名をエクスポート関数に渡し、isRawSql = false を設定しました。エクスポート関数は “SELECT * FROM givenTableName/QueryName;” に似たSQLステートメントを生成し、JSONエクスポートを実行します。

SaveAs（ターゲットファイル名）が空の場合、ファイルはエクスポートされませんが、変換は行われ、変換された文字列が結果として返されます。

サンプルプロジェクトをダウンロードして遊んでみてください。


# [今後の機能]
たくさん... :)
特定の機能が必要な場合は、メールまたはコメントを残してください :)

# 待ちきれませんか？今すぐダウンロードして楽しんでください！
<a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> サンプル</a>

# [著作権、ライセンス、クレジット]

Copyright © 2018 Krish

非商用目的であれば、dllを自由に使用できます。商用ユーザーの場合、1つの条件付きでdllを使用できます。あなたが誰であるかをお知らせください。私たちのクライアントリストにあなたの/会社名があることを非常に嬉しく思います。

私のGitHubページへのクレジットとリンクをいただければ幸いです。




<hr>
<hr>
<hr>
# Raw methods from class
<hr>
<hr>

```C#
/// <summary>
/// デスクトップにトーストを表示します
/// </summary>
public async void FN_SHOW_TOAST(string iMessage, int iDuration, string iBG_COLOR, long iANIME_DURATION, string iFONT_COLOR, int iX, int iY, int iANIM_DIRECTION, bool iAUTO_CLOSE = true)
{
}



/// <summary>
/// バイト配列を画像に変換し、指定された場所に保存します。
/// </summary>
/// <returns>ローカルに保存された画像のパス</returns>
public string ByteToImage(byte[] byteArrayIn, string iTempPath, bool useCache)
{
}

/// <summary>
/// バイト配列をビットマップに変換します
/// </summary>
public Bitmap ByteToBitmap(byte[] byteArr)
{
}

/// <summary>
/// 画像を含むバイト配列を返します
/// </summary>
/// <param name="hWND"></param>
public byte[] TakeScreenShotFromHwnd(long hWND)
{
}

/// <summary>
/// デスクトップ全体のスクリーンショットを撮ります。バイト配列を返します
/// </summary>
public byte[] TakeScreenShot()
{
}

/// <summary>
/// デスクトップ全体のスクリーンショットを撮ります。場所に保存し、その場所を返します
/// </summary>
public string TakeScreenShot1(string SavePath)
{
}

/// <summary>
/// URLから受信した画像を含むバイト配列を返します
/// </summary>
/// <param name="URL"></param>
public byte[] PictureFromUrl(string URL, bool ShowError = false, long sender = 0)
{
}
/// <summary>
/// winScpを使用します。ファイルを指定されたホストに安全にアップロードします
/// </summary>
public string FTPS_UPLOAD(string iHost, int iPort, string iUsername, string iPassword, string iLocalFileName, string iRemoteLocation, string iHostCertificateFingerprint = "")
{
}

/// <summary>
/// C#のstring.format()を使用してフォーマットされた文字列を返します
/// </summary>
public string FN_STRING_FORMAT(string iString, params object[] iParams)
{
}


public string FN_SERIALIZE(dynamic iObject)
{
}

/// <summary>
/// 画面に対する相対的なカーソル位置を返します
/// </summary>
public string getCursorPosition()
{
}

/// <summary>
/// 親ウィンドウのダイアログフォームを表示します.. カスタマイズ不可
/// </summary>
/// <param name="iHWND"></param>
public int AreYouSure(int iHWND)
{
}

/// <summary>
/// 確認ダイアログを表示します、カスタマイズ可能
/// </summary>
public int ShowDialog(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}

/// <summary>
/// 確認ダイアログを表示します、カスタマイズ可能
/// </summary>
public int ShowDialogRich(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}

/// <summary>
/// JSON構成を使用してリッチダイアログフォームを表示します
/// </summary>
public int ShowDialogJSON(string JSONConfig)
{
}

/// <summary>
/// インプットボックスフォームを表示します
/// </summary>
public string ShowInputBox(InputBoxType Type = InputBoxType.Text, string Title = "", string Message = "", int PosX=0, int PosY=0, string ThemeBg = "", string ThemeForeColour = "")
{
}

/// <summary>
/// プログレスバーを表示します
/// </summary>
public long OpenProgressBar(string Title, string Message, int Total, bool AutoClose, string ThemeBg, string TitleForeColour)
{
}

/// <summary>
/// 既存のプログレスバーの値を設定するか、エラーを表示します
/// </summary>
public long SetProgressBar(long Handle, int CurrentValue, string Message, int NewMaxValue, bool AutoClose = false)
{
}

/// <summary>
/// すでに開いているプログレスバーを閉じます。
/// </summary>
/// <param name="Handle"></param>
public void CloseProgressBar(long Handle)
{
}

/// <summary>
/// クリップボードに画像が含まれている場合、一時的な場所に保存し、ファイルパスを返します
/// </summary>
public string SaveClipboardToImage(string path, string FileName, string ImageType)
{
}

/// <summary>
/// Webからファイルをダウンロードしてローカルパスに保存します。保存されたファイルパスを返します
/// </summary>
public string DownloadAFile(string url, string destination, bool overWrite, bool ShowProgress)
{
}



public string PadLeft(string Input, int Length, string PaddingChar="")
{
}

public string PadRight(string Input, int Length, string PaddingChar="")
{
}


/// <summary>
/// JSON文字列を動的型に逆シリアル化します。動的オブジェクトを返します
/// </summary>
public object JSONToObject(string json)
{
}

/// <summary>
/// JSON動的オブジェクトからプロパティを読み取り、プロパティ値を返します。
/// </summary>
public string JSONGetValue(object iObject, string propertyName)
{
}

/// <summary>
/// 指定されたJSONオブジェクトからJSONプロパティを抽出し、値をJSONオブジェクトとして返します。
/// </summary>
public object JSONGetObject(object jsonParsedObject, string propertyName)
{
}

/// <summary>
/// VBAユーザー向けのモーダルモダンUIカレンダーを表示します
/// </summary>
/// <returns></returns>
public DateTime ShowCalendar()
{
}

/// <summary>
/// カスタムファイルを開くダイアログを表示します。ドラッグアンドドロップも可能です。
/// </summary>
/// <returns>Json形式の文字列</returns>
public string ShowDialogForFile(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX =0, int PosY =0, string ThemeBg="", string ThemeForeColour="", bool closeAfterFileDrop = true)
{
}

/// <summary>
/// カスタムファイルを開くダイアログを表示します。ドラッグアンドドロップも可能です。
/// </summary>
/// <returns>String[] 配列</returns>
public string[] ShowDialogForFileArray(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX = 0, int PosY = 0, string ThemeBg = "", string ThemeForeColour = "", bool closeAfterFileDrop = true)
{
}

/// <summary>
/// HTMLカラーをアクセスカラーコードに変換します
/// </summary>
public int ColorHexToAccess(string HTMLColor)
{
}

/// <summary>
/// MS ACCESSカラーをHTMLカラーコードに変換します
/// </summary>
public string ColorAccessToHex(long AccessColor)
{
}


/// <summary>
/// URLが到達可能かどうか、trueまたはfalseを返します
/// </summary>
public bool UrlIsReachable(string url)
{
}

/// <summary>
/// URLの形式が正しいかどうか、trueまたはfalseを返します
/// </summary>
public bool UrlIsValid(string url)
{
}
/// <summary>
/// 指定されたURLはローカルファイルパスか？
/// </summary>
public bool UrlIsLocalPath(string p)
{
}

/// <summary>
/// 指定されたURLはローカルファイルパスか？
/// </summary>
public bool UriIsLocalPath(string p)
{
}

// ------------------  Dell specific functions------------------------

/// <summary>
/// アプリバージョンを返します
/// </summary>
public string version()
{
}

public string copyright()
{
}
```
