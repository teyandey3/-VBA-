このファイルは、Excelのマクロを動かすためのVBAコードです。

インポートの手順
1. Excelの上タブの「開発」→「Visial Basic」→「(Visial Basic内の上タブ)ファイル」→「ファイルのインポート」
2. ダウンロードしたbasファイルかclsファイルを選択して"開く"ボタンを押す
3. 標準モジュールまたはクラスモジュールにファイルが追加されていれば完了。

コードをコピペする際は、標準モジュールファイル内のコードはVisual Basicの標準モジュールへ、クラスファイル内のコードはVisual Basicのクラスモジュールへコピペしてください。

**GitHub上でコードを編集しないでください。**

コードの中に大量の日本語を使ってしまったため、GitHub上でコードを編集すると、Visual Basicにインポートしたときに文字化けを起こしてしまいます。

コードを編集する際は、**必ずGitHubからダウンロードして**PC上のコードエディターで編集してください

VSCodeで開いたときに文字化けする際は以下の手順を試してください。

-----------Files: Auto Guess Encodingの機能を有効化する--------------
1. 「File」メニューから「Preferences」を選択し、「Settings」をクリック。
2. 検索バーに「Auto Guess Encoding」と入力。
3. 「Files: Auto Guess Encoding」の設定を見つけ、チェックボックスをオンにする。

※この機能はすべてのケースで正確に文字コードを推測できるわけではないため、問題が解決しない場合は、手動で文字コードを設定する必要がある。

参考
https://wk-partners.co.jp/homepage/blog/software/how-to-solve-garbled-characters-problem-in-vscode/

https://dianxnao.com/%E6%96%87%E5%AD%97%E5%8C%96%E3%81%91%E5%AF%BE%E7%AD%96%EF%BC%9Avscode%E3%81%A7%E6%96%87%E5%AD%97%E3%82%B3%E3%83%BC%E3%83%89%E3%82%92%E8%87%AA%E5%8B%95%E5%88%A4%E5%88%A5%E3%81%99%E3%82%8B%E8%A8%AD/
