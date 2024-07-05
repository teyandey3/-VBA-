このファイルは、Excelのマクロを動かすためのVBAコードです。

標準モジュールファイル内のコードはVisual Basicの標準モジュールへ、クラスファイル内のコードはVisual Basicのクラスモジュールへインポートしてください（インポートすると自動で）。

コードの中に大量の日本語を使ってしまったため、このレポジトリからbasファイルもしくはclsファイルをダウンロードしてVisual Basicにインポートしても文字化けを起こしてしまいます。(いつか英語に書き換える予定)。

そのため、今のところGitHub上にあるコードを利用する際は、各モジュールのコードをVisual Basicの標準モジュールもしくはクラスモジュールに直接コピペするしか方法がありません。

VSCodeで開いたときに文字化けする際は以下の手順を試してください。

-----------Files: Auto Guess Encodingの機能を有効化する--------------
1. 「File」メニューから「Preferences」を選択し、「Settings」をクリック。
2. 検索バーに「Auto Guess Encoding」と入力。
3. 「Files: Auto Guess Encoding」の設定を見つけ、チェックボックスをオンにする。

※この機能はすべてのケースで正確に文字コードを推測できるわけではないため、問題が解決しない場合は、手動で文字コードを設定する必要がある。

参考
https://wk-partners.co.jp/homepage/blog/software/how-to-solve-garbled-characters-problem-in-vscode/

https://dianxnao.com/%E6%96%87%E5%AD%97%E5%8C%96%E3%81%91%E5%AF%BE%E7%AD%96%EF%BC%9Avscode%E3%81%A7%E6%96%87%E5%AD%97%E3%82%B3%E3%83%BC%E3%83%89%E3%82%92%E8%87%AA%E5%8B%95%E5%88%A4%E5%88%A5%E3%81%99%E3%82%8B%E8%A8%AD/
