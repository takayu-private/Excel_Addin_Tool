## ライセンス
本ツールは非商用目的で利用可能です。商用利用を希望される場合は、事前に作者に連絡し、許可を得る必要があります。詳細は [LICENSE.txt](./LICENSE.txt) をご確認ください。

## 検索拡張ツール使用方法

〇インストール
1. fileCp.ps1を実行
　実行することにより、検索拡張ツールがエクセルアドインの格納場所にコピーされます。
　格納場所：C:\Users\<ユーザ名>\AppData\Roaming\Microsoft\AddIns
　処理終了後、格納場所のショートカットリンクが作成されます。
　コピーに失敗した場合は「copy_ng」が作成されます。
2. エクセルを起動
3. エクセル画面の「開発」タブ→「EXCEL アドイン」を押下
4. アドインのポップアップ画面上の「参照」ボタンを押下→ファイルの参照画面で「SearchExtensionAddin.xlam」を選択し「OK」を押下
　　→アドイン画面に戻り「Searchextensionaddin」にチェックが入っていることを確認し「OK」をクリック
5. 「エクセル検索拡張ツール」タブが表示されることを確認
完了

〇使用方法
・「インストール」で追加した「エクセル検索拡張ツール」タブからメニューを選択
　◆ブック内文字列検索・置換
　　検索・ブック内検索、ブック内置換
　　・使用感は通常の検索画面と同様でテキストボックスに文字をいれ、実行したいボタンをクリックしてください。
　　・検索範囲では「すべて」「セル」「図形」が選択でき、セルのみ、図形のみの検索も可能です。
　　　(デフォルトは「すべて」です。)
　　・正規表現にチェックを入れた場合はキャプチャ置換も可能です。
　◆ファイル検索
　　・使用感は検索画面とサクラエディタのGrep機能を足した感じになってます。
　　・検索場所は「・・・」ボタンからフォルダを選択可能です。(直接入力も可能、パスの最後に「\」は不要)
　　・対象ファイルは「;」区切りで複数設定可能です。ファイル名の一部のみでも大丈夫です。
　　　(「a」だけを設定した場合、ファイル名の一部に「a」が含まれるファイルが対象となる)
　　・検索結果は「【ファイル検索】検索結果」シート、置換結果は「【ファイル検索】置換結果」シートが追加されます。
　　　(既に存在する場合はシートの内容を削除して結果が出力されます。)
　　・検索・置換の仕様は「ブック内検索・置換」と同様です。

○注意事項
　・正規表現のオプションは検索・置換文字列に適用され、ファイル検索・置換の「検索場所」、「対象ファイル」には適用されません。
　　(対象ファイルはチェックの有無にかかわらず、正規表現として認識されます。)
　・置換実行時はエクセルの仕様上、置換前に戻す(Ctlr + z)ができなくなります。
　　「すべて置換」、「ファイル置換」をする場合は事前にバックアップをとることをお勧めします。
　・ファイル検索について
　　「対象ファイル」は正規表現として検索されます。
　　　特殊文字を使用する際は注意してください。(正規表現で使われる記号はエスケープ文字「\」が必要)
　　　※記号が無ければ支障は出ません。
　　　例：
　　　　○：test.xlsx;test.*\.xlsx
　　　　×：*.xlsx ←ワイルドカードとしては使用できない。
　　　また、ファイル名が「~$」始まりのファイルは検索対象外となります。
　　　(「~$」始まりのファイルはエクセルの一時ファイル)
　　検索対象のディレクトリ内にエクセルでサポートされていない拡張子のファイルが存在する場合は想定外の動作をする可能性があります。

○アドインの無効
1. エクセルを起動
2. エクセル画面の「開発」タブ→「EXCEL アドイン」を押下
3. アドインのポップアップ画面上の「Searchextensionaddin」のチェックを外し「OK」をクリック
4. 「エクセル検索拡張ツール」タブが削除されることを確認

○アンインストール
1. 「アドインの無効」の手順を実施
2. 「インストール」で作成したショートカットリンクからアドインの格納先に遷移し「SearchExtensionAddin.xlam」を削除
　ショートカットリンク先：C:\Users\<ユーザ名>\AppData\Roaming\Microsoft\AddIns
3. エクセル画面の「開発」タブ→「EXCEL アドイン」を押下
4. アドインのポップアップ画面上の「Searchextensionaddin」の押下すると確認メッセージが出るため「はい」を押下
　メッセージ内容：「アドイン'<ファイルパス>'が見つかりません。リストから削除しますか？」
5. アドインリストから「Searchextensionaddin」が削除されていることを確認
完了

○アップデート予定機能(リリース時期は未定・しない可能性あり)
既存機能のアップデート
・特になし
新規追加機能(検討中)

○最後に
バグや追加・既存機能のアップデートのアイデアは常に受け付けてます。
ただし、そのバグや機能のアップデートはいつされるかは未定です。
