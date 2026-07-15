# VBA Sprint 公開手順

このフォルダ内のファイルを、すべて同じ階層のままGitHubへアップロードします。

1. GitHubで公開リポジトリを新規作成します（例：`vba-sprint`）。
2. このフォルダの全ファイルをリポジトリ直下へアップロードします。
3. リポジトリの `Settings` → `Pages` を開きます。
4. `Source` で `Deploy from a branch` を選択します。
5. Branchを `main`、フォルダを `/ (root)` にして保存します。
6. 表示された `https://ユーザー名.github.io/vba-sprint/` をスマホで開きます。

## ホーム画面への追加

- iPhone/iPad：Safariの共有メニュー →「ホーム画面に追加」
- Android：Chromeのメニュー →「ホーム画面に追加」または「アプリをインストール」

## 更新時の注意

`index.html`などを更新したときは、`sw.js`先頭のキャッシュ名を `vba-sprint-v2` のように変更すると、端末へ新しいファイルが反映されやすくなります。

学習履歴は各ブラウザの端末内に保存され、PCとスマホ間では同期されません。
