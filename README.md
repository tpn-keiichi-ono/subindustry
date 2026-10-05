# subindustry

Google スプレッドシート「シナリオ攻略先リスト」に組み込む Google Apps Script（コンテナバインド）です。
次の3つの仕組みで構成されています。

1. **差分追跡**：追跡シート（`新FMT`・`新FMT2`）のセルの変更を、文字単位の差分・編集者・変更時点の行全体のスナップショット付きで記録する。行の追加・削除も自動で記録して追跡を続け、並べ替えなどの行のずれにも気づいて突き合わせ直す。
2. **クレデンシャル・オファリング履歴**：得意先ごとに、クレデンシャル（実績紹介）とオファリング（提案）の予定・実績を登録・編集・削除する。オファリングはクレデンシャルに紐づけられる。
3. **履歴の閲覧**：行ごとのボタン（チェックボックス）から開くモーダル（クレデンシャル・オファリング履歴／変更履歴）と、選んだ行に合わせて表示が切り替わる履歴サイドバー。

ほかに、アカウントプランシートの作成（テンプレートのコピー）と、「案件_」シートの列の保護も同じプロジェクトに入っています。

## フォルダ構成

```
（リポジトリ直下）            Apps Script のファイル。名前はエディタ側と同じ。フォルダに入れないこと
  test.gs                       差分追跡（v6）
  CredentialHistory.gs          クレデンシャル・オファリング履歴のサーバー側、メニュー（onOpen）、ボタン列の起動
  CredentialDialog.html         クレデンシャル・オファリング履歴のモーダル
  Changehistory.gs              変更履歴モーダルのサーバー側
  ChangeHistoryDialog.html      変更履歴のモーダル
  HistorySidebar.gs             履歴サイドバーのサーバー側（メニュー「履歴機能」から開く）、承認用のウェブアプリ
  HistorySidebarView.html       履歴サイドバー
  AuthorizeView.html            承認用のウェブアプリ（HistorySidebar.gs の doGet）のページ
  コード.gs                     アカウントプランシートの作成（メニューは onOpen から追加）
  Loading.html                  アカウントプランシート作成の進み具合の画面
  自動入力部分のシート保護.gs   「案件_」シートの列の保護
  サンプルデータ.gs             画面確認用のサンプル履歴の追加・削除（エディタから実行）
  OpportunityCanvas.gs          オポチュニティキャンバス（行の内容をキャンバスの枠に並べる）のサーバー側
  OpportunityCanvasDialog.html  オポチュニティキャンバスのモーダル
  TaskManagement.gs             タスク管理シート（サービス案ごとの検討の進捗）
  ServiceManagement.gs          サービスリクエスト（新FMT のリクエストを選択パネルで登録する。サービス案を付けてまとめる。メニューは onOpen から追加）
  RequestPickerView.html        サービスリクエストの選択パネル（サイドバー）
  RequestListDialog.html        サービスリクエストの未登録・登録済みの一覧（パネルの件数のバッジから開くモーダル）
  appsscript.json               マニフェスト（タイムゾーン・ランタイム）
docs/
  DESIGN.md                     設計思想・データモデル・処理の流れ・判断の理由・画面のデザイン・テスト
  OPERATIONS.md                 導入・運用・同期・トラブル対応の手順
tests/                        Node.js で動くロジックのテスト（Apps Script には送らない）
  lib/gas-mock.js               Apps Script のモック
  lib/fixture.js                テスト用の追跡シート（新FMT・新FMT2）
samples/新FMT_サンプルデータ.tsv  新FMT に貼り付けられるサンプルの5行（オポチュニティキャンバスの確認用。Apps Script には送らない）
scripts/check-syntax.js       構文・ファイルの置き場所・名前の重複・onOpen の数などのチェック
.claspignore                  clasp で push するファイルを Apps Script のファイルだけに絞る
CLAUDE.md                     Claude Code 向けの作業ルール
```

Apps Script のファイルをフォルダ（`src/` など）に入れると、同期したときにエディタのファイル名が `src/HistorySidebarView` のようになり、
`createTemplateFromFile('HistorySidebarView')` で HTML が見つからなくなります。必ずリポジトリ直下に置いてください（`npm run check` で確かめられます）。

## よく使うコマンド

```bash
npm test          # ロジックのテスト（依存パッケージなし。Node.js 18 以上）
npm run check     # 構文・ファイルの置き場所・名前の重複・onOpen の数などをチェック
npx clasp pull    # Apps Script エディタの最新を取り込む
npx clasp push    # Apps Script に反映する（必ず pull・確認してから）
```

clasp の準備は `docs/OPERATIONS.md` の「clasp で同期する」を参照してください。

## ドキュメント

- 設計の考え方と全体像：[docs/DESIGN.md](docs/DESIGN.md)
- 導入・運用・トラブル対応：[docs/OPERATIONS.md](docs/OPERATIONS.md)
- Claude Code で作業するときのルール：[CLAUDE.md](CLAUDE.md)
