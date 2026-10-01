# CLAUDE.md — このリポジトリで作業するときのルール

Google スプレッドシートにバインドされた Google Apps Script（V8）のプロジェクトです。
全体像と設計の理由は `docs/DESIGN.md`、運用手順は `docs/OPERATIONS.md` を先に読んでください。

## 前提

- 実行環境は Apps Script。リポジトリ直下の `.gs`・`.html` が Apps Script のファイルで、全ファイルが **1つのグローバルスコープ** に読み込まれる。ファイルをまたいで関数・定数を参照できる一方、**同じ名前の関数・`const` が2つあると壊れる**（`const` はエラー、`function` はどちらが動くか不定）。
- 利用者は日本語で操作する。**画面の文言・エラーメッセージ・コードのコメントは日本語**で、既存の言い回し（です・ます調、「〜してください」）に合わせる。
- ユーザーはスプレッドシートの利用者で、Apps Script エディタに貼り付けて反映することもある。変更は「どのファイルを差し替えれば済むか」が分かるようにまとめる。

## 守ること

1. **`.gs` と `.html` に同じ名前を付けない**（Apps Script では拡張子違いでも同名不可）。HTML を増やすときは `XxxDialog` / `XxxView` のように別名にし、`createTemplateFromFile()` の名前と一致させる。
2. **`onOpen` はプロジェクト全体で1つだけ**（現在は `CredentialHistory.gs`）。メニュー項目は `credAddMenu_()` に足す。別のメニューは、メニューを作る関数（例：`コード.gs` の `apAddMenu_()`）を `onOpen` から呼ぶ。
3. **`DIFF_RULES` の `ranges` にボタン列（チェックボックスの列）を含めない**。範囲を変えたら `setupDiffTracking()` の実行が必要（運用側の作業として必ず伝える）。
4. **差分追跡の記録を壊さない**：`変更履歴_差分`・`変更時点スナップショット`・非表示の `__CHAR_DIFF_*` シートの列構成や書き方を変えるときは、既存の読み手（`Changehistory.gs`・`HistorySidebar.gs`）も合わせて直す。記録は追記のみで、過去の行を書き換えない。
5. **シートへの文字列の書き込みは数式にならない方法で**：差分ログはリッチテキスト、クレデンシャル記録は先頭に `'` を付けた値（`credText_()`）。`setValue(ユーザー入力)` をそのまま使わない。
6. **書き込みはドキュメントロックの中で**：`LockService.getDocumentLock()`。差分追跡は `diff*`、クレデンシャルは `credWithLock_()` を使う。読み取りだけの処理（サイドバー・変更履歴の表示）はロックを取らない。
7. **トリガーの所有者は1人**（管理者アカウント）。インストール型トリガーを増やすときは `diffInstallTriggers_()` / `credInstallTrigger_()` のように「1つだけ存在する」ことを保証する関数を通す。ボタン列のチェック検知は `onCredentialLauncherEdit` 1つで全ボタン列をまかなう（トリガーを増やさない）。
8. **UI はデザインの決まりに従う**（`docs/DESIGN.md` の「画面のデザイン」）。色・フォント・余白はモーダル間でそろえる。
9. **`clasp push` の前に必ず `clasp pull` と差分の確認**。push はエディタ側にしかないファイルを消す。ファイル名はエディタ側に合わせてある（`test.gs`・`Changehistory.gs` など）ので、勝手に変えない（`clasp pull` で古い名前も戻り、同じ関数が2つになる）。
10. **Apps Script のファイルはリポジトリ直下に置き、フォルダ（`src/` など）に入れない**。同期するとエディタのファイル名が `src/HistorySidebarView` のようになり、`createTemplateFromFile()` で HTML が見つからなくなる（実際に起きた）。`tests/`・`scripts/` の `.js` は Apps Script に送らない（`.claspignore` で除外している。送ると `require` でプロジェクト全体が動かなくなる）。

## 確認のしかた

- 変更したら `npm run check` と `npm test` を通す。
- ロジックを変えたら `tests/` にテストを足す（`tests/lib/gas-mock.js` のモックを使う）。モックは `SpreadsheetApp` の一部しか再現していないので、使う API が無ければモックに足す。
- 画面（`.html`）は、Playwright などでモックの `google.script.run` を差し込んで表示を確かめられる（`docs/DESIGN.md` の「テスト」）。
- Apps Script 上でしか確かめられないこと（トリガーの起動時間、`google.script.host.setWidth` の効き方、権限）は、ユーザーに確認手順を伝える。

## 命名

- 接頭辞でファイル（機能）を表す：`diff*`（差分追跡）、`cred*`（クレデンシャル）、`chg*`（変更履歴）、`hs*`（サイドバー）、`ap*`（アカウントプラン作成。`コード.gs` の既存の関数は接頭辞なし）、`smp*`（サンプルデータ）、`oc*`（Opportunity Canvas）。
- 末尾 `_` の関数は内部用（メニューや `google.script.run` から直接呼ばない）。
- 設定は各ファイル先頭の `DIFF_OPTIONS` / `CRED_OPTIONS` / `CHG_OPTIONS` / `HS_OPTIONS` に集める。
