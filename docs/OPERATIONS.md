# 導入・運用の手順

管理者（トリガーを所有するアカウント）向けの手順です。トリガーは**管理者1人だけ**が設置してください。

## 1. はじめて導入する

1. スプレッドシートの「拡張機能 > Apps Script」を開き、`src/` のファイルを同じ名前で追加する。
   - `.gs` は「スクリプト」、`.html` は「HTML」で作る。名前は拡張子なしで入力する（例：`CredentialDialog`）。
   - `.gs` と `.html` に同じ名前は付けられない。
   - `onOpen` はプロジェクト全体で1つだけ。ほかのファイルにもある場合は `CredentialHistory.gs` の `onOpen` を消し、既存の `onOpen` から `credAddMenu_()` を呼ぶ。
2. 追跡シート（`新FMT`・`新FMT2`）の2行目に次の見出しを用意する。
   - `得意先`
   - `最終更新日時`（差分追跡が自動で書き込む列）
   - `クレデンシャル`（改行）`オファリング登録`（ボタン列。改行・空白は無視して照合）
   - `変更履歴`（ボタン列）
3. `DiffTracking.gs` の `DIFF_RULES` の範囲から、2つのボタン列を外す（例：ボタン列が AG・AH なら `A3:AF`）。
4. エディタで次の順に実行する（初回は権限の承認が出る）。
   1. `setupDiffTracking()`：比較基準を作り、差分追跡のトリガー（編集・変更・5分ごと）を設置
   2. `setupCredentialLauncher()`：クレデンシャルのボタン列にチェックボックスを入れ、起動トリガーを設置
   3. `setupChangeHistoryLauncher()`：変更履歴のボタン列にチェックボックスを入れる
   4. （任意）メニュー「（管理者）開いたときにサイドバーを自動表示」
5. スプレッドシートを再読み込みし、メニュー「クレデンシャル」が出ることを確かめる。

## 2. clasp で同期する（Claude Code で作業するとき）

```bash
npm install -g @google/clasp      # または npx clasp ...
clasp login                       # 管理者アカウントで
cp .clasp.json.example .clasp.json
# .clasp.json の scriptId に「プロジェクトの設定 > ID」を入れる（rootDir は src のまま）
```

**最初に必ず `clasp pull` する。** `clasp push` は、エディタ側にしかないファイル（もともとあった別のスクリプトなど）を削除します。

```bash
clasp pull          # エディタの今の内容を src/ に取り込む
git diff            # リポジトリの内容と違うところを確認する
```

- エディタ上で差分追跡のファイル名が `test.gs` の場合、pull すると `src/test.gs` ができる。`DiffTracking.gs` と中身が同じなら、どちらかに名前をそろえる（両方を push すると同じ関数が2つになり壊れる）。
- `appsscript.json`（マニフェスト）も pull で取り込まれる。リポジトリに含めてよい。

変更を反映するとき：

```bash
npm run check && npm test
clasp push
```

push しただけではトリガーは変わりません。`DIFF_RULES` を変えたとき、権限（スコープ）が増えたときは、下の「3. 運用」の該当手順を行ってください。

## 3. 運用

| こんなとき | すること |
|---|---|
| 行を追加・削除した | 何もしなくてよい（自動で記録して追跡を続ける）。追加した行のボタンはメニュー「ボタン列を設定・補充」「変更履歴ボタン列を設定・補充」で入れる |
| 列を追加・削除した | 追跡が止まる。`DIFF_RULES` の範囲を確かめて `setupDiffTracking()` を実行 |
| 並べ替えた・行をドラッグで動かした | 検知されない。`setupDiffTracking()` を実行（比較基準の取り直し） |
| `DIFF_RULES` を変えた | `setupDiffTracking()` を実行 |
| コードの更新で権限が増えた | 管理者がエディタで関数を1つ実行して承認し、`reinstallDiffTriggers()` を実行 |
| 状態を確かめたい | `diagnoseDiffState()`・`diagnoseDiffTriggers()` を実行して実行ログを見る |
| 全員のサイドバー自動表示を止めたい | `removeHistorySidebarAutoOpen()` を実行 |

`setupDiffTracking()` は比較基準を取り直すだけで、記録（`変更履歴_差分` など）は消しません。

## 4. トラブル対応

| 症状 | 原因と対処 |
|---|---|
| `SyntaxError: Identifier 'xxx' has already been declared` | 同じ名前の `const` が2回ある。貼り付けで既存の行と重なった、または同じ中身のファイルが2つある |
| トリガー作成画面で「同じ名前の関数が複数」 | `onOpen` などが2つある。1つにする |
| `No HTML file named XXX was found` | HTML ファイルが無い・名前が違う（`XXX.html.html`、スクリプトとして作った、大文字小文字）・未保存 |
| 「同じ名前は付けられません」 | `.gs` と `.html` を同名にしようとしている。HTML 側を `…View` などにする |
| 「〜列が DIFF_RULES の ranges に含まれています」 | ボタン列が追跡範囲に入っている。範囲から外して `setupDiffTracking()`、そのあとボタンの設定をやり直す |
| 「差分記録を停止しました」 | 列の変更などで追跡が止まった。`diagnoseDiffState()` で原因を見て `setupDiffTracking()` |
| モーダルの中身が「読み込めませんでした」 | 利用者がスクリプトを未承認。メニューから何か1つ実行して承認してもらう |
| チェックしてもモーダルが開かない | 起動トリガー（`onCredentialLauncherEdit`）が無い。`setupCredentialLauncher()` を実行。見出しの文字が設定と違う場合も反応しない |
| モーダルが95%に広がらない | 1回目は既定の大きさで開くことがある。2回目以降は保存した大きさで開く |
