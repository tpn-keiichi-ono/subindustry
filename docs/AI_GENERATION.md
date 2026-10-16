# 会議サマリを AI で作る：GAS・GCP での実現方法（検討案）

会議サマリ（RX Business MTG の事前資料）を、スプレッドシートを読み取って AI に作らせる場合の進め方の検討です。
今の版（`MeetingSummary.gs`）は、集計も文もコードで作っています。この文書は案で、まだ実装していません。

## 結論（推奨）

- **数字はコード、文章は AI**：件数・金額・前回比・区分・並び・停滞などの判定は、今の GAS のコードで決めます。
  AI には、コードが確かめた事実を文章にする部分だけを任せます（活動の要点、長い文の要約、リード文の言い回し）。
- **まず GAS から Vertex AI の Claude を呼ぶ（案A）**：今の仕組み（スプレッドシートにバインドした GAS、試作 → 確認 → 本番の流れ）を変えずに足せます。
- **規模や運用が広がったら Cloud Run に移す（案B）**：実行時間の制限、定期実行、ログ・監査が必要になったときです。
- **AI エージェントに資料を丸ごと作らせる形（案C）は、定例の配布資料には使わない**：単発の分析には向きます。

### 数字を AI に任せない理由

- **同じ数を指すところは全ページで一致していなければならない**：サマリーの「変化のある案件 14件」は区分の行の合計・P4・Appendix ① と一致し、
  期待値の前回比は滝の図の内訳の合計と一致します。これらは変更履歴（`変更履歴_差分`）との照合で決まります。
  AI は数え漏れや計算違いをすることがあり、同じ入力でも毎回少し違う結果を返すので、配布資料の数字の正しさを保証できません。
- **事実だけの方針を守る**：資料には判断・論点を書きません。AI の文には評価の言葉が混ざることがあるので、検査して外す仕組みが要ります（後述）。
- **根拠を説明できる**：数字の根拠を、コードとテスト（`tests/meeting-summary.test.js`）で説明できます。

## AI に任せる候補

| 箇所 | 今 | AI に任せると | 優先度 |
|---|---|---|---|
| 活動の要点（サマリーの「活動更新」、P4・Appendix ① の活動状況） | 最新の活動状況をそのまま出す（長いと「…」で切る） | 前回の活動状況と今の活動状況から、この期間に何をしたかを事実だけで1文に（例：「10/13 常務への説明の場を依頼（前回：提案書の方向性を合意）」） | 高（会議で「活動の内容」を見るため） |
| 長い文の要約（活動状況・活動における課題） | セルに入らない分は「…」で切る | 意味を保ってセルに入る長さに縮める | 中 |
| リード文の言い回し | 決まった型の文（数字を差し込む） | コードが渡した数字だけを使い、ページの状態を事実として書く | 低（今の型で足りている） |
| 記載の矛盾の検出（例：活動状況に「失注」とあるのにステータスが進行中） | なし | 入力内容の確認事項に「記載の不一致（AI による検出）」として出す | 後で（人の確認が前提） |

## 構成案

### 案A：GAS ＋ Vertex AI（Claude）

```
スプレッドシート ─(GAS：mtgSummarize_ で集計)→ 事実の JSON
  → UrlFetchApp → Vertex AI（Claude）→ 文の JSON → 検査（不合格ならテンプレートの文）
  → mtgSlides_ / mtgSlideRequests_ → Google スライド・PDF（今と同じ）
```

**準備**

1. 標準の GCP プロジェクトを用意し、Apps Script の「プロジェクトの設定」で紐付けます（既定のプロジェクトでは Vertex AI を有効にできません）。
2. Vertex AI API を有効にし、Model Garden で使う Claude のモデルを有効にします。必要なら割り当て（クォータ）を申請します。
3. 実行する人（管理者）に、そのプロジェクトの Vertex AI ユーザー（`roles/aiplatform.user`）を付けます。GAS は実行する人の権限で呼びます。
4. `appsscript.json` の `oauthScopes` に `https://www.googleapis.com/auth/cloud-platform` と `https://www.googleapis.com/auth/script.external_request` を足します。
   **`oauthScopes` を書くとスコープの自動検出が止まる**ので、今使っているスコープ（スプレッドシート・ドライブ・スライド・UI など）もすべて書きます。
   書いたあとは、`requireAllScopes()` の流れで全員が1回承認し直します（`docs/OPERATIONS.md` の権限の節も直す）。
5. プロジェクト ID・リージョン・モデルはスクリプトのプロパティに置き、`MTG_OPTIONS.ai`（`enabled` など）で切り替えます。

**呼び出し（GAS ではAnthropic の SDK を使えないので、Vertex AI の REST を直接呼ぶ）**

```js
// 例：活動の要点を JSON で受け取る（実装時に Vertex AI のドキュメントで URL・項目を確かめる）
function mtgAiCall_(system, facts, schema) {
  const props = PropertiesService.getScriptProperties();
  const project = props.getProperty('MTG_AI_PROJECT');
  const region = props.getProperty('MTG_AI_REGION') || 'global';
  const model = props.getProperty('MTG_AI_MODEL') || 'claude-opus-5-5';
  const host = region === 'global' ? 'aiplatform.googleapis.com' : region + '-aiplatform.googleapis.com';
  const url = 'https://' + host + '/v1/projects/' + project + '/locations/' + region + '/publishers/anthropic/models/' + model + ':rawPredict';
  const body = {
    anthropic_version: 'vertex-2023-10-16',
    max_tokens: 16000,
    system,
    messages: [{role: 'user', content: JSON.stringify(facts)}],
    output_config: {effort: 'medium', format: {type: 'json_schema', schema}}
  };
  const res = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', payload: JSON.stringify(body),
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
  if (res.getResponseCode() !== 200) return null;            // 429・5xx は少し待って1回だけやり直し、だめならテンプレートの文
  const msg = JSON.parse(res.getContentText());
  if (msg.stop_reason !== 'end_turn') return null;            // refusal・max_tokens などはテンプレートの文に戻す
  const text = msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
  return JSON.parse(text);
}
```

- 1シートにつき1回の呼び出しにまとめます（案件ごとに呼ばない）。出力は短い JSON にし、`effort` は文の要約なら `low`〜`medium`。
- GAS の1回の実行は6分までで、`UrlFetchApp` にも待ち時間の上限があります。シートが多いときは、シートごとに分けて実行します。
- Claude Opus 5.5 は温度などのサンプリングの設定を受け付けません。同じ入力でも文が少し変わるので、後述の「同じデータなら同じ文」で抑えます。

**長所**：今の仕組み・権限・運用のまま足せる。集計とスライドの作り方は今のコードとテストをそのまま使える。データは自社の GCP の中で扱える。
**短所**：実行時間の制限。`oauthScopes` を明示する手間。AI 呼び出しのログ・監査は GAS のログ（Cloud Logging）だけ。

### 案B：Cloud Run（＋ Cloud Scheduler）

```
Cloud Scheduler（または GAS のメニューから呼ぶ）→ Cloud Run（Node.js / TypeScript）
  → Sheets API で読む → 集計（MeetingSummary.gs のロジックを移す）→ Vertex AI（@anthropic-ai/vertex-sdk の AnthropicVertex）
  → 検査 → Slides / Drive API → 共有ドライブに保存
```

- 集計のロジックは今の `.gs` が JavaScript なので、モジュールに分けて移せます。テスト（`tests/`）もそのまま使えます。
- 認証はサービスアカウントです。スプレッドシートと保存先のフォルダをサービスアカウントに共有します（またはドメイン全体の委任）。
  Vertex AI はサービスアカウントの権限（ADC）で呼びます。
- **長所**：実行時間の制限が無い。リトライ・ログ・監査（Cloud Logging）、CI/CD、秘密情報の管理、複数シートの並列処理ができる。
- **短所**：インフラと運用が要る。差分の記録（GAS のトリガー）とは別の系統になり、両方を保守する。
- **移る目安**：シートが増えて6分に収まらない、AI の処理が増える、定期の自動作成や監査ログが必要になったとき。

### 案C：AI エージェントにシートを読ませて資料を作らせる

- Claude にツール（シートを読む・スライドを作る）を渡して、作り方も任せる形です。
  Vertex AI ではマネージドのエージェント（Managed Agents）やコード実行は使えないので、Cloud Run 上でツール実行のループを自分で持ちます。
- **定例の配布資料には向かない理由**：毎回の構成・数字がぶれる、数字の正しさを確かめにくい、所要時間・費用が読みにくい。
- **向く用途**：単発の分析（「この案件の前回からの動きを調べて」「停滞の理由を活動状況から整理して」）。会議サマリとは別の機能として考えます。

## 品質を守る仕組み（どの案でも）

1. **入力は集計済みの事実だけ**：案件のキー・得意先・案件名・前回と今の活動状況など、必要な項目だけを JSON で渡します（メールアドレスなどは渡さない）。
2. **出力は JSON スキーマで固定**：例 `{"items": [{"dealKey": "...", "summary": "..."}]}`（`output_config.format`）。
3. **検査**（不合格ならもう1回作り直し、それでも不合格ならテンプレートの文＝今のコードの文に戻す）：
   - 文の中の数字・日付が、入力にある数字・日付だけか（正規表現で抜き出して照合）
   - 判断・評価の言葉が無いか（「要対応」「懸念」「好調」「推奨」「すべき」「リスク」など。`docs/DESIGN.md` の使わない言葉と合わせる）
   - 空欄を推定で埋めていないか：入力で空の項目は「未入力」と書かせ（指示に明記する）、入力に無い値（金額・日付・ステータス・活動の中身）が文にあれば不合格
     （`docs/DESIGN.md` の「空欄と推定」と同じ決まり）
   - 長さ（全角の文字数）、書き言葉（体言止め・である調）、`dealKey` が入力にある案件か
4. **同じデータなら同じ文**：データのハッシュをキーにして、試作で作った文を保存し、本番でデータが変わっていなければその文を使います（本番で文が変わらない）。
5. **試作 → 人の確認 → 本番**の流れは今のままです。AI の文は試作で確かめます。
6. **AI が作った文であることの表示**：表紙の出典の行に「文の一部は AI で作成」などを出すかは、社内のルールに合わせて決めます。
7. **評価してから使う**：過去の会議日のデータで作った文を、検査の合格率と人の評価（正確さ・事実だけか・読みやすさ）で確かめてから、本番に入れます。

## モデル・費用・制限

- **モデル**：Claude Opus 5.5（`claude-opus-5-5`。Vertex AI でも同じ ID）。費用や速度を優先する場合は Claude Sonnet 5.5（`claude-sonnet-5-5`）・
  Claude Haiku 5.5（`claude-haiku-5-5`）も選べます（品質を評価してから決める）。
- **費用の目安**：1シート1回あたり、入力 1〜2万トークン・出力 2千トークン程度の想定です。Anthropic の直販価格（Opus 5.5：入力 $4・出力 $20／100万トークン）で
  1回 0.1ドル前後です。Vertex AI の価格は Google の価格表に従います。週1回 × シート数なら、月に数ドルの規模です。
- **Vertex AI で使えないもの**：Message Batches・Files API・コード実行・Managed Agents（今回の用途では不要）。
- **拒否（`stop_reason: "refusal"`）**：テンプレートの文に戻します（SDK の自動の切り替えは GAS からは使えないため）。
- **リージョン**：`global`（推奨）か特定のリージョン。データの所在について社内のルールを確かめます（日本のリージョンでの提供状況は Model Garden で確認）。

## 進め方

0. 今の版（AI なし）をリリースします。
1. 案A で「活動の要点」だけを AI にします。試作でだけ使い、テンプレートの文と並べて比べます（`MTG_OPTIONS.ai.enabled` で切り替え）。
2. 評価がよければ本番でも使い、長い文の要約・リード文に広げます。
3. 必要になったら案B に移します。

## 決めてもらうこと

- 使う AI と社内の利用ルール（Vertex AI の Claude でよいか、データの扱い、AI が作った文の表示の要否）
- GCP プロジェクト（課金と権限の持ち主）、リージョン
- どこから AI に任せるか（上の候補の順でよいか）
