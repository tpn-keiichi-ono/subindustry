/**
 * 会議資料を作る画面（MeetingRunDialog）の代わり（テスト用）。
 * 資料を作る関数（createBusinessMtgSummaries など）は画面を開くだけなので、テストでは、画面での操作を gas.promptAnswers の答えでまねる：
 *   1つ目の答え：作るシートの番号（空欄なら、画面で初めに選ばれているシート。「1,3」のように複数も。統合版には無い）
 *   2つ目の答え：会議日（「10/23」など。画面のプルダウンで選んだ会議日、または「一覧に無い日を入力」）
 * 比べる時点は gas.runPoint（画面のプルダウンの key。関数なら画面の内容から選ぶ。空なら前回の会議）。
 * 作っている途中の仕事があれば、画面の知らせを gas.alerts に残し、gas.confirmAnswer が YES なら「やめて新しく作る」を選んだとみなす。
 * 「次へ（確かめる）」の文は、本番なら gas.alerts に残し、gas.confirmAnswer が YES のときだけ「作成する」を押す。結果の文も gas.alerts に残す。
 * 本物の関数（画面を開く）は gas.realRun に残す。
 */
'use strict';

const RUNS = {
  createBusinessMtgSummaries: ['business', true],
  previewBusinessMtgSummaries: ['business', false],
  createBusinessMtgCombinedSummary: ['combined', true],
  createServiceMtgSummaries: ['service', true],
  previewServiceMtgSummaries: ['service', false]
};

function withRunDialog(gas) {
  const g = gas.global;
  gas.realRun = {};
  Object.keys(RUNS).forEach(fn => {
    if (typeof g[fn] !== 'function') return;
    gas.realRun[fn] = g[fn];
    const [kind, final] = RUNS[fn];
    g[fn] = () => emulate(gas, kind, final);
  });
  return gas;
}

function emulate(gas, kind, final) {
  const g = gas.global;
  g.mtgUi_();               // スプレッドシートの画面に出せなければ止める（本物と同じ）
  g.mtgRequireSlides_();    // Slides API が無ければ止める（本物と同じ）
  const data = g.getMeetingRunData(kind, final);
  const take = () => String(gas.promptAnswers.length ? gas.promptAnswers.shift() : gas.promptAnswer || '').trim();
  let cancelJob = false;
  if (data.job) {
    gas.alerts.push({title: 'MeetingRunDialog', message: data.job});
    if (gas.confirmAnswer !== 'YES') return;
    cancelJob = true;
  }
  let names = [];
  if (kind !== 'combined') {
    const text = take().normalize('NFKC');
    names = text ? text.split(/[,、・]/).filter(Boolean).map(part => {
      const sheet = data.sheets[Number(part) - 1];
      if (!sheet) throw new Error('「' + part + '」に当たるシートがありません（画面の選択肢は ' + data.sheets.length + '件）。');
      return sheet.name;
    }) : data.sheets.filter(s => s.checked).map(s => s.name);
  }
  const meeting = take();
  const point = typeof gas.runPoint === 'function' ? gas.runPoint(data) : gas.runPoint || '';
  const check = g.checkMeetingRun(kind, final, meeting, names, point);
  if (check.error) {
    gas.alerts.push({title: 'MeetingRunDialog', message: check.error});
    return;
  }
  if (final) {
    gas.alerts.push({title: 'MeetingRunDialog', message: check.text});
    if (gas.confirmAnswer !== 'YES') return;
  }
  const result = g.runMeetingMaterials(kind, final, meeting, names, cancelJob, point);
  gas.alerts.push({title: 'MeetingRunDialog', message: result.message});
}

module.exports = {withRunDialog};
