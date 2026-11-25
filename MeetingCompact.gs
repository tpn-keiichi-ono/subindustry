/**
 * 統合版（報告用）：全サブインダストリーの資料を、数字の要点と変化のあった案件だけにまとめた資料（RX Business Meeting）。2026/11/09 から。
 * 統合版（createBusinessMtgCombinedSummary()。目次・全サブインダストリーの状況・計画・各サブインダストリーの資料のページ）とは別の資料で、統合版は変えない。
 *
 * ページの並び：
 *   1. 表紙（統合版と同じひな形の表紙。見出しの下は「事前資料」ではなく「報告資料」（compact.coverWord）。mtgCombinedCover_）
 *   2. 全サブインダストリーの状況・計画進捗（mtgCompactOverviewSlide_）：サブインダストリーごとに、今年度の確定（確度 A＋B）の想定売上の横棒（A・B を積み上げ、
 *      データラベル付き）と計画線（透明度のあるグレーの点線。上に「計画」）。右に、計画・進捗率（A＋B）・確定（A＋B）/件・見込（C）/件と、
 *      想定売上合計（失注除く）・前回比（比べた日付。増は青・減は赤）・案件数・変化のある案件（まとまりの見出しは付けない）。シートの順（DIFF_RULES）、いちばん下に合計
 *   2'. 同じ（見込（C）を含む。棒に続けて見込（C）を点線で足す。mtgCompactOverviewSlide_(…, {withC: true})）
 *   2''. 状況・計画進捗（FY26〜28売上）（mtgCompactOutlookSlide_。2026/11/12）：FY26売上の確度 A〜D の横棒と、FY26・FY27 の確定・見込、FY26〜28 の A〜D の合計
 *   3. 計画に対する積み上げ（全体。統合版と同じ mtgPlanSlide_ で、想定売上を積み上げる）と、確定（A＋B）だけの積み上げ（mtgPlanRanksOnly_）
 *   （変化のあった案件の前）RXシナリオ一覧（シナリオNo）（mtgCompactScenarioSlides_。全体サマリだけのときは入れない）
 *   4〜 変化のあった案件：{サブインダストリー}（mtgCompactChangeSlides_。シートの順にサブインダストリーごと）：No（資料の通し番号）・得意先／案件名（／シナリオNo）・
 *      ステータス・想定売上・受注月（前回▶今回）・活動状況・活動における課題。変化が無ければ「該当なし」の1ページ
 * 集計はどれも期待値ではなく想定売上（summary.planAmount = mtgPlanOf_(…, 'amount')・bridge.amountStart / amountEnd。2026/11/09）。
 * 金額は小数点以下を繰り上げた整数で出す（compact.ceilAmounts。mtgWithAmountCeil_）
 *
 * 作り方は統合版と同じ：エディタで createBusinessMtgReportingSummary() を実行し、画面（MeetingRunDialog）で会議日を選ぶ。すべての追跡シートの会議サマリを
 * 今の値で作り（読むだけ。会議日・時点コピーは記録しない）、本番として保存して会議のフォルダに格納する。インデックス（INDEX_会議資料）には書かない
 * （統合版の列は統合版のもの）。各サブインダストリーの本番の資料ができていなくても作れる。時間がかかれば区切って続ける（統合版と同じトリガー
 * continueBusinessMtgSummaries。仕事の種類 job.kind = 'compact'。mtgDeckJob_）
 */

/** 統合版（報告用）を作る（エディタから実行）。画面で会議日（と比べる時点）を選ぶ */
function createBusinessMtgReportingSummary() {
  mschOpenRunDialog_('compact', true);
}

/**
 * 統合版（報告用）の全体サマリだけを作る（エディタから実行。数字の確かめ用）：表紙・全サブインダストリーの状況・計画進捗・計画に対する積み上げのページだけで、
 * 変化のあった案件のページは作らない（そのぶん早く終わる）。画面は統合版（報告用）と同じ（final: false で区別する）。
 * 保存先のフォルダに「…_報告用_全体サマリ.pdf」として保存し、会議のフォルダには格納しない
 */
function createBusinessMtgReportingOverview() {
  mschOpenRunDialog_('compact', false);
}

/**
 * 統合版（報告用）の変化のあった案件だけを作る（エディタから実行。2026/11/12）：表紙と「変化のあった案件：{サブインダストリー}」のページだけで、
 * 状況・計画進捗と計画に対する積み上げのページは作らない（全体サマリ createBusinessMtgReportingOverview() と対になる）。画面は統合版（報告用）と同じ。
 * 保存先のフォルダに「…_報告用_変化のあった案件.pdf」として保存し、会議のフォルダには格納しない
 */
function createBusinessMtgReportingChanges() {
  mschOpenRunDialog_('compact', 'changes');
}

/**
 * 統合版（報告用）のどこまでを作るか（画面から届く final から）：true なら ''（すべて）・'changes' なら 'changes'（変化のあった案件だけ）・
 * それ以外（false）なら 'overview'（全体サマリだけ）
 */
function mtgCompactPart_(final) {
  return final === true ? '' : final === 'changes' ? 'changes' : 'overview';
}

/** 仕事（job）の作る範囲（mtgCompactPart_ の値。前の版で保存した仕事は overview で区別していた） */
function mtgCompactJobPart_(job) {
  return job.part || (job.overview ? 'overview' : '');
}

/**
 * 「統合版（報告用）（全サブインダストリー）」（知らせの名前）。part が 'overview' なら「統合版（報告用）の全体サマリ（全サブインダストリー）」、
 * 'changes' なら「統合版（報告用）の変化のあった案件（全サブインダストリー）」
 */
function mtgCompactLabel_(part) {
  const word = part === 'overview' ? 'の全体サマリ' : part === 'changes' ? 'の変化のあった案件' : '';
  return MTG_OPTIONS.compact.label + word + '（' + MTG_OPTIONS.combined.name + '）';
}

/**
 * 変化のあった案件のページの前に入るページの数（表紙・全サブインダストリーの状況・計画進捗（見込（C）を含むページ・FY26〜28 のページも）
 * （・計画に対する積み上げ 2枚））。変化のあった案件だけ（part 'changes'）なら表紙の1枚
 */
function mtgCompactFrontCount_(part) {
  if (part === 'changes') return 1;
  const C = MTG_OPTIONS.compact;
  const kinds = (C.salesPage ? 1 : 0) + (mtgCompactAmountOn_() ? 1 : 0);
  // 売上のページを FY26〜28 のページにする（outlookPage）ときは、状況・計画進捗が2枚（①FY26の内訳・②FY26〜28。2026/11/25 から。前は1枚）で、
  // 計画に対する積み上げは FY26だけの2枚と FY26〜28 の2枚
  const outlook = C.salesPage && C.outlookPage ? (MTG_OPTIONS.plan.enabled ? 2 : 0) + mtgCompactSpanPatterns_().length : 0;
  return 1 + kinds * (MTG_OPTIONS.plan.enabled ? 4 : 2) + outlook;
}

/** 想定売上で集計したページを入れるか（compact.amountPages。今年度の売上のページを入れないときは、想定売上のページだけにする） */
function mtgCompactAmountOn_() {
  const C = MTG_OPTIONS.compact;
  return !!C.amountPages || !C.salesPage;
}

/** fn を、金額の小数点以下を繰り上げる設定（MTG_OPTIONS.amountCeil。mtgAmount_）で作る（compact.ceilAmounts のとき。終わったら元に戻す） */
function mtgWithAmountCeil_(fn) {
  const o = MTG_OPTIONS;
  if (!o.compact.ceilAmounts) return fn();
  const before = o.amountCeil;
  o.amountCeil = true;
  try {
    return fn();
  } finally {
    if (before === undefined) delete o.amountCeil;
    else o.amountCeil = before;
  }
}

/**
 * ファイル名（統合版の名前の後ろに fileSuffix。「20261113_全サブインダストリー_報告用」。全体サマリだけなら後ろに overviewSuffix、
 * 変化のあった案件だけなら changesSuffix）
 */
function mtgCompactFileName_(meeting, part) {
  const C = MTG_OPTIONS.compact;
  const tail = part === 'overview' ? C.overviewSuffix : part === 'changes' ? C.changesSuffix : '';
  return mtgFileName_(MTG_OPTIONS.combined.name, meeting, true) + C.fileSuffix + (tail || '');
}

/**
 * 統合版（報告用）のスライドの並び（一度に作るとき。テストと見た目の確認に使う。実際の資料は mtgCompactRun_ で区切りながら同じ並びに作る）。
 * summaries はシートの順。ページ番号は表紙を1枚目とした通し番号
 */
function mtgCompactSlides_(summaries, meeting, final) {
  const now = summaries.length ? summaries[0].now : new Date();
  // 変化のあった案件のページの前に RXシナリオ一覧
  const scenario = mtgNumberSlides_(mtgCombinedFrame_(summaries, meeting, final), mtgCompactScenarioSlides_(summaries), mtgCompactFrontCount_());
  let page = mtgCompactFrontCount_() + scenario.length;
  let no = 0;
  const pages = [];
  summaries.forEach(s => {
    const part = mtgCompactChangeSlides_(s, no);
    no += part.count;
    pages.push.apply(pages, mtgNumberSlides_(mtgCombinedFrame_([s], meeting, final), part.slides, page));
    page += part.slides.length;
  });
  return mtgCompactFront_(summaries, meeting, final, now).concat(scenario, pages);
}

/**
 * 前のページ（表紙・全サブインダストリーの状況・計画進捗・計画に対する積み上げ（全体））。ページ番号は1から。
 * summaries は会議サマリの中身か、その一部（mtgCombinedLite_。区切って作るとき）。now は中身を作った時刻
 */
function mtgCompactFront_(summaries, meeting, final, now, part) {
  const o = MTG_OPTIONS;
  const C = o.compact;
  // 変化のあった案件だけ（createBusinessMtgReportingChanges()）は表紙だけ
  if (part === 'changes') {
    return mtgNumberSlides_(mtgCombinedFrame_(summaries, meeting, final), [mtgCombinedCover_(summaries, meeting, final, now, C.coverWord)]);
  }
  return mtgWithAmountCeil_(() => {
    const amountOn = mtgCompactAmountOn_();
    const merged = o.plan.enabled ? mtgMergePlans_(summaries.map(mtgCompactPlanOf_)) : null;
    // 計画に対する積み上げ：今年度の売上（全確度・確定（A＋B））のあとに、想定売上（同じ2枚。amountPages のとき）
    const plan = [];
    // タイトルは括弧書きを付けない（2026/11/10）：「計画に対する積み上げ」「計画に対する積み上げ：確定（A＋B）」。今年度の売上と想定売上の両方を
    // 入れるとき（amountPages）は、想定売上のページに「：想定売上」を付けて見分ける
    const T = MTG_TITLES.plan;
    const amountWord = C.salesPage ? C.measure : '';
    const titles = {sales: T, salesSure: T + '：' + C.sureLabel,
      amount: T + (amountWord ? '：' + amountWord : ''), amountSure: T + '：' + [amountWord, C.sureLabel].filter(Boolean).join('・')};
    if (merged && C.salesPage) {
      // 売上の列が無いシートは除いて足す（どのシートにも無ければ「データなし」）
      const pick = list => (list.some(x => !x.gap) ? list.filter(x => !x.gap) : list.slice(0, 1));
      // 案件の今年度の売上の合計を、受注月の月に置いて積み上げたもの（FY26〜28 のページにも使う）
      const byMonth = mtgMergePlans_(pick(summaries.map(s => s.planSalesByMonth || mtgSalesPlanByMonthOf_([], s.name, now, {}))));
      byMonth.year = byMonth.year || mtgFiscalYearKey_(now);
      const fy = 'FY' + byMonth.year;
      const measure = fy + '売上';
      // FY26〜28 のページを入れる（outlookPage）ときは、FY26だけのページのタイトルに「（FY26売上）」を付けて見分ける（2026/11/12）
      const salesTitles = C.outlookPage ? {sales: T + '（' + measure + '）', salesSure: T + '（' + measure + '）：' + C.sureLabel} : titles;
      if (C.salesPlanBy === 'half') {
        const sales = mtgMergeSalesPlans_(pick(summaries.map(s => s.planSales || mtgSalesPlanOf_([], s.name, now, {}))));
        plan.push(mtgSalesPlanSlide_(sales, o.combined.name, {title: salesTitles.sales}), mtgSalesPlanSlide_(mtgPlanRanksOnly_(sales, C.sureRanks), o.combined.name,
          {only: C.sureRanks, label: C.sureLabel, title: salesTitles.salesSure}));
      } else {
        // 案件の今年度の売上の合計を、受注月の月に置いて積み上げる（想定売上の積み上げと同じ形）
        // リード文は期間を書かない（図の見出しにある。折り返さない長さにするため。shortLead）
        plan.push(mtgPlanSlide_(byMonth, now, o.combined.name, {measure, titleNote: measure, title: salesTitles.sales, shortLead: true}),
          mtgPlanSlide_(mtgPlanRanksOnly_(byMonth, C.sureRanks), now, o.combined.name, {only: C.sureRanks, label: C.sureLabel, measure, titleNote: measure,
            title: salesTitles.salesSure, shortLead: true}));
      }
      // FY26〜28 のページ（outlookPage）：FY26だけの2枚のあとに、今年度の四半期と来年度以降の半期を続けて積み上げる2枚（A〜D・確定（A＋B））
      if (C.outlookPage) {
        plan.push(mtgCompactPlanYearsSlide_(byMonth, summaries, now), mtgCompactPlanYearsSlide_(byMonth, summaries, now, {only: C.sureRanks, label: C.sureLabel}));
      }
    }
    if (merged && amountOn) {
      plan.push(mtgPlanSlide_(merged, now, o.combined.name, {measure: C.measure, title: titles.amount}),
        mtgPlanSlide_(mtgPlanRanksOnly_(merged, C.sureRanks), now, o.combined.name, {only: C.sureRanks, label: C.sureLabel, measure: C.measure,
          title: titles.amountSure}));
    }
    // 状況・計画進捗：今年度の売上（と見込（C）を含む）のあとに、想定売上（同じ2枚。amountPages のとき）
    // 売上のページ：FY26〜28 のページ（outlookPage。2026/11/12 から）か、前の今年度の売上の2枚（FY26売上・見込（C）を含む）
    // FY26〜28 の版は、①FY26の内訳（確度 A〜D と計画）と②FY26〜28（FY26 は合算・FY27・FY28・総額）の2枚（2026/11/25 に1枚から分けた）
    const status = [].concat(!C.salesPage ? [] : C.outlookPage ? [mtgCompactOutlookSlide_(summaries), mtgCompactOutlookSpanSlide_(summaries)]
      .concat(mtgCompactSpanPatterns_().map(pt => mtgCompactOutlookSpanSlide_(summaries, pt)))
      : [mtgCompactOverviewSlide_(summaries, {sales: true}), mtgCompactOverviewSlide_(summaries, {sales: true, withC: true})],
      amountOn ? [mtgCompactOverviewSlide_(summaries), mtgCompactOverviewSlide_(summaries, {withC: true})] : []);
    return mtgNumberSlides_(mtgCombinedFrame_(summaries, meeting, final),
      [mtgCombinedCover_(summaries, meeting, final, now, C.coverWord)].concat(status, plan));
  });
}

/**
 * 全サブインダストリーの状況・計画進捗（FY26売上）（①。2026/11/12 の FY26〜28 の1枚を、2026/11/25 に①FY26の内訳と②FY26〜28 に分けた）。
 * 横棒は今年度の売上（受注月が今年度の案件）を確度 A〜D（compact.outlookRanks）で積み上げ（棒の中に値。右に合計）、計画の位置に点線（compact.planLine）。
 * 右の列：計画・進捗率（A＋B）・確定（A＋B）・見込（C）・見込（D）（FY26）。FY27 以降は②（mtgCompactOutlookSpanSlide_）。
 * 件数は出さない。並びはシートの順（ほかのページと同じ）、いちばん下に合計。中身は summary.salesOutlook
 * （mtgSalesOutlookOf_）。売上・受注月の列が無いシートは、棒の場所に「データなし」と理由（合計はそのシートを除いて ※）
 */
function mtgCompactOutlookSlide_(summaries) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const G = o.gapText;
  const C = o.compact;
  const now = summaries.length && summaries[0].now ? new Date(summaries[0].now) : new Date();
  const keys = C.outlookRanks || ['A', 'B', 'C', 'D'];
  const sure = C.sureRanks || ['A', 'B'];
  const outlookOf = s => s.salesOutlook || mtgSalesOutlookOf_([], s.name, now, {});
  const items = summaries.map(s => ({name: s.name, p: outlookOf(s)}));
  const ok = items.filter(z => !z.p.gap);
  const year = (ok[0] || items[0] || {p: {}}).p.year || mtgFiscalYearKey_(now);
  const next = ('0' + ((Number(year) + 1) % 100)).slice(-2);
  const last = ('0' + ((Number(year) + 2) % 100)).slice(-2);
  // 合計（作れたシートの分だけ）
  const total = {name: '合計', total: true, p: {plan: 0, all: 0, byYear: {[year]: {}}, later: {[next]: 0, [last]: 0}, gap: ''}};
  ok.forEach(z => {
    total.p.plan += z.p.plan || 0;
    total.p.all += z.p.all || 0;
    keys.forEach(k => { total.p.byYear[year][k] = (total.p.byYear[year][k] || 0) + ((z.p.byYear[year] || {})[k] || 0); });
    [next, last].forEach(y => { total.p.later[y] += (z.p.later || {})[y] || 0; });
  });
  const laterOf = (p, y) => (p.later || {})[y] || 0;
  const sumOf = (p, y, list) => list.reduce((t, k) => t + (((p.byYear || {})[y] || {})[k] || 0), 0);
  const rateOf = p => (p.plan > 0 ? sumOf(p, year, sure) / p.plan * 100 : null);
  const fy = y => 'FY' + y;
  const sureWord = C.sureLabel || '確定（A＋B）';
  const rows = items.concat([total]);
  const marked = ok.length < items.length;
  const totalRate = ok.length ? rateOf(total.p) : null;
  const range = keys[0] + '〜' + keys[keys.length - 1];
  // リード文は2行（金額が5桁でも折り返さない長さ。2026/11/16）：1行目に今年度、2行目に棒の説明と次のページ
  const lead = fy(year) + ' の' + sureWord + 'は ' + mtgAmount_(sumOf(total.p, year, sure)) + '百万円（計画比 ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') +
    '）、' + range + ' の合計は ' + mtgAmount_(sumOf(total.p, year, keys)) + '百万円（失注除く）。\n' +
    '棒は ' + fy(year) + '売上（' + mtgPlanMonthName_() + 'が ' + fy(year) + 'の案件）を ' + range + ' で積み上げ、点線は計画。' + fy(next) + '以降は次のページ。';
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 480, 20, fy(year) + ' の売上の確度別（' + keys[0] + '〜' + keys[keys.length - 1] + '）と計画（百万円）',
    {size: 10.5, bold: true, color: c.ink})];
  const nameW = 96;
  // 列：今年度（計画・進捗率・確定（A＋B）・確定のほかの確度ごとの見込（C・D。outlookRanks に創注を入れれば見込（創注）も））。
  // 棒を広くとる（来年度以降の列は②のページ）
  const prospects = keys.filter(k => sure.indexOf(k) < 0);
  const cols = [{key: 'plan', label: '計画', period: '（' + fy(year) + '）', w: 48}, {key: 'rate', label: '進捗率（A＋B）', period: '（' + fy(year) + '）', w: 62},
    {key: 'ab0', label: sureWord, period: '（' + fy(year) + '）', w: 62}]
    .concat(prospects.map(k => ({key: 'r_' + k, rank: k, label: '見込（' + k + '）', period: '（' + fy(year) + '）', w: k.length > 1 ? 62 : 56})));
  // いちばん右に今年度の合計（棒の長さ。A〜D。2026/11/25）
  const sumPeriod = '（' + fy(year) + '・' + range + '）';
  cols.push({key: 'sum', label: '合計', period: sumPeriod, w: Math.max(62, Math.ceil(mtgTextWidth_(sumPeriod, 6.5)) + 2)});
  const colsW = cols.reduce((t, col) => t + col.w, 0);
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  const headY = y + 36;
  const labelH = 12;
  const periodH = 10;
  const headH = labelH + periodH;
  const rowH = Math.min(22, (S.bottom - headY - headH - 2) / rows.length);
  // 見出し・期間の字は折り返さない大きさ（状況・計画進捗のページと同じ決め方）
  let size = 7.5;
  while (cols.some(col => mtgTextWidth_(col.label, size) > col.w * 0.9) && size > 6) size -= 0.5;
  let periodSize = 6.5;
  while (cols.some(col => mtgTextWidth_(col.period, periodSize) > col.w) && periodSize > 5.5) periodSize -= 0.5;
  let cx = barX + barW + 10;
  cols.forEach(col => {
    el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, labelH, col.label, {size, color: c.sub, align: 'CENTER', valign: 'BOTTOM', role: 'compact-head'}));
    el.push(mtgBox_(cx - S.padX / 2, headY + labelH, col.w + S.padX, periodH, col.period,
      {size: periodSize, color: c.muted, align: 'CENTER', valign: 'MIDDLE', role: 'compact-head-period'}));
    col.x = cx;
    cx += col.w;
  });
  // 凡例（棒のすぐ上。状況・計画進捗のページと同じ大きさ）
  const ranks = keys.map(k => o.plan.ranks.find(r => r.key === k)).filter(Boolean);
  let lx = barX;
  ranks.forEach(r => {
    const ly = headY + headH - periodH / 2 - 4;
    el.push({type: 'box', x: lx, y: ly, w: 8, h: 8, fill: r.color, shape: 'RECTANGLE', role: 'legend'});
    const bw = mtgTextWidth_(r.key, 8.5) + S.padX + 2;
    el.push(mtgBox_(lx + 13 - S.padX / 2, ly - 4, bw, 16, r.key, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 13 + bw - S.padX + 10;
  });
  // 目盛り：いちばん大きい計画の planScale 倍か、いちばん長い棒を、棒の幅いっぱいに
  const max = Math.max.apply(null, [1].concat(ok.map(z => Math.max((z.p.plan || 0) * C.planScale, sumOf(z.p, year, keys)))));
  const scale = barW / max;
  const line = C.planLine;
  const light = {C: true, D: true, '創注': true};   // 薄い色の確度は、棒の中の値を濃い緑の字に
  rows.forEach((item, i) => {
    const ry = headY + headH + 2 + rowH * i;
    const p = item.p;
    if (item.total) el.push({type: 'box', x, y: ry, w, h: 0.75, fill: c.line, shape: 'RECTANGLE'});
    else if (i % 2 === 1) el.push({type: 'box', x, y: ry, w, h: rowH, fill: c.greenSoft, alpha: 0.35, shape: 'RECTANGLE', role: 'compact-band'});
    el.push(mtgBox_(x, ry, nameW, rowH, item.name, {size: 8.5, bold: true, color: c.ink, role: 'compact-name'}));
    if (!item.total && !p.gap) {
      const bh = Math.max(rowH - 6, 6);
      const by = ry + (rowH - bh) / 2;
      let level = 0;
      ranks.forEach(r => {
        const v = p.byYear[year][r.key] || 0;
        if (!v) return;
        const bx = barX + level * scale;
        const bw = Math.max(v * scale, 0.75);
        el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: r.color, line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role: 'compact-bar'});
        const label = mtgAmount_(v);
        let pt = 7;
        while (mtgTextWidth_(label, pt) + 3 > bw && pt > 6) pt -= 0.5;
        if (mtgTextWidth_(label, pt) + 3 <= bw) {
          el.push(mtgBox_(bx - S.padX / 2, ry, bw + S.padX, rowH, label, {size: pt, bold: true, color: light[r.key] ? c.greenDark : c.white, align: 'CENTER',
            role: 'compact-bar-label'}));
        }
        level += v;
      });
      // 棒の右に A〜D の合計
      if (level) {
        const text = mtgAmount_(level);
        el.push(mtgBox_(barX + level * scale + 2 - S.padX / 2, ry, mtgTextWidth_(text, 7.5) + S.padX + 2, rowH, text,
          {size: 7.5, bold: true, color: c.sub, role: 'compact-bar-total'}));
      }
      if (p.plan > 0) {
        el.push({type: 'line', x: barX + p.plan * scale, y: ry + 1, w: 0, h: rowH - 2, color: line.color, alpha: line.alpha, weight: line.weight, dash: true,
          role: 'compact-plan'});
      }
    } else if (!item.total) {
      el.push(mtgBox_(barX, ry, barW, rowH, mtgWrap_(G + '（' + p.gap + '）', barW, 7.5, 1)[0], {size: 7.5, color: c.muted, role: 'compact-gap'}));
    }
    const none = p.gap || (item.total && !ok.length);
    const mark = item.total && marked ? '※' : '';
    const amount = v => (none ? G : mtgAmount_(v) + mark);
    const rate = none ? null : rateOf(p);
    const values = {plan: amount(p.plan || 0), rate: rate == null ? '—' : Math.round(rate) + '%', ab0: amount(sumOf(p, year, sure)),
      sum: amount(sumOf(p, year, keys))};
    prospects.forEach(k => { values['r_' + k] = amount(sumOf(p, year, [k])); });
    cols.forEach(col => {
      const text = values[col.key];
      const gap = text === G;
      const strong = col.key === 'rate' || col.key === 'ab0' || col.key === 'sum';
      el.push(mtgBox_(col.x - S.padX / 2, ry, col.w + S.padX, rowH, text,
        {size: gap ? 7.5 : col.key === 'rate' ? 9 : 8, bold: !gap && (!!item.total || strong),
          color: gap ? c.muted : col.key === 'rate' && rate != null && rate >= 100 ? c.greenDark : c.ink, align: 'END', role: 'compact-value'}));
    });
  });
  return Object.assign({title: MTG_TITLES.compactOverview + '（' + fy(year) + '売上）', lead, elements: el},
    marked ? {note: '※ のある合計は「' + G + '」のシートを除く'} : {});
}

/**
 * 状況・計画進捗（FY26〜28売上）（②）のあとに足すパターンのページの options（compact.outlookSpanPatterns の順）：
 * 'withSouchu' → FY26 に創注まで入れる {keys, titleNote: '創注を含む'}（outlookRanks に創注が入っていれば作らない）、'merged' → FY26 を確度で分けない
 * {merged, titleNote: 'FY26合算'}。titleNote はタイトルの括弧の中に足す（「（FY26〜28売上・創注を含む）」。1行に入る長さ）
 */
function mtgCompactSpanPatterns_() {
  const o = MTG_OPTIONS;
  const C = o.compact;
  const keys = C.outlookRanks || ['A', 'B', 'C', 'D'];
  const souchu = o.plan.ranks.map(r => r.key).filter(k => keys.indexOf(k) < 0);
  return (C.outlookSpanPatterns || []).map(name => {
    if (name === 'withSouchu') return souchu.length ? {keys: keys.concat(souchu), titleNote: souchu.join('・') + 'を含む'} : null;
    if (name === 'merged') return {merged: true, titleNote: 'FY' + mtgFiscalYearKey_(new Date()) + '合算'};
    if (name === 'fy27Sure') {
      const y = Number(mtgFiscalYearKey_(new Date()));
      return {laterSure: true, oneYear: true, titleNote: 'FY' + ('0' + ((y + 1) % 100)).slice(-2) + 'は' + (C.sureRanks || ['A', 'B']).join('＋')};
    }
    if (name === 'laterSure') {
      const y = Number(mtgFiscalYearKey_(new Date()));
      return {laterSure: true, titleNote: 'FY' + ('0' + ((y + 1) % 100)).slice(-2) + '〜' + ('0' + ((y + 2) % 100)).slice(-2) + 'は' + (C.sureRanks || ['A', 'B']).join('＋')};
    }
    return null;
  }).filter(Boolean);
}

/**
 * 全サブインダストリーの状況・計画進捗（FY26〜28売上）（②。2026/11/25）：横棒は今年度の売上（確度 A〜D の内訳。①と同じ色）に、
 * FY27〜28 の売上の合計（受注月・確度によらない。薄いグレー）を続けた棒（右に総額）。右の列：売上（FY26。A〜D の合算）・売上（FY27）・売上（FY28）｜
 * FY27〜28 の合計｜全体合計（FY26〜28。棒の長さと同じ。状況・計画進捗の全体合計）。計画は今年度だけの目標なので入れない（①のページ）。
 * 並びはシートの順、いちばん下に合計。中身は summary.salesOutlook。列が無いシートは、棒の場所に「データなし」と理由（合計はそのシートを除いて ※）
 */
function mtgCompactOutlookSpanSlide_(summaries, options) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const G = o.gapText;
  const C = o.compact;
  const now = summaries.length && summaries[0].now ? new Date(summaries[0].now) : new Date();
  // options（パターンのページ。mtgCompactSpanPatterns_）：keys は FY26 に入れる確度（創注まで入れる）、merged は FY26 を確度で分けない（合算の1色）、
  // titleNote はタイトルの後ろ（「：創注を含む」）
  const opt = options || {};
  const keys = opt.keys || C.outlookRanks || ['A', 'B', 'C', 'D'];
  const merged = !!opt.merged;
  // laterSure：FY27〜28 は確定（A＋B）の案件の売上だけ（salesOutlook.laterSure。前の版の中身に無ければ 0）
  const laterSure = !!opt.laterSure;
  // oneYear：来年度（FY27）だけを続ける（FY28 は入れない。列も FY26・FY27・全体合計）
  const oneYear = !!opt.oneYear;
  const sure = C.sureRanks || ['A', 'B'];
  const outlookOf = s => s.salesOutlook || mtgSalesOutlookOf_([], s.name, now, {});
  const items = summaries.map(s => ({name: s.name, p: outlookOf(s)}));
  const ok = items.filter(z => !z.p.gap);
  const year = (ok[0] || items[0] || {p: {}}).p.year || mtgFiscalYearKey_(now);
  const next = ('0' + ((Number(year) + 1) % 100)).slice(-2);
  const last = ('0' + ((Number(year) + 2) % 100)).slice(-2);
  const endY = oneYear ? next : last;   // 全体合計の最後の年度
  const fy = y => 'FY' + y;
  const range = keys[0] + '〜' + keys[keys.length - 1];
  const thisOf = p => keys.reduce((t, k) => t + (((p.byYear || {})[year] || {})[k] || 0), 0);
  const laterOf = (p, y) => ((laterSure ? p.laterSure : p.later) || {})[y] || 0;
  const rowOf = (name, p, total) => ({name, total, p, gap: p.gap, plan: p.plan || 0, v0: thisOf(p), v1: laterOf(p, next), v2: oneYear ? 0 : laterOf(p, last)});
  const rows = items.map(z => rowOf(z.name, z.p, false));
  const sum = {name: '合計', total: true, gap: '', v0: 0, v1: 0, v2: 0};
  rows.filter(r => !r.gap).forEach(r => { sum.v0 += r.v0; sum.v1 += r.v1; sum.v2 += r.v2; });
  rows.push(sum);
  rows.forEach(r => { r.later = r.v1 + r.v2; r.all = r.v0 + r.later; });
  const marked = ok.length < items.length;
  const laterWord = fy(next) + (oneYear ? '' : '〜' + last) + (laterSure ? '（' + sure.join('＋') + '）' : '');
  const lead = fy(year) + '〜' + endY + 'の売上の全体合計は ' + mtgAmount_(sum.all) + '百万円（失注除く）。' + fy(year) + ' ' + mtgAmount_(sum.v0) + '（' + range +
    '）・' + laterWord + ' ' + mtgAmount_(sum.later) + '。\n' + '棒は ' + fy(year) + ' の ' + range + (merged ? '（合算）' : '') + ' に ' + laterWord + (oneYear ? ' の売上' : ' の合計') + '（グレー。' + mtgPlanMonthName_() +
    (laterSure ? 'によらない' : '・確度によらない') + '）を続けた総額。点線は ' + fy(year) + 'の計画。';
  // 棒の色：今年度は確度 A〜D の色で内訳を残し（①と同じ）、FY27〜28 は薄いグレー（compact.laterBarColor。受注月・確度によらない売上なので確度の色と分ける）
  const light = {C: true, D: true, '創注': true};   // 薄い色の確度は、棒の中の値を濃い緑の字に
  // merged なら今年度は合算の1色（compact.thisYearColor）
  const segs = (merged ? [{key: 'v0', label: fy(year) + '（' + range + '）', color: C.thisYearColor || '#70AD47', text: c.white}]
    : keys.map(k => o.plan.ranks.find(r => r.key === k)).filter(Boolean)
      .map(r => ({key: r.key, label: r.key, color: r.color, text: light[r.key] ? c.greenDark : c.white})))
    .concat([{key: 'later', label: laterWord, color: C.laterBarColor || '#D9D9D9', text: c.sub}]);
  const segValue = (r, key) => (key === 'later' ? r.later : key === 'v0' ? r.v0
    : ((r.p || {}).byYear || {})[year] ? (r.p.byYear[year][key] || 0) : 0);
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 480, 20, fy(year) + '〜' + endY + 'の売上（百万円）', {size: 10.5, bold: true, color: c.ink})];
  const nameW = 96;
  // FY26 の列は、期間（「（FY26・A〜創注）」など）が 6.5pt で1行に入る幅
  const v0Period = '（' + fy(year) + '・' + range + '）';
  // 期間は 6.5pt で1行に入る幅（確定（A＋B）の案件だけのときも期間は同じ。タイトル・凡例・リード文に「A＋B」と書く）
  const fit = (text, min) => Math.max(min, Math.ceil(mtgTextWidth_(text, 6.5)) + 2);
  const p1 = '（' + fy(next) + '）';
  const p2 = '（' + fy(last) + '）';
  const pL = '（' + fy(next) + '〜' + last + '）';
  const cols = [{key: 'v0', label: '売上', period: v0Period, w: fit(v0Period, 62)},
    {key: 'v1', label: '売上', period: p1, w: fit(p1, 52)}]
    .concat(oneYear ? [] : [{key: 'v2', label: '売上', period: p2, w: fit(p2, 52)}, {key: 'later', label: '合計', period: pL, w: fit(pL, 58)}])
    .concat([{key: 'all', label: '全体合計', period: '（' + fy(year) + '〜' + endY + '）', w: 62}]);
  const colsW = cols.reduce((t, col) => t + col.w, 0);
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  const headY = y + 36;
  const labelH = 12;
  const periodH = 10;
  const headH = labelH + periodH;
  const rowH = Math.min(22, (S.bottom - headY - headH - 2) / rows.length);
  let size = 7.5;
  while (cols.some(col => mtgTextWidth_(col.label, size) > col.w * 0.9) && size > 6) size -= 0.5;
  let periodSize = 6.5;
  while (cols.some(col => mtgTextWidth_(col.period, periodSize) > col.w) && periodSize > 5.5) periodSize -= 0.5;
  let cx = barX + barW + 10;
  cols.forEach(col => {
    el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, labelH, col.label, {size, color: c.sub, align: 'CENTER', valign: 'BOTTOM', role: 'compact-head'}));
    el.push(mtgBox_(cx - S.padX / 2, headY + labelH, col.w + S.padX, periodH, col.period,
      {size: periodSize, color: c.muted, align: 'CENTER', valign: 'MIDDLE', role: 'compact-head-period'}));
    col.x = cx;
    cx += col.w;
  });
  // 凡例（棒のすぐ上）：A〜D（FY26）・FY27〜28
  let lx = barX;
  segs.forEach(sg => {
    const ly = headY + headH - periodH / 2 - 4;
    el.push({type: 'box', x: lx, y: ly, w: 8, h: 8, fill: sg.color, shape: 'RECTANGLE', role: 'legend'});
    const bw = mtgTextWidth_(sg.label, 8.5) + S.padX + 2;
    el.push(mtgBox_(lx + 13 - S.padX / 2, ly - 4, bw, 16, sg.label, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 13 + bw - S.padX + 10;
  });
  // 目盛り：いちばん長い棒（総額）か計画を、右の総額の数の分をあけて棒の幅いっぱいに
  const max = Math.max.apply(null, [1].concat(rows.filter(r => !r.total && !r.gap).map(r => Math.max(r.all, r.plan))));
  const scale = (barW - 34) / max;
  const line = C.planLine;
  rows.forEach((r, i) => {
    const ry = headY + headH + 2 + rowH * i;
    if (r.total) el.push({type: 'box', x, y: ry, w, h: 0.75, fill: c.line, shape: 'RECTANGLE'});
    else if (i % 2 === 1) el.push({type: 'box', x, y: ry, w, h: rowH, fill: c.greenSoft, alpha: 0.35, shape: 'RECTANGLE', role: 'compact-band'});
    el.push(mtgBox_(x, ry, nameW, rowH, r.name, {size: 8.5, bold: true, color: c.ink, role: 'compact-name'}));
    if (!r.total && !r.gap) {
      const bh = Math.max(rowH - 6, 6);
      const by = ry + (rowH - bh) / 2;
      // 今年度の計画の位置に点線（①と同じ。今年度の分・総額と比べやすいように。2026/11/25）。棒の下に描く（棒の中の数に線が重ならないように）
      const planX = r.plan > 0 ? barX + r.plan * scale : null;
      if (planX != null) {
        el.push({type: 'line', x: planX, y: ry + 1, w: 0, h: rowH - 2, color: line.color, alpha: line.alpha, weight: line.weight, dash: true,
          role: 'compact-plan'});
      }
      let level = 0;
      segs.forEach(sg => {
        const v = segValue(r, sg.key);
        if (!v) return;
        const bx = barX + level * scale;
        const bw = Math.max(v * scale, 0.75);
        el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: sg.color, line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role: 'compact-bar'});
        const label = mtgAmount_(v);
        let pt = 7;
        while (mtgTextWidth_(label, pt) + 3 > bw && pt > 6) pt -= 0.5;
        if (mtgTextWidth_(label, pt) + 3 <= bw) {
          el.push(mtgBox_(bx - S.padX / 2, ry, bw + S.padX, rowH, label, {size: pt, bold: true, color: sg.text, align: 'CENTER', role: 'compact-bar-label'}));
        }
        level += v;
      });
      // 棒の右に総額。計画の点線にかかるときは、点線の右に出す（数の上に線が重ならないように）
      if (level) {
        const text = mtgAmount_(level);
        const tw = mtgTextWidth_(text, 7.5);
        let tx = barX + level * scale + 2;
        if (planX != null && tx - 1 <= planX && planX <= tx + tw + 2) tx = planX + 3;
        el.push(mtgBox_(tx - S.padX / 2, ry, tw + S.padX + 2, rowH, text, {size: 7.5, bold: true, color: c.sub, role: 'compact-bar-total'}));
      }
    } else if (!r.total) {
      el.push(mtgBox_(barX, ry, barW, rowH, mtgWrap_(G + '（' + r.gap + '）', barW, 7.5, 1)[0], {size: 7.5, color: c.muted, role: 'compact-gap'}));
    }
    const none = r.gap || (r.total && !ok.length);
    const mark = r.total && marked ? '※' : '';
    cols.forEach(col => {
      const text = none ? G : mtgAmount_(r[col.key]) + mark;
      const gap = text === G;
      const strong = col.key === 'later' || col.key === 'all';
      el.push(mtgBox_(col.x - S.padX / 2, ry, col.w + S.padX, rowH, text,
        {size: gap ? 7.5 : 8, bold: !gap && (!!r.total || strong), color: gap ? c.muted : c.ink, align: 'END', role: 'compact-value'}));
    });
  });
  return Object.assign({title: MTG_TITLES.compactOverview + '（' + fy(year) + '〜' + endY + '売上' + (opt.titleNote ? '・' + opt.titleNote : '') + '）', lead,
    elements: el}, marked ? {note: '※ のある合計は「' + G + '」のシートを除く'} : {});
}

/**
 * 1行の文字が、表のセル（幅 width）に折り返さずに入るか。文字の幅は見積もり（mtgTextWidth_）なので、本文は compact.bodyFit、
 * 太字の見出し（bold）は compact.headFit の余裕をとる（Google スライドの字は見積もりより広く、ぎりぎりだと折り返す。2026/11/16）
 */
function mtgCompactLineFits_(text, size, width, bold) {
  const C = MTG_OPTIONS.compact;
  const ratio = bold ? C.headFit || 0.85 : C.bodyFit || 0.92;
  return String(text).split('\n').every(line => mtgTextWidth_(line, size) <= (Number(width) - MTG_SLIDE.cellPadX) * ratio);
}

/** 1行の文字を、折り返さずに入れるのに要るセルの幅（mtgCompactLineFits_ の逆） */
function mtgCompactLineWidth_(lines, size, bold) {
  const C = MTG_OPTIONS.compact;
  const ratio = bold ? C.headFit || 0.85 : C.bodyFit || 0.92;
  const widest = Math.max.apply(null, [0].concat(lines.map(line => mtgTextWidth_(line, size))));
  return Math.ceil(widest / ratio + MTG_SLIDE.cellPadX);
}

/**
 * 変化のあった案件の表の列の幅。specs は活動状況・課題より前の列（No・得意先／案件名・ステータス・金額・受注月）：{min, size, texts, arrow}。
 * 見出し（labels。太字。改行で行を分ける）と中身（texts。「前 ▶ 今」は arrow の側で2行に分けた、長いほうの行）が折り返さずに入る幅
 * （min より狭くはしない）。残りを活動状況・課題に 55:45 で分ける（2026/11/16）
 */
function mtgCompactChangeColumns_(labels, specs) {
  const S = MTG_SLIDE;
  const fixed = specs.map((spec, i) => {
    const size = spec.size || S.bodySize;
    const head = mtgCompactLineWidth_(String(labels[i]).split('\n'), S.bodySize, true);
    const parts = [];
    (spec.texts || []).forEach(t => {
      const text = String(t == null ? '' : t);
      if (!spec.arrow || text.indexOf(' ▶ ') < 0) {
        parts.push(text);
        return;
      }
      const [from, to] = text.split(' ▶ ');
      parts.push.apply(parts, spec.arrow === 'before' ? [from, '▶ ' + to] : [from + ' ▶', to]);
    });
    return Math.max(spec.min || 0, head, mtgCompactLineWidth_(parts, size, false));
  });
  const rest = S.w - fixed.reduce((a, b) => a + b, 0);
  const activity = Math.round(rest * 0.55);
  return mtgTableWidths_(fixed.concat([activity, rest - activity]));
}

/**
 * 変化のあった案件の表の見出しの行（mtgHeadRow_ と同じ形）。見出しの各行が太字で折り返さずに入るように、入らなければ字を小さくする（7pt まで）
 */
function mtgCompactHeadRow_(columns, labels) {
  const S = MTG_SLIDE;
  const c = MTG_OPTIONS.colors;
  return mtgRow_(columns, labels.map((text, i) => {
    let size = S.bodySize;
    while (!mtgCompactLineFits_(text, size, columns[i], true) && size > 7) size -= 0.5;
    return {text, size, bold: true, color: c.white, fill: c.green, align: 'CENTER', lines: 3};
  }), S.headRow);
}

/**
 * 変化のあった案件のページに載る案件（その他の項目のみの更新を除く）で選ばれているシナリオの数 {シナリオNo: 件数}（Appendix の RXシナリオ一覧）
 */
function mtgCompactScenarioUsed_(s) {
  const out = {};
  (s.changes || []).filter(ch => ch.category !== 'updated').forEach(ch => {
    (ch.scenarios || []).map(mtgScenarioNumber_).filter((no, i, list) => list.indexOf(no) === i).forEach(no => { out[no] = (out[no] || 0) + 1; });
  });
  return out;
}

/** プルダウンの選択肢「01-04.顧客情報活用パッケージ」を {no: '01-04', name: '顧客情報活用パッケージ'} に（番号が無ければ no は空） */
function mtgCompactScenarioItem_(text) {
  const t = String(text || '').normalize('NFKC').trim();
  const m = /^(\d+(?:[-‐−ー]\d+)*)[.．、\s]*(.*)$/.exec(t);
  return m ? {no: m[1].replace(/[‐−ー]/g, '-'), name: m[2].trim() || t} : {no: '', name: t};
}

/**
 * 区切って作るときに、作り始めに各シートの RXシナリオ構成のプルダウンの選択肢だけを読む（見出しの行と、見出しの下の1セルの入力規則）。
 * RXシナリオ一覧は変化のあった案件のページの前に入れるが、変化のあった案件のページはシートを読むたびにページ番号を付けて足すので、
 * 先に一覧のページの数を決めておくため（一覧の形は選択肢だけで決まる。数は仕上げで、読んだシートの中身から入れる）。
 * [{name, scenarioColumn, scenarioOptions}]（読めないシートは列なし）
 */
function mtgCompactScenarioOptions_(ss, names) {
  return names.map(name => {
    try {
      const sheet = ss.getSheetByName(name);
      const rule = DIFF_RULES[name];
      if (!sheet || !rule) return {name, scenarioColumn: false, scenarioOptions: []};
      const headers = sheet.getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
      const col = mtgColumnIndex_(headers, 'scenario');
      if (col < 0) return {name, scenarioColumn: false, scenarioOptions: []};
      return {name, scenarioColumn: true, scenarioOptions: mtgListOptions_(sheet.getRange(rule.headerRow + 1, col + 1).getDataValidation()) || []};
    } catch (error) {
      console.log('RXシナリオ一覧：「' + name + '」のプルダウンを読めませんでした（' + (error && error.message || error) + '）');
      return {name, scenarioColumn: false, scenarioOptions: []};
    }
  });
}

/**
 * 「RXシナリオ一覧（シナリオNo）」（2026/11/16。変化のあった案件のページの前）：変化のあった案件のページの「シナリオNo」と、シナリオの名前の一覧。
 * シナリオは各シートの RXシナリオ構成のプルダウンの選択肢（各シートには、そのサブインダストリーのシナリオだけがある）で、シナリオNo の上2桁が
 * サブインダストリー。サブインダストリーごとに見出しの帯（「01　スーパー・GMS」。上2桁の順）と、No・名前・件数（この資料の変化のあった案件のうち、
 * そのシナリオを選んだ案件の数。選ばれているものは太字、無ければ「—」）。4列に流し込み（サブインダストリーは列の途中で分けない）、入らなければ次のページ（「（続き）」）。
 * どのシートにも選択肢が無ければページを作らない（[]）。summaries は中身か、その一部（mtgCombinedLite_）。
 * options（[{name, scenarioColumn, scenarioOptions}]。mtgCompactScenarioOptions_）があれば、一覧の形はそれで決める（作り始めに数えたページの数と合わせるため）
 */
function mtgCompactScenarioSlides_(summaries, options) {
  const o = MTG_OPTIONS;
  if (!o.compact.scenarioPage) return [];
  const sources = options || summaries;
  const groups = sources.filter(s => s.scenarioColumn).map((s, i) => {
    const items = (s.scenarioOptions || []).map(mtgCompactScenarioItem_);
    // 上2桁（いちばん多いもの）。無ければシートの順
    const counts = {};
    items.forEach(it => { const pre = it.no.slice(0, 2); if (/^\d{2}$/.test(pre)) counts[pre] = (counts[pre] || 0) + 1; });
    const pre = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || '';
    return {pre, name: s.name, items, order: i};
  });
  if (!groups.some(gr => gr.items.length)) return [];
  groups.sort((a, b) => (a.pre && b.pre ? a.pre.localeCompare(b.pre) : a.pre ? -1 : b.pre ? 1 : 0) || (a.order - b.order));
  const used = {};
  summaries.forEach(s => {
    // 区切って作るときの中身の一部（mtgCombinedLite_）には数を入れてある。中身そのものなら変化のある案件から数える
    const counts = s.scenarioUsed || (s.changes ? mtgCompactScenarioUsed_(s) : {});
    Object.keys(counts).forEach(no => { used[no] = (used[no] || 0) + counts[no]; });
  });
  const total = groups.reduce((t, gr) => t + gr.items.length, 0);
  const usedCount = Object.keys(used).length;
  return (() => {
    const S = MTG_SLIDE;
    const c = o.colors;
    const P = S.padX;
    const nc = 4;
    const gap = 10;
    const cw = (S.w - gap * (nc - 1)) / nc;
    const headH = 15;
    const rowH = 13;
    const noW = 38;
    const cntW = 26;
    const top = S.top - 4;
    const bottom = S.bottom;
    const pages = [];
    let el = null;
    let col = nc;
    let y = top;
    const nextCol = () => {
      col++;
      y = top;
      if (col >= nc) {
        el = [];
        pages.push(el);
        col = 0;
      }
    };
    const header = (gr, more) => {
      const x0 = S.x + col * (cw + gap);
      el.push({type: 'box', x: x0, y, w: cw, h: headH, fill: c.green, shape: 'RECTANGLE', role: 'scenario-group'});
      el.push(mtgBox_(x0 + 2 - P / 2, y, cw + P, headH, mtgWrap_((gr.pre ? gr.pre + '　' : '') + gr.name + (more ? '（続き）' : ''), cw - 4 + P, 8, 1)[0],
        {size: 8, bold: true, color: c.white, role: 'scenario-group'}));
      y += headH;
    };
    nextCol();
    groups.forEach(gr => {
      const rows = Math.max(gr.items.length, 1);
      // サブインダストリーは列の途中で分けない（1列に入らないほど多ければ、列の終わりで分けて「（続き）」）
      if (y > top && y + headH + rowH * Math.min(rows, Math.floor((bottom - top - headH) / rowH)) > bottom) nextCol();
      header(gr);
      if (!gr.items.length) {
        el.push(mtgBox_(S.x + col * (cw + gap) + 2 - P / 2, y, cw + P, rowH, 'プルダウンの選択肢なし', {size: 7, color: c.muted, role: 'scenario-none'}));
        y += rowH + 5;
        return;
      }
      gr.items.forEach(it => {
        if (y + rowH > bottom) {
          nextCol();
          header(gr, true);
        }
        const x0 = S.x + col * (cw + gap);
        const n = used[it.no] || 0;
        el.push({type: 'box', x: x0, y, w: noW, h: rowH, fill: c.greenSoft, shape: 'RECTANGLE'});
        el.push(mtgBox_(x0 - P / 2, y, noW + P, rowH, it.no || '—', {size: 7.5, bold: true, color: c.ink, align: 'CENTER', role: 'scenario-no'}));
        // 名前は1行。入らなければ 6pt まで小さくし、それでも入らなければ「…」で切る
        const nameW = cw - noW - 4 - cntW;
        let size = 7.5;
        while (mtgWrap_(it.name, nameW + P, size, 2).length > 1 && size > 6) size -= 0.5;
        el.push(mtgBox_(x0 + noW + 4 - P / 2, y, nameW + P, rowH, mtgWrap_(it.name, nameW + P, size, 1)[0],
          {size, bold: !!n, color: n ? c.ink : c.sub, role: 'scenario-name'}));
        el.push(mtgBox_(x0 + cw - cntW - P / 2, y, cntW + P, rowH, n ? n + '件' : '—', {size: 7, color: n ? c.ink : c.muted, align: 'END', role: 'scenario-count'}));
        el.push({type: 'line', x: x0, y: y + rowH, w: cw, h: 0, color: c.line, weight: 0.5});
        y += rowH;
      });
      y += 5;
    });
    const title = 'RXシナリオ一覧（シナリオNo）';
    const lead = 'シナリオNo の上2桁はサブインダストリー（RXシナリオ構成の選択肢 ' + total + '件）。\n' +
      '右の件数は、次のページからの変化のあった案件のうち、そのシナリオを選んだ案件の数（' + usedCount + 'シナリオ。太字）。';
    return pages.map((elements, n) => ({title: title + (pages.length > 1 ? '（' + (n + 1) + '/' + pages.length + 'ページ）' : ''), lead, elements}));
  })();
}

/**
 * 変化のあった案件の並びに使う、今回のステータスの期待値（%）。ステータスの名前の「(70%)」から読む（「A-受注済み(100%)」は 100）。
 * 名前に % が無ければ確度の順（plan.ranks の順に 90・70・50・30・10）の目安。失注・保留・削除・確度の無いステータスは -1（最後）
 */
function mtgCompactStatusRate_(ch) {
  const group = mtgStatusGroupOf_(ch);
  if (group === 'closed') return -1;
  const m = /\((\d+(?:\.\d+)?)\s*%\)/.exec(String(ch.statusTo || '').normalize('NFKC'));
  if (m) return Number(m[1]);
  const i = MTG_OPTIONS.plan.ranks.findIndex(r => r.key === group);
  // 確度の順に 90・70・50・30・10（A〜創注のふつうの期待値）。同じ確度で % のあるステータスより少し後ろ
  return Math.max(0, (MTG_OPTIONS.plan.ranks.length - i) * 20 - 10) - 0.5;
}

/**
 * 計画に対する積み上げ（FY26〜28売上）（2026/11/12）：左に今年度の売上を受注月の四半期（Q1〜Q4）ごとに確度別に積み上げ（前の四半期の上に続ける。
 * 上に「+増えた分」）、今年度の合計・計画（白抜き）・不足（計画との差。グレーの点線の枠。計画を超えたら「超過」）。区切り線の右に、来年度・再来年度の
 * 半期ごとの売上（受注月・確度によらない。年度ごとに通期で1色（compact.laterColors）。累計の棒には前半期・後半期の境目に点線。summary.salesOutlook.halves。確定（A＋B）のページも同じ値）を、今年度の合計の上に
 * 続けて積み上げ、いちばん右に3年度分の累計。
 * sales は今年度の売上を受注月の月に置いた積み上げ（mtgSalesPlanByMonthOf_ を足したもの）。options.only（['A', 'B']）なら確定（A＋B）だけのページ。
 * 受注月が空・今年度の外の案件の今年度の売上は入れない（note に金額）。失注は除く
 */
function mtgCompactPlanYearsSlide_(sales, summaries, now, options) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const C = o.compact;
  const only = options && options.only;
  const label = (options && options.label) || C.sureLabel;
  // 今年度の分の確度：確定（A＋B）のページは only、ほかは状況・計画進捗（FY26〜28売上）と同じ A〜D（compact.outlookRanks。創注・未設定は除く。
  // 累計を状況・計画進捗の全体合計とそろえるため。2026/11/12）
  const outlookKeys = C.outlookRanks || ['A', 'B', 'C', 'D'];
  const keys = o.plan.ranks.map(r => r.key).concat([o.plan.otherLabel]).filter(k => (only || outlookKeys).indexOf(k) >= 0);
  const ranks = o.plan.ranks.concat([{key: o.plan.otherLabel, color: o.plan.otherColor}]).filter(r => keys.indexOf(r.key) >= 0);
  const p = mtgPlanRanksOnly_(sales, keys);
  const rangeWord = keys[0] + '〜' + keys[keys.length - 1];
  // 除いた確度（「創注・未設定」）
  const dropped = o.plan.ranks.map(r => r.key).concat([o.plan.otherLabel]).filter(k => keys.indexOf(k) < 0);
  const year = Number(p.year || mtgFiscalYearKey_(now));
  const yy = k => ('0' + ((year + k) % 100)).slice(-2);
  const pick = values => keys.reduce((m, k) => (values && values[k] ? (m[k] = values[k], m) : m), {});
  const totalOf = values => keys.reduce((t, k) => t + ((values || {})[k] || 0), 0);
  const sumValues = list => list.reduce((m, v) => { Object.keys(v).forEach(k => { m[k] = (m[k] || 0) + v[k]; }); return m; }, {});
  // 今年度：受注月の四半期
  const quarters = [0, 1, 2, 3].map(q => {
    const values = pick(sumValues(p.months.slice(q * 3, q * 3 + 3).map(mo => mo.values)));
    const a = mtgMonthOf_(p.start + q * 3);
    const b = mtgMonthOf_(p.start + q * 3 + 2);
    return {label: 'Q' + (q + 1), sub: String(a.y).slice(2) + '/' + a.m + '〜' + b.m, values, total: totalOf(values)};
  });
  const thisValues = sumValues(quarters.map(q => q.values));
  const thisTotal = totalOf(thisValues);
  // 来年度・再来年度：半期ごとの売上（受注月・確度によらない。確定（A＋B）のページも同じ値。売上の列が無い・受注月の列が無いシートは除く）。
  // 確度で分けず、年度ごとに通期で1色（compact.laterColors）。凡例は「FY27」「FY28」。累計の棒では前半期・後半期の境目に点線
  const colors = C.laterColors || ['#5B9BD5', '#966BC7'];
  const outlooks = summaries.map(sm => sm.salesOutlook).filter(x => x && !x.gap && x.halves);
  const later = [yy(1), yy(2)].reduce((list, y, k) => list.concat(['前', '後'].map(h => {
    const key = 'FY' + y;
    const total = outlooks.reduce((t, x) => t + ((x.halves.find(v => v.year === y && v.half === h) || {}).value || 0), 0);
    const color = Array.isArray(colors[k]) ? colors[k][0] : colors[k] || '#5B9BD5';
    return {label: 'FY' + y, sub: h + '半期', year: y, half: h, key, color, values: total ? {[key]: total} : {}, total};
  })), []);
  const laterRanks = later.filter(l => l.half === '前').map(l => ({key: l.key, color: l.color}));
  const yearTotal = y => later.filter(l => l.year === y).reduce((t, l) => t + l.total, 0);
  const grand = thisTotal + later.reduce((t, l) => t + l.total, 0);
  const plan = p.plan || 0;
  const fy = 'FY' + yy(0);
  const span = fy + '〜' + yy(2);
  const word = only ? label + 'の' : '';
  const x = S.x;
  const w = S.w;
  const top = S.top;
  const el = [mtgBox_(x, top - 2, 400, 20, span + 'の' + (only ? label + 'の' : '') + '売上の積み上げ（百万円）', {size: 10.5, bold: true, color: c.ink})];
  // 凡例（右上）：確度・来年度以降・計画（白抜き）
  const legend = ranks.map(r => ({key: r.key, color: r.color})).concat(laterRanks, [{key: '計画', outline: true}]);
  const legendW = legend.reduce((t, r) => t + 12 + mtgTextWidth_(r.key, 8) + 2 + 8, 0);
  let lx = x + w - legendW;
  legend.forEach(r => {
    el.push(r.outline ? {type: 'box', x: lx, y: top + 4, w: 8, h: 8, fill: c.white, line: c.ink, lineWeight: 0.75, shape: 'RECTANGLE', role: 'legend'}
      : {type: 'box', x: lx, y: top + 4, w: 8, h: 8, fill: r.color, shape: 'RECTANGLE', role: 'legend'});
    const bw = mtgTextWidth_(r.key, 8) + S.padX + 2;
    el.push(mtgBox_(lx + 12 - S.padX / 2, top, bw, 16, r.key, {size: 8, color: c.sub, role: 'legend'}));
    lx += 12 + bw - S.padX + 8;
  });
  const chartTop = top + 46;
  const base = S.bottom - 34;
  const max = Math.max(plan, grand, 1) * 1.04;
  const k = (base - chartTop) / max;
  // 列：Q1〜Q4・今年度の合計・計画・不足｜来年度の前・後・再来年度の前・後・累計（区切りの分 0.6 列あける）
  const slot = (w - 16) / (12 + 0.6);
  const bw = slot * 0.62;
  const xOf = i => x + slot * i + (slot - bw) / 2 + (i >= 7 ? slot * 0.6 : 0);
  const stack = (bx, from, values) => {
    let level = from;
    ranks.concat(laterRanks).forEach(r => {
      const v = values[r.key] || 0;
      if (!v) return;
      el.push({type: 'box', x: bx, y: base - (level + v) * k, w: bw, h: Math.max(v * k, 0.5), fill: r.color, line: c.white, lineWeight: 0.4, shape: 'RECTANGLE',
        role: 'plan-years-bar'});
      level += v;
    });
    return level;
  };
  const valueLabel = (bx, yTop, text, color) => el.push(mtgBox_(bx - 14, yTop - 13, bw + 28, 12, text,
    {size: 7.5, bold: true, color: color || c.ink, align: 'CENTER', valign: 'BOTTOM', role: 'plan-years-value'}));
  const axis = (bx, text, sub) => el.push(mtgBox_(bx - 14, base + 2, bw + 28, 22, text + '\n' + sub,
    {size: 7.5, bold: true, color: c.ink, align: 'CENTER', valign: 'TOP', role: 'plan-years-axis',
      spans: [{start: text.length + 1, end: text.length + 1 + sub.length, size: 6, bold: false, color: c.muted}]}));
  let level = 0;
  quarters.forEach((q, i) => {
    const bx = xOf(i);
    level = stack(bx, level, q.values);
    if (q.total) valueLabel(bx, base - level * k, '+' + mtgAmount_(q.total));
    axis(bx, q.label, q.sub);
  });
  let bx = xOf(4);
  stack(bx, 0, thisValues);
  valueLabel(bx, base - thisTotal * k, mtgAmount_(thisTotal));
  axis(bx, fy, '合計');
  bx = xOf(5);
  if (plan > 0) {
    el.push({type: 'box', x: bx, y: base - plan * k, w: bw, h: plan * k, fill: c.white, line: c.ink, lineWeight: 0.75, shape: 'RECTANGLE', role: 'plan-years-plan'});
    valueLabel(bx, base - plan * k, mtgAmount_(plan));
    el.push({type: 'line', x, y: base - plan * k, w: xOf(5) - x, h: 0, color: C.planLine.color, weight: 0.75, dash: true, role: 'plan-years-plan-line'});
  }
  axis(bx, '計画', fy);
  bx = xOf(6);
  const gap = plan - thisTotal;
  if (plan > 0 && gap > 0) {
    el.push({type: 'box', x: bx, y: base - plan * k, w: bw, h: gap * k, fill: '#EEEEEE', line: '#BBBBBB', lineWeight: 0.5, dashStyle: 'DASH', shape: 'RECTANGLE',
      role: 'plan-years-gap'});
    valueLabel(bx, base - plan * k, '▲' + mtgAmount_(gap), c.minus);
  } else if (plan > 0) {
    valueLabel(bx, base - thisTotal * k, '+' + mtgAmount_(-gap), c.greenDark);
  }
  axis(bx, plan > 0 && gap <= 0 ? '超過' : '不足', fy);
  // 区切り線と、来年度以降（今年度の合計の高さから点線でつなぎ、続けて積み上げる）
  const sepX = xOf(7) - (slot - bw) / 2 - slot * 0.3;
  el.push({type: 'line', x: sepX, y: chartTop - 10, w: 0, h: base - chartTop + 10, color: c.line, weight: 0.75});
  el.push(mtgBox_(sepX + 4, chartTop - 22, 320, 12, 'FY' + yy(1) + '以降：半期の売上（' + mtgPlanMonthName_() + '・確度によらない）', {size: 7, color: c.muted}));
  el.push({type: 'line', x: xOf(4) + bw, y: base - thisTotal * k, w: xOf(7) - xOf(4) - bw, h: 0, color: '#B0B0B0', weight: 0.5, dash: true});
  level = thisTotal;
  later.forEach((l, i) => {
    const lbx = xOf(7 + i);
    level = stack(lbx, level, l.values);
    if (l.total) valueLabel(lbx, base - level * k, '+' + mtgAmount_(l.total));
    axis(lbx, l.label, l.sub);
  });
  bx = xOf(11);
  stack(bx, 0, sumValues([thisValues].concat(later.map(l => l.values))));
  // 累計の棒：年度の中の前半期・後半期の境目に白い点線（どちらの半期にも売上があるときだけ）
  let cum = thisTotal;
  later.forEach(l => {
    cum += l.total;
    const pair = later.find(v => v.year === l.year && v.half === '後');
    if (l.half === '前' && l.total > 0 && pair && pair.total > 0) {
      el.push({type: 'line', x: bx, y: base - cum * k, w: bw, h: 0, color: c.white, weight: 1.5, dashStyle: 'DOT', role: 'plan-years-half-line'});
    }
  });
  valueLabel(bx, base - grand * k, mtgAmount_(grand));
  axis(bx, '累計', span);
  el.push({type: 'line', x, y: base, w, h: 0, color: c.line, weight: 0.75});
  // 確定（A＋B）のページでないときは、今年度の分の確度（A〜D。除いた確度）と、累計が状況・計画進捗の全体合計と同じことを書く
  const lead = (only ? word + ' ' + fy + '売上' : fy + '売上（' + rangeWord + (dropped.length ? '。' + dropped.join('・') + 'を除く' : '') + '）') + 'の積み上げは ' +
    mtgAmount_(thisTotal) + '百万円' +
    (plan > 0 ? 'で、計画 ' + mtgAmount_(plan) + '百万円に対して ' + (gap > 0 ? '▲' + mtgAmount_(gap) : '+' + mtgAmount_(-gap)) +
      '（' + (only ? '進捗率' : '達成率') + ' ' + Math.round(thisTotal / plan * 100) + '%）' : '') + '。\n' +
    'FY' + yy(1) + ' ' + mtgAmount_(yearTotal(yy(1))) + '・FY' + yy(2) + ' ' + mtgAmount_(yearTotal(yy(2))) + 'を続けた ' + span + 'の累計は ' +
    mtgAmount_(grand) + '百万円（' + (only ? '' : '状況・計画進捗の全体合計と同じ。') + '失注除く）。';
  // 受注月が空・今年度の外の案件の今年度の売上（四半期に置けないので入れない）
  const aside = ['undated', 'before', 'after'].reduce((t, key) => t + ((p[key] || {}).value || 0), 0);
  const note = aside > 0 ? '※ ' + mtgPlanMonthName_() + 'が空・' + fy + '外の ' + fy + '売上 ' + mtgAmount_(aside) + '百万円は除く' : '';
  return Object.assign({title: MTG_TITLES.plan + '（' + span + '売上）' + (only ? '：' + label : ''), lead, elements: el}, note ? {note} : {});
}

/**
 * 変化のある案件のうち、今ある案件（失注を除く。案件数と同じ数え方）の数。削除した行・失注になった案件は数えない
 * （案件数より多くならないように。2026/11/09）。中身の一部（mtgCombinedLite_）には changedNow を入れてある
 */
function mtgCompactChangedNow_(s) {
  if (s.changedNow != null) return s.changedNow;
  if (!s.changes) return s.kpi.changed;
  return s.changes.filter(ch => ch.category !== 'updated' && !ch.removed && ch.deal && !mtgExcludedStatus_(ch.statusTo)).length;
}

/** 統合版（報告用）で積み上げる計画（想定売上。summary.planAmount。前の版で保存した仕事の中身に無ければ期待値の plan） */
function mtgCompactPlanOf_(s) {
  return s.planAmount || s.plan;
}

/**
 * 全サブインダストリーの状況・計画進捗（統合版の「計画に対する積み上げ（サブインダストリー別）」の図に、「全サブインダストリーの状況」の数字をまとめたもの）。
 * 横棒は確定（確度 A・B）だけを積み上げる（棒の中に値。入らない値は、いちばん右の確度なら棒のすぐ右に「B 7」、間の確度は出さない）。計画線は計画の位置に、透明度のあるグレーの点線（上に「計画」）。
 * 目盛りは、いちばん大きい計画の planScale 倍（か、それより大きい確定）を棒の幅いっぱいにする。進捗率（A＋B）＝ 確定（A＋B）÷ 計画。
 * 金額はどれも想定売上：確定・見込・進捗率は今年度の想定売上の積み上げ（plan.monthColumn の月が今年度の案件。summary.planAmount）、
 * 想定売上合計・前回比（bridge.amountEnd − amountStart）・案件数・変化のある案件は失注を除くすべて。想定売上・月の列が無いシートは、棒の場所に「データなし」と理由。
 * options.withC：次のページ（見込（C）を含む）。棒の確定（A・B）に続けて見込（C）を、薄い地と上・下・右の細かい点線で足す（compact.prospect。
 * 並び・右の数字は同じ。目盛りは A＋B＋C も入るように）
 */
function mtgCompactOverviewSlide_(summaries, options) {
  // options.sales：今年度の売上（FY前半期＋後半期の列。summary.planSales）で集計したページ（棒・計画・進捗率・確定・見込）。
  // 右の列の売上合計は3年度分（今年度から）の売上、前回比は今年度の売上の増減（summary.salesBridge）。案件数・変化のある案件は同じ
  const bySales = !!(options && options.sales);
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const G = o.gapText;
  const C = o.compact;
  const withC = !!(options && options.withC);
  // 棒は確定（A・B）。withC なら見込（C）を点線の枠で続ける（dashed）
  const ranks = o.plan.ranks.filter(r => r.key === 'A' || r.key === 'B' || (withC && r.key === 'C'))
    .map(r => Object.assign({}, r, {dashed: r.key === 'C'}));
  const barKeys = ranks.map(r => r.key);
  // 前回比の見出し：いちばん多い比べ方（統合版の状況のページと同じ）。前回の会議と比べるなら、比べた会議日を括弧書き（「前回比（11/6）」）
  const versus = s => mtgVersus_(s).trim().replace(/で$/, '');
  const counts = {};
  summaries.forEach(s => { counts[versus(s)] = (counts[versus(s)] || 0) + 1; });
  const common = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || '前回比';
  const commonOf = summaries.find(s => versus(s) === common && s.previous && s.previous.date);
  const deltaLabel = !/比$/.test(common) ? common + 'の増減'
    : common === '前回比' && commonOf ? '前回比（' + mtgShortDate_(commonOf.previous.date) + '）' : common;
  const sum = (p, keys) => keys.reduce((t, k) => t + (p.byRank[k] || 0), 0);
  const cnt = (p, keys) => keys.reduce((t, k) => t + ((p.byRankCount || {})[k] || 0), 0);
  const rate = p => (p.gap || !(p.plan > 0) ? null : sum(p, ['A', 'B']) / p.plan * 100);
  const planOf = s => (bySales ? s.planSales || mtgSalesPlanOf_([], s.name, new Date(), {}) : mtgCompactPlanOf_(s));
  // 並び：シートの順（DIFF_RULES の順。summaries の順）に固定する。どのページも同じ並び（2026/11/10。進捗率の順に並べ替えると、
  // ページや会議ごとにサブインダストリーの位置が変わり、隣の行の数字と見間違えやすかったため）
  const items = summaries.map(s => ({s, p: planOf(s), lack: s.gaps || {}}));
  const ok = items.filter(x => !x.p.gap);
  const merged = mtgMergePlans_(ok.map(x => x.p));
  const first = mtgMonthOf_(merged.start);
  const last = mtgMonthOf_(merged.start + 11);
  const period = String(first.y).slice(2) + '/' + first.m + '〜' + String(last.y).slice(2) + '/' + last.m;
  // 合計（作れたシートの分だけ。作れないシートがあれば ※）
  // 売上のページ（bySales）の右の2列：3年度分（今年度から）の売上の合計と、今年度の売上の前回比（summary.salesBridge）。
  // 想定売上のページは、全期間の想定売上合計と、その前回比（summary.bridge）
  const yearKey = merged.year || mtgFiscalYearKey_(summaries.length && summaries[0].now ? new Date(summaries[0].now) : new Date());
  const span = 'FY' + yearKey + '〜' + ('0' + ((Number(yearKey) + 2) % 100)).slice(-2);
  const lackOf = s => (bySales ? (s.planSales && s.planSales.gap) || !s.salesBridge || s.salesBridge.sureStart == null : (s.gaps || {}).amount);
  const endOf = s => (bySales ? s.salesBridge.total : s.bridge.amountEnd || 0);
  // 前回比：売上のページは確定（A＋B）の今年度の売上の増減（確定（A＋B）/件の右の列）。想定売上のページは全期間の想定売上の増減
  const deltaOf = s => (bySales ? s.salesBridge.sureEnd - s.salesBridge.sureStart : (s.bridge.amountEnd || 0) - (s.bridge.amountStart || 0));
  items.forEach(z => { z.lackAmount = !!lackOf(z.s); });
  const total = {active: 0, amount: 0, delta: 0, changed: 0, lack: {}};
  items.forEach(({s, lack, lackAmount}) => {
    const b = s.bridge;
    total.active += b.endCount;
    if (lackAmount) total.lack.amount = true;
    else total.amount += endOf(s);
    if (lackAmount || lack.changes) total.lack.delta = true;
    else total.delta += deltaOf(s);
    if (lack.changes) total.lack.changed = true;
    else total.changed += mtgCompactChangedNow_(s);
  });
  const reached = ok.filter(x => rate(x.p) >= 100).length;
  const fy = 'FY' + (merged.year || mtgFiscalYearKey_(summaries.length && summaries[0].now ? new Date(summaries[0].now) : new Date()));
  const totalRate = ok.length ? rate(merged) : null;
  const lead = '進捗率（A＋B）は ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') + '（計画に届いているのは ' + ok.length + 'のうち ' + reached + '）。' +
    '案件 ' + total.active + '件・' + (bySales ? span + '売上合計 ' + mtgAmount_(total.amount) + '百万円（失注除く）。\n'
      : '想定売上合計 ' + mtgAmount_(total.amount) + '百万円（全期間・失注除く）。\n') +
    'サブインダストリーはシートの順。' +
    (bySales ? '確定・見込は' + mtgPlanMonthName_() + 'が今年度（' + period + '）の案件の ' + fy + '売上。進捗率 ＝ A＋B ÷ 計画。'
      : '進捗率・確定（A＋B）・見込（C）は今年度（' + period + '）の' + mtgPlanMonthName_() + 'の案件の想定売上（進捗率 ＝ A＋B ÷ 計画）。');
  // 見込（C）を含むページのリード文：確定＋見込の計画比
  const withCRate = ok.length && merged.plan > 0 ? Math.round(sum(merged, ['A', 'B', 'C']) / merged.plan * 100) : null;
  const leadC = '確定（A＋B）に見込（C）を足すと ' + mtgAmount_(sum(merged, ['A', 'B', 'C'])) + '百万円で、計画の ' + (withCRate == null ? '—' : withCRate + '%') +
    '（進捗率（A＋B）は ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') + '）。\n' +
    '見込（C）は点線の枠。' + (bySales ? mtgPlanMonthName_() + 'が今年度（' + period + '）の案件の ' + fy + '売上。' : '') +
    '並び・右の数字は前のページと同じ。';
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 480, 20, bySales ? fy + ' の売上（前半期＋後半期）の確定（A＋B）' + (withC ? 'と見込（C）' : '') + 'と計画（百万円）'
    : '今年度（' + period + '）の確定（A＋B）' + (withC ? 'と見込（C）' : '') + 'の想定売上と計画（百万円）',
    {size: 10.5, bold: true, color: c.ink})];
  // 列：名前・横棒｜計画・進捗率（A＋B）・確定（A＋B）/件・（売上のページは前回比）・見込（C）/件｜想定売上合計（売上のページは売上合計）・（想定売上のページは前回比）・案件数・変化のある案件
  const nameW = 96;
  const thisYear = bySales ? '（FY' + yearKey + '）' : '（今年度）';
  // 案件数・変化のある案件は受注月・売上で絞らない（失注を除くすべての案件）。売上のページは売上合計と同じ「（FY26〜28）」と書く
  const allYears = bySales ? '（' + span + '）' : '（全期間）';
  const groups = [
    // 計画・進捗率・確定・見込は今年度の受注月の案件だけ：見出しの下に「（今年度）」。売上のページは今年度の売上の列なので「（FY26）」
    // 見出しは折り返さず1行（入らなければ字を小さくする）。その下の行に対象の期間
    {cols: [{key: 'plan', label: '計画', w: 36, period: thisYear}, {key: 'rate', label: '進捗率（A＋B）', w: 52, period: thisYear},
      {key: 'ab', label: '確定（A＋B）/件', w: 66, period: thisYear}]
      // 売上のページの前回比は確定（A＋B）の増減なので、確定（A＋B）/件の右
      .concat(bySales ? [{key: 'delta', label: deltaLabel, w: 58, period: thisYear}] : [])
      .concat([{key: 'c', label: '見込（C）/件', w: 60, period: thisYear}])},
    // 想定売上合計・前回比・案件数・変化のある案件は、受注月で絞らない（今年度の列と違う）ので、見出しの下に「（全期間）」
    // 売上のページは、売上合計（3年度分。見出しの下に「（FY26〜28）」）。前回比は左のまとまり（確定（A＋B）の右）
    {cols: [bySales ? {key: 'amount', label: '売上合計', w: 48, period: '（' + span + '）'} : {key: 'amount', label: '想定売上合計', w: 48, period: '（全期間）'}]
      .concat(bySales ? [] : [{key: 'delta', label: deltaLabel, w: 58, period: '（全期間）'}]).concat([
      // 売上のページは期間が「（FY26〜28）」と長いので、案件数の列を広げる（折り返さないように。2026/11/10）
      {key: 'active', label: '案件数', w: bySales ? 44 : 36, period: allYears}, {key: 'changed', label: '変化のある案件', w: 56, period: allYears}])}
  ];
  // 列はまとまりに分けて見出しを付けない（2026/11/09。計画進捗と案件の状況を分ける必要はないため）。groups は列の並びのまとまり
  const colsW = groups.reduce((t, gr) => t + gr.cols.reduce((u, col) => u + col.w, 0), 0);
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  // 計画線の見出し（「計画」）の行と、列の見出しの行
  const groupY = y + 21;
  const headY = groupY + 15;
  // 見出しの行（1行）と、その下の期間の行の高さ
  const labelH = 12;
  const periodH = 10;
  const headH = labelH + periodH;
  const rows = items.concat([{total: true, s: {name: '合計'}, p: merged, lack: total.lack}]);
  const rowH = Math.min(22, (S.bottom - headY - headH - 2) / rows.length);
  // 目盛り：いちばん大きい計画が棒の幅の 1/planScale ほどの位置に来るように（計画線の右に、計画を超えた分を描ける）
  const max = Math.max.apply(null, [1].concat(ok.map(item => Math.max(sum(item.p, ['A', 'B']), item.p.plan * C.planScale))).concat(ok.map(item => sum(item.p, barKeys))));
  const scale = barW / max;
  let cx = barX + barW + 10;
  // 見出しは折り返さずに1行。どの列も同じ大きさの字で、いちばん長い見出しが列の幅の 9割に入る大きさ（7.5pt から 6pt まで。
  // Google スライドで折り返すと見出しの高さがそろわず見づらいため）。その下の同じ高さの行に、対象の期間（「（今年度）」「（全期間）」。小さめの灰色の字）
  let size = 7.5;
  const fitsAll = pt => groups.every(gr => gr.cols.every(col => mtgTextWidth_(col.label, pt) <= col.w * 0.9));
  while (!fitsAll(size) && size > 6) size -= 0.5;
  // 期間の字も折り返さない：どの列も同じ大きさで、いちばん長い期間が列の幅に入る大きさ（6.5pt から 5.5pt まで）
  let periodSize = 6.5;
  const periodsFit = pt => groups.every(gr => gr.cols.every(col => !col.period || mtgTextWidth_(col.period, pt) <= col.w));
  while (!periodsFit(periodSize) && periodSize > 5.5) periodSize -= 0.5;
  groups.forEach(gr => {
    gr.cols.forEach(col => {
      el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, labelH, col.label,
        {size, color: c.sub, align: 'CENTER', valign: 'BOTTOM', role: 'compact-head'}));
      if (col.period) {
        el.push(mtgBox_(cx - S.padX / 2, headY + labelH, col.w + S.padX, periodH, col.period,
          {size: periodSize, color: c.muted, align: 'CENTER', valign: 'MIDDLE', role: 'compact-head-period'}));
      }
      col.x = cx;
      cx += col.w;
    });
  });
  const cols = [].concat.apply([], groups.map(gr => gr.cols));
  // 凡例：グラフの上（列の見出しの行の、棒の左端から）に確度 A・B の色。計画線は線の上に見出し（「計画」）を付けるので凡例に入れない
  let lx = barX;
  // 見本は A・B・C とも同じ大きさ（8 × 8）。見込（C）は棒と同じ薄い地と細かい点線の枠。見本と文字の間は 5pt あける
  const P = C.prospect;
  ranks.forEach(r => {
    const sw = 8;
    const sh = 8;
    // 凡例は期間の行の高さ（棒のすぐ上）
    const ly = headY + headH - periodH / 2 - sh / 2;
    el.push(r.dashed ? {type: 'box', x: lx, y: ly, w: sw, h: sh, fill: r.color, alpha: P.alpha, line: P.line,
      lineWeight: P.lineWeight, dashStyle: P.dashStyle, shape: 'RECTANGLE', role: 'legend'}
      : {type: 'box', x: lx, y: ly, w: sw, h: sh, fill: r.color, shape: 'RECTANGLE', role: 'legend'});
    const bw = mtgTextWidth_(r.key, 8.5) + S.padX + 2;
    el.push(mtgBox_(lx + sw + 5 - S.padX / 2, ly - 4, bw, sh + 8, r.key, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += sw + 5 + bw - S.padX + 10;
  });
  // 計画線の見出し：まとまりの見出しの高さに「計画」と、そこから1行目までの点線（計画の位置ごと。近い位置（24pt 以内）は1つにまとめる）
  const line = C.planLine;
  const planXs = [];
  ok.filter(z => z.p.plan > 0).forEach(z => {
    const px = barX + z.p.plan * scale;
    if (!planXs.some(v => Math.abs(v - px) < 24)) planXs.push(px);
  });
  planXs.forEach(px => {
    el.push(mtgBox_(px - 30, groupY, 60, 12, '計画', {size: 8, bold: true, color: line.color, align: 'CENTER', role: 'compact-plan-label'}));
    el.push({type: 'line', x: px, y: groupY + 13, w: 0, h: headY + headH + 3 - groupY - 13, color: line.color, alpha: line.alpha, weight: line.weight,
      dash: true, role: 'compact-plan'});
  });
  rows.forEach((item, i) => {
    const ry = headY + headH + 2 + rowH * i;
    const p = item.p;
    const s = item.s;
    if (item.total) el.push({type: 'box', x, y: ry, w, h: 0.75, fill: c.line, shape: 'RECTANGLE'});
    else if (i % 2 === 1) el.push({type: 'box', x, y: ry, w, h: rowH, fill: c.greenSoft, alpha: 0.35, shape: 'RECTANGLE', role: 'compact-band'});
    el.push(mtgBox_(x, ry, nameW, rowH, s.name, {size: 8.5, bold: true, color: c.ink, role: 'compact-name'}));
    // 横棒（確度 A・B の積み上げ）と計画線。合計の行は描かない（目盛りが違う。全体の図は次のページ）
    if (!item.total && !p.gap) {
      let level = 0;
      const bh = Math.max(rowH - 6, 6);
      const shown = ranks.filter(r => (p.byRank[r.key] || 0) > 0);
      let end = barX;
      shown.forEach((r, k) => {
        const v = p.byRank[r.key];
        const bx = barX + level * scale;
        const bw = Math.max(v * scale, 0.75);
        const by = ry + (rowH - bh) / 2;
        if (r.dashed) {
          // 見込（C）は確定の棒に続けて（積み上げ）、薄い地（枠なし）と、上・下・右の3辺の細かい点線（左の辺は確定の棒との境目なので引かない。
          // 左の辺が確定の棒の端に重なると、つなぎ目が乱れて見えるため）。確定の棒が無ければ左の辺も引く
          el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: r.color, alpha: P.alpha, shape: 'RECTANGLE', role: 'compact-bar', prospect: true});
          const edge = (x1, y1, w1, h1) => el.push({type: 'line', x: x1, y: y1, w: w1, h: h1, color: P.line, weight: P.lineWeight, dashStyle: P.dashStyle,
            role: 'compact-bar-edge'});
          edge(bx, by, bw, 0);
          edge(bx, by + bh, bw, 0);
          edge(bx + bw, by, 0, bh);
          if (level === 0) edge(bx, by, 0, bh);
        } else {
          el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: r.color, line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role: 'compact-bar'});
        }
        end = bx + bw;
        // データラベル：棒の中（白い字。見込は濃い緑の字。入らなければ 6pt まで小さく）。それでも入らない値は、棒のいちばん右（最後の確度）なら
        // 棒のすぐ右に「C 8」のように出し、間の確度（右に別の確度が続く）は出さない（離れた場所に出すと、どの棒の値か分かりにくいため。値は右の表にある）
        const label = mtgAmount_(Math.round(v * 10) / 10);
        let size = 7.5;
        while (mtgTextWidth_(label, size) + 3 > bw && size > 6) size -= 0.5;
        if (mtgTextWidth_(label, size) + 3 <= bw) {
          el.push(mtgBox_(bx - S.padX / 2, ry, bw + S.padX, rowH, label, {size, bold: true, color: r.dashed ? c.greenDark : c.white, align: 'CENTER',
            role: 'compact-bar-label'}));
        } else if (k === shown.length - 1) {
          const text = r.key + ' ' + label;
          el.push(mtgBox_(end + 3 - S.padX / 2, ry, mtgTextWidth_(text, 7.5) + S.padX + 2, rowH, text,
            {size: 7.5, bold: true, color: r.dashed ? c.greenDark : r.color, role: 'compact-bar-label'}));
        }
        level += v;
      });
      if (p.plan > 0) {
        el.push({type: 'line', x: barX + p.plan * scale, y: ry + 1, w: 0, h: rowH - 2, color: line.color, alpha: line.alpha, weight: line.weight, dash: true,
          role: 'compact-plan'});
      }
    } else if (!item.total) {
      el.push(mtgBox_(barX, ry, barW, rowH, mtgWrap_(G + '（' + p.gap + '）', barW, 7.5, 1)[0], {size: 7.5, color: c.muted, role: 'compact-gap'}));
    }
    const lack = item.lack || {};
    const b = s.bridge || {};
    const r = item.total ? totalRate : p.gap ? null : rate(p);
    const amountCount = keys => (p.gap || (item.total && !ok.length) ? G : mtgAmount_(Math.round(sum(p, keys) * 10) / 10) + '/' + cnt(p, keys) + '件');
    const mark = key => (total.lack[key] ? '※' : '');
    const values = item.total ? {
      plan: ok.length ? mtgAmount_(p.plan) : G, ab: amountCount(['A', 'B']), c: amountCount(['C']),
      amount: items.every(z => z.lackAmount) ? G : mtgAmount_(total.amount) + mark('amount'),
      delta: items.every(z => z.lackAmount || z.lack.changes) ? G : mtgSigned_(total.delta) + mark('delta'),
      active: total.active + '件', changed: items.every(z => z.lack.changes) ? G : total.changed + '件' + mark('changed')
    } : {
      plan: p.gap ? G : mtgAmount_(p.plan), ab: amountCount(['A', 'B']), c: amountCount(['C']),
      amount: item.lackAmount ? G : mtgAmount_(endOf(s)), delta: item.lackAmount || lack.changes ? G : mtgSigned_(deltaOf(s)),
      active: b.endCount + '件', changed: lack.changes ? G : mtgCompactChangedNow_(s) + '件'
    };
    cols.forEach(col => {
      if (col.key === 'rate') {
        el.push(mtgBox_(col.x - S.padX / 2, ry, col.w + S.padX, rowH, r == null ? '—' : Math.round(r) + '%',
          {size: 9, bold: true, color: r != null && r >= 100 ? c.greenDark : c.ink, align: 'END', role: 'compact-value'}));
        return;
      }
      const text = String(values[col.key]);
      const gap = text === G;
      // 前回比：増えたら青、減ったら赤（報告案件版の期待値の増減と同じ色）
      const sign = col.key === 'delta' && !gap ? (/^\+/.test(text) ? c.plus : /^▲/.test(text) ? c.minus : null) : null;
      el.push(mtgBox_(col.x - S.padX / 2, ry, col.w + S.padX, rowH, text,
        {size: gap ? 7.5 : col.key === 'ab' || col.key === 'c' ? 8 : 8.5, bold: !!item.total || col.key === 'ab',
          color: gap ? c.muted : sign || c.ink, align: 'END', role: 'compact-value'}));
    });
  });
  const marked = Object.keys(total.lack).length > 0;
  // タイトル：「（見込（C）を含む）」「（FY26売上）」「（FY26売上・見込（C）を含む）」
  const suffix = [bySales ? fy + '売上' : '', withC ? '見込（C）を含む' : ''].filter(Boolean).join('・');
  return Object.assign({title: MTG_TITLES.compactOverview + (suffix ? '（' + suffix + '）' : ''), lead: withC ? leadC : lead, elements: el}, marked ? {note: '※ は「' + G + '」のシートを除いた合計'} : {});
}

/**
 * 変化のあった案件：{サブインダストリー}（1つのサブインダストリーのページ。行が入らなければ次のページに分け、どのページにも「（1/2ページ）」）。変化のある案件（その他の項目のみの更新を除く。
 * 各サブインダストリーの資料の「前回から変化のある案件」と同じ案件・同じ並び：今回のステータスの確度 A から順（失注・保留・削除は最後）、その中は想定売上の大きい順）を、No.（start + 1 からの資料の通し番号）とともに全件。
 * 変化が無ければ表の見出しの下に「該当なし」（変更の記録が無ければ「データなし」と理由）。
 * 表は出典の行にかからないよう、区切り線より tableMargin 上までで次のページに分ける。戻り値 {slides（ページ番号・出典の行は付けない）, count: 載せた案件の数}
 */
function mtgCompactChangeSlides_(s, start) {
  return mtgWithAmountCeil_(() => mtgCompactChangeSlidesOf_(s, start));
}

/** mtgCompactChangeSlides_ の中身（金額の繰り上げの設定の中で作る） */
function mtgCompactChangeSlidesOf_(s, start) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const word = mtgFromWord_(s);
  const title = o.compact.changesTitle + s.name;
  // 金額の列：今年度の売上（salesPage。今の値。前回の売上は記録に無いため）か、想定売上（前回▶今回）
  const bySales = !!o.compact.salesPage;
  const year = mtgFiscalYearKey_(s.now ? new Date(s.now) : new Date());
  const salesOf = ch => (ch.deal && ch.deal.sales ? ch.deal.sales[year] : null);
  const lack = s.gaps || {};
  // 並び：今回のステータスの期待値の高い順 → 失注・保留・削除、同じ期待値は今年度の売上の大きい順
  // （salesPage。想定売上のときは各サブインダストリーの資料と同じ想定売上の大きい順）
  // 2026/11/16 から、今回のステータスの期待値（「B-クロージング中(70%)」の %。mtgCompactStatusRate_）の高い順に並べ、同じ期待値は売上の大きい順
  // （同じ確度の文字でも期待値の違うステータスがあれば、期待値で分ける）。失注・保留・削除は最後
  const moves = s.changes.filter(ch => ch.category !== 'updated').map((ch, i) => ({ch, i, rate: mtgCompactStatusRate_(ch)}))
    .sort((a, b) => (b.rate - a.rate) || (bySales ? (salesOf(b.ch) || 0) - (salesOf(a.ch) || 0) : 0) || (a.i - b.i)).map(x => x.ch);
  // セルの中身（列の幅を決めるのにも使う）。「前 ▶ 今」は、1行に入らなければ矢印の前（ステータス）・あと（受注月・金額）で改行する
  const statusText = ch => mtgStatusMove_(ch);
  const salesText = ch => (ch.removed ? '（削除）' : salesOf(ch) != null ? mtgAmount_(salesOf(ch)) : o.blankText);
  const amountMove = ch => mtgValueMove_(ch.base ? ch.base.amount : null, ch.deal ? ch.amount : null, ch);
  const labels = ['No', mtgDealHead_(s), 'ステータス\n' + word + '▶今回', bySales ? 'FY' + year + '売上\n（今回）' : '想定売上\n' + word + '▶今回',
    '受注月\n' + word + '▶今回', '活動状況（最新）', '活動における課題'];
  // 列の幅（2026/11/16）：活動状況・活動における課題は文が長くなりやすいので広く、ほかの列は見出しと中身が折り返さずに入る幅まで狭める。
  // 幅は見出し・中身の実際の文字から決める（見出しは太字で、見積もりより Google スライドの字が広いので余裕を大きく（compact.headFit）。
  // 「FY26売上」が「FY26売 / 上」と折り返した。2026/11/16）。No は3桁が 7.5pt で1行に入る幅（26pt では「39／2」と折り返した）
  const columns = mtgCompactChangeColumns_(labels, [
    {min: 30},
    {min: 100},
    {size: 8, texts: moves.map(statusText), arrow: 'before'},
    {texts: moves.map(ch => (bySales ? salesText(ch) : String(amountMove(ch)))), arrow: 'after'},
    {texts: moves.map(ch => mtgMonthMove_(ch)), arrow: 'after'}
  ]);
  // 得意先／案件名の列：RXシナリオ構成の列があれば、案件名の下にシナリオの番号（mtgDealScenarioCell_。見出しは「得意先／案件名／シナリオNo」。2026/11/16）
  const head = mtgCompactHeadRow_(columns, labels);
  const noSize = text => {
    let size = 7.5;
    while (mtgTextWidth_(text, size) > (columns[0] - S.cellPadX) * 0.8 && size > 5.5) size -= 0.5;
    return size;
  };
  const gone = moves.filter(ch => ch.removed || !ch.deal || mtgExcludedStatus_(ch.statusTo)).length;
  // リード文は2行（どちらも折り返さない長さ。2026/11/16）：1行目に件数（失注・削除の案件も載せる。状況・計画進捗の「変化のある案件」は今ある案件だけを
  // 数える（mtgCompactChangedNow_）ので、その分を書き添える）、2行目に並びと、その他の項目のみの更新の件数
  const updated = s.kpi && s.kpi.updated ? 'ほかにその他の項目のみの更新 ' + s.kpi.updated + '件。' : '';
  const lead = lack.changes ? mtgChangedLabel_(s) + 'は' + o.gapText + '（' + lack.changes + '）。'
    : moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + (gone ? '（うち失注・削除 ' + gone + '件は、状況・計画進捗の件数に含めない）' : '') + '。\n' +
      '並びはステータスの期待値の高い順・' + (bySales ? 'FY' + year + '売上' : '想定売上') + 'の大きい順。No は通し番号。' + updated
      : mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし' + mtgUpdatedNote_(s) + '。';
  if (!moves.length) {
    return {slides: [{title, lead, elements: mtgEmptyTable_(columns, head, lack.changes ? 'gap' : 'none', lack.changes || '')}], count: 0};
  }
  const rows = moves.map((ch, i) => {
    const status = statusText(ch);
    return mtgRow_(columns, [
      // No：3桁（「392」）でも折り返さないように、小さめの字（7.5pt）で1行。セルの幅の 8割に入らなければさらに小さくする
      // （見積もりぎりぎりだと、Google スライドでは折り返すことがあるため）
      {text: String(start + i + 1), size: noSize(String(start + i + 1)), lines: 1, align: 'CENTER', fill: c.greenSoft, color: c.ink, bold: true},
      mtgDealScenarioCell_(s, ch, columns[1]),
      // ステータスの動きが1行に入らなければ「▶」の前で改行（「B-／クロージング中」のように名前の途中で切らない）
      // 字は 8pt（列を狭めて活動状況・課題を広くしたため。「▶ B-クロージング中(70%)」が1行に入る大きさ）
      {text: mtgCompactLineFits_(status, 8, columns[2]) ? status : status.replace(' ▶ ', '\n▶ '), size: 8, lines: 3, fill: c.white},
      // 今年度の売上（今の値。空なら「未入力」、削除した行は「（削除）」）か、想定売上（前回▶今回。変わっていなければ今の値）
      bySales ? {text: salesText(ch), lines: 1, fill: c.white, align: 'CENTER',
        color: salesOf(ch) != null ? c.body : c.muted}
        : mtgFieldCell_(s, 'amount', amountMove(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      // 活動状況・課題は4行まで今の字の大きさで（長い文を小さくしすぎないように。2026/11/16）。入らなければ小さくし、それでも入らなければ行を伸ばす
      mtgFieldCell_(s, 'activity', ch.activity, {lines: 4, fill: c.white, fit: true, grow: true}),
      mtgFieldCell_(s, 'issue', ch.issue, {lines: 4, fill: c.white, fit: true, grow: true})
    ]);
  });
  // ページに分ける（区切り線より tableMargin 上まで。入らない行は次のページ）
  const limit = S.bottom - o.compact.tableMargin;
  const pages = [];
  rows.forEach(row => {
    let page = pages[pages.length - 1];
    if (!page || (page.rows.length > 1 && page.y + row.h > limit)) {
      page = {rows: [head], y: S.top + head.h};
      pages.push(page);
    }
    page.rows.push(row);
    page.y += row.h;
  });
  // 2ページ以上なら、どのページのタイトルにも「（1/3ページ）」（「（続き）」より、全部で何ページかが分かる）
  return {slides: pages.map((p, n) => ({title: title + (pages.length > 1 ? '（' + (n + 1) + '/' + pages.length + 'ページ）' : ''), lead, elements: [{type: 'table', x: S.x, y: S.top, columns, rows: p.rows}]})),
    count: moves.length};
}

/* ---------------- 作る（区切りながら） ---------------- */

/** 統合版（報告用）を作れるか（作れなければ {error: 理由}、作れれば {text: 確かめの文}） */
function mtgCompactCheck_(meeting, point, cancelJob, part) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const o = MTG_OPTIONS;
  if (!o.compact.enabled) return {error: '統合版（報告用）を使わない設定です（MeetingSummary.gs の MTG_OPTIONS.compact.enabled）。'};
  // 作っている途中の仕事：「やめて新しく作る」（cancelJob）なら、作り始めるときにやめる（mtgJobCancel_）
  const pending = mtgJobLoad_();
  if (pending && !cancelJob) {
    return {error: (mtgDeckJob_(pending) ? mtgDeckJob_(pending).label : 'サブインダストリーの資料') + 'を作っている途中です（自動で続けています）。' +
      '終わってから（showBusinessMtgSummariesResult() で確かめてから）実行してください。'};
  }
  const names = mtgTrackedNames_(ss);
  if (!names.length) return {error: '追跡しているシート（test.gs の DIFF_RULES）が見つかりません。'};
  const pointText = '比べる時点：' + (point ? mtgChosenText_(point) + '（全シート共通）' : '前回の会議（シートごと）');
  if (part === 'changes') {
    return {text: '「' + o.combined.name + '」の' + o.compact.label + 'の変化のあった案件（表紙と、' + names.length + 'サブインダストリーの変化のあった案件のページ。' +
      '状況・計画進捗と計画に対する積み上げのページは作りません）を、今のシートの値で作ります。\n\n' +
      '会議日 ' + mtgDateLabel_(meeting) + ' の資料として作り、保存先のフォルダに保存します（会議のフォルダには格納しません。会議日の記録・時点コピー・インデックスも変えません）。\n' +
      pointText};
  }
  if (part === 'overview') {
    return {text: '「' + o.combined.name + '」の' + o.compact.label + 'の全体サマリ（' + names.length + 'サブインダストリーの状況・計画進捗と計画に対する積み上げ。' +
      '変化のあった案件のページは作りません）を、今のシートの値で作ります。\n\n' +
      '会議日 ' + mtgDateLabel_(meeting) + ' の資料として作り、保存先のフォルダに保存します（会議のフォルダには格納しません。会議日の記録・時点コピー・インデックスも変えません）。\n' +
      '比べる時点：' + (point ? mtgChosenText_(point) + '（全シート共通）' : '前回の会議（シートごと）')};
  }
  return {text: '「' + o.combined.name + '」の' + o.compact.label + '（' + names.length + 'サブインダストリーの状況・計画進捗と、変化のあった案件）を、' +
    '今のシートの値で作ります。\n\n' +
    '会議日 ' + mtgDateLabel_(meeting) + ' の資料として作り、会議のフォルダに格納します（会議日の記録・時点コピー・インデックスは変えません）。\n' +
    '比べる時点：' + (point ? mtgChosenText_(point) + '（全シート共通）' : '前回の会議（シートごと）')};
}

/**
 * 統合版（報告用）を作り始める（画面の「作成する」から）。時間がかかれば区切って続ける（mtgCompactRun_。続きは continueBusinessMtgSummaries）。
 * 戻り値は終わり（区切ったとき）の知らせ
 */
function mtgCompactStart_(meeting, point, overwrite, cancelJob, part) {
  const started = mtgNow_();   // 6分の上限は、この実行を始めたときから
  if (cancelJob) mtgJobCancel_();   // 「やめて新しく作る」：作っている途中の仕事をやめる
  const check = mtgCompactCheck_(meeting, point);
  if (check.error) throw new Error(check.error);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = mtgTrackedNames_(ss);
  // 変化のあった案件のページの前の RXシナリオ一覧：プルダウンの選択肢だけを先に読み、ページの数を決める（全体サマリだけのときは入れない）
  const scenarioOptions = part !== 'overview' && MTG_OPTIONS.compact.scenarioPage ? mtgCompactScenarioOptions_(ss, names) : [];
  const scenarioPages = scenarioOptions.length ? mtgCompactScenarioSlides_([], scenarioOptions).length : 0;
  const job = {kind: 'compact', final: true, meeting, now: Date.now(), names, point: point || null, items: [], scenarioOptions, scenarioPages,
    page: mtgCompactFrontCount_(part) + scenarioPages, no: 0, deck: null, result: null, folderUrl: '', sourceError: '', overwrite: !!overwrite,
    // 作る範囲：'overview' 全体サマリだけ（createBusinessMtgReportingOverview()。変化のあった案件のページを作らない）・
    // 'changes' 変化のあった案件だけ（createBusinessMtgReportingChanges()。前のページは表紙だけ）
    part: part || ''};
  mtgJobSafetyTrigger_();   // 実行が強制終了されたら（メモリ不足など）、最後に保存した進み具合から続ける
  let done;
  try {
    done = mtgCompactRun_(ss, job, started);
  } catch (error) {
    mtgDeckDiscard_(job.deck);   // 作りかけの資料はゴミ箱へ
    mtgJobClear_();
    mtgJobDeleteTriggers_();
    throw error;
  }
  if (done) mtgJobFinish_();
  const message = mtgCompactMessage_(job, done);
  console.log(message);
  return message;
}

/**
 * 統合版（報告用）を作る仕事を、時間の許す限り進める（統合版の mtgCombinedRun_ と同じ進め方）。
 * job = {kind: 'compact', final: true, meeting, now, names, point, items: [{name, start, pages, lite: mtgCombinedLite_}], page: 次のページの前のページの数,
 *  no: 載せた変化のあった案件の数（通し番号）, deck, result, folderUrl, sourceError, overwrite}。
 * 順に、各サブインダストリーの変化のあった案件のページ（sheetsPerRead シートずつ読む）→ 前のページ（表紙・状況と計画進捗・計画）を足して仕上げ（PDF）→
 * 会議のフォルダへの格納（インデックスには書かない）。上限を越えそうなら、仕事を保存して続きを予約し、false を返す
 */
function mtgCompactRun_(ss, job, started) {
  const o = MTG_OPTIONS;
  const R = o.run;
  const fits = ms => mtgNow_() - started + ms <= R.limitMs;
  // この実行で1つでも進んだか。まだ進んでいなければ、見積もりの時間が足りなくても区切らずに続ける（シートと記録を読むのに時間がかかる
  // ときに区切ると、次の実行も同じだけ読んで区切り、いつまでも終わらなかった。2026/11/10）
  let progressed = false;
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  job.folderUrl = folder.getUrl();
  const part = mtgCompactJobPart_(job);
  const fileName = mtgCompactFileName_(job.meeting, part);
  if (!job.deck) {
    job.deck = mtgDeckStart_(fileName, folder);
    mtgJobCheckpoint_(job);
  }
  while (job.items.length < job.names.length) {
    if (progressed && !fits(R.summaryMs + R.deckMs)) return mtgJobPause_(job);
    const batch = R.readMode === 'api' ? job.names.slice(job.items.length, job.items.length + Math.max(R.sheetsPerRead || 1, 1))
      : job.names.slice(job.items.length);
    const source = mtgJobSource_(ss, job, batch);
    if (progressed && !fits(R.summaryMs + R.deckMs)) return mtgJobPause_(job);
    const summaries = mtgSummaries_(source, batch, new Date(job.now), true, job.meeting, job.point).summaries;
    for (let i = 0; i < summaries.length; i++) {
      if (i > 0 && !fits(R.deckMs)) return mtgJobPause_(job);
      const summary = summaries[i];
      if (part === 'overview') {
        // 全体サマリだけ：前のページに使う中身だけを残す（変化のあった案件のページは作らない）
        job.items.push({name: summary.name, start: 0, pages: 0, lite: mtgCombinedLite_(summary)});
        progressed = true;
        job.crashes = 0;   // 進んだので、止められた回数も数え直す
        mtgJobCheckpoint_(job);
        continue;
      }
      const timer = mschTimer_(o.compact.label + 'のスライド（' + summary.name + '）', true);
      const changes = mtgCompactChangeSlides_(summary, job.no);
      const slides = mtgNumberSlides_(mtgCombinedFrame_([summary], job.meeting, true), changes.slides, job.page);
      mtgDeckAdd_(job.deck, slides, job.page);
      timer.lap('スライドを足す（' + slides.length + '枚）');
      job.items.push({name: summary.name, start: job.page + 1, pages: slides.length, lite: mtgCombinedLite_(summary)});
      job.page += slides.length;
      job.no += changes.count;
      progressed = true;
      job.crashes = 0;   // 進んだので、止められた回数も数え直す
      mtgJobCheckpoint_(job);   // 資料ごとに進み具合を保存（実行が止められても、この資料の次から。mtgCombinedRecover_）
    }
  }
  if (!job.result) {
    if (progressed && !fits(R.combinedFinishMs)) return mtgJobPause_(job);
    const timer = mschTimer_(o.compact.label + 'の仕上げ', true);
    job.deck.finishing = true;   // 仕上げの途中で止められたら、続きからは作れない（作り直す）
    mtgJobCheckpoint_(job);
    // 変化のあった案件のページの前（前のページのすぐあと）に RXシナリオ一覧。形は作り始めに読んだ選択肢で決め（取っておいたページの数と同じ）、
    // 数は読んだシートの中身から
    if (job.scenarioPages) {
      const lites = job.items.map(item => item.lite);
      const start = mtgCompactFrontCount_(part);
      const scenario = mtgNumberSlides_(mtgCombinedFrame_(lites, job.meeting, true), mtgCompactScenarioSlides_(lites, job.scenarioOptions), start)
        .slice(0, job.scenarioPages);
      mtgDeckAdd_(job.deck, scenario, start);
      timer.lap('RXシナリオ一覧');
    }
    mtgDeckAdd_(job.deck, mtgCompactFront_(job.items.map(item => item.lite), job.meeting, true, new Date(job.now), part), 0);
    timer.lap(part === 'changes' ? '表紙' : '表紙・状況と計画進捗・計画のページ');
    const existing = job.overwrite ? mtgExistingPdf_(fileName, o.combined.name, job.meeting, true, folder) : null;
    const out = mtgDeckFinish_(job.deck, fileName, folder, existing);
    timer.lap('並べて PDF に保存');
    const result = {name: o.combined.name, file: out.pdf, pages: job.page, overwritten: !!out.overwritten};
    // 会議のフォルダに格納する（できなくても止めない）。インデックスの統合版の列には書かない。全体サマリだけ・変化のあった案件だけは格納しない
    if (!part) {
      const stored = mtgStoreToMeetingFolder_(o.combined.name, job.meeting, out.pdf);
      if (stored.error) result.storeError = stored.error;
      else result.stored = stored;
    }
    job.result = mtgCombinedPlain_(result);
    job.deck = null;
  }
  return true;
}

/** 統合版（報告用）の仕事の、終わり（done）・区切ったときの知らせ */
function mtgCompactMessage_(job, done) {
  const part = mtgCompactJobPart_(job);
  const label = mtgCompactLabel_(part);
  if (!done) {
    const rest = job.names.slice(job.items.length);
    const left = part === 'overview' ? [].concat(rest.length ? ['シートの読み込み ' + rest.length + '件（' + rest.join('・') + '）'] : [], ['表紙・状況と計画進捗・計画のページと PDF'])
      : [].concat(rest.length ? ['変化のあった案件のページ ' + rest.length + '件（' + rest.join('・') + '）'] : [],
        [(part === 'changes' ? '表紙' : '表紙・状況と計画進捗・計画のページ') + (job.scenarioPages ? '・RXシナリオ一覧' : '') + 'と PDF' +
          (part === 'changes' ? '' : '・会議のフォルダへの格納')]);
    return ['時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りは約' + Math.round(MTG_OPTIONS.run.continueAfterMs / 60000) +
      '分後に自動で続けます：' + left.join('・') + '。',
    '続きが終わったら、エディタで showBusinessMtgSummariesResult() を実行すると結果が出ます（実行ログにも出ます）。終わるまで、会議サマリの関数は実行しないでください。',
    '', label + '：' + job.items.length + '/' + job.names.length + 'サブインダストリーの' + (part === 'overview' ? 'シートを読みました。' : 'ページを作りました。'),
    '', '保存先：' + job.folderUrl].join('\n');
  }
  const r = job.result;
  return ['「' + label + '」：' + r.fileName + '（' + r.pages + 'ページ）\n　' + r.fileUrl +
    (r.overwritten ? '（前の資料に上書き。リンクはそのまま）' : '') +
    (r.folderName ? '\n　会議のフォルダ「' + r.folderName + '」' + (r.created ? 'を作って' : 'に') + '格納しました：' + r.folderUrl : '') +
    (r.storeError ? '\n　※ 会議のフォルダに格納できませんでした（' + r.storeError + '）。資料は保存先にあります。' : ''),
  '', '保存先：' + job.folderUrl].concat(mtgSourceNote_(job)).join('\n');
}
