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
 *   4〜 変化のあった案件：{サブインダストリー}（mtgCompactChangeSlides_。シートの順にサブインダストリーごと）：No（資料の通し番号）・得意先／案件名・
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
  // 売上のページを FY26〜28 のページにする（outlookPage）ときは、状況・計画進捗が1枚（計画に対する積み上げは同じ2枚）
  return 1 + kinds * (MTG_OPTIONS.plan.enabled ? 4 : 2) - (C.salesPage && C.outlookPage ? 1 : 0);
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
  let page = mtgCompactFrontCount_();
  let no = 0;
  const pages = [];
  summaries.forEach(s => {
    const part = mtgCompactChangeSlides_(s, no);
    no += part.count;
    pages.push.apply(pages, mtgNumberSlides_(mtgCombinedFrame_([s], meeting, final), part.slides, page));
    page += part.slides.length;
  });
  return mtgCompactFront_(summaries, meeting, final, now).concat(pages);
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
      if (C.salesPlanBy === 'half' && !C.outlookPage) {
        const sales = mtgMergeSalesPlans_(pick(summaries.map(s => s.planSales || mtgSalesPlanOf_([], s.name, now, {}))));
        plan.push(mtgSalesPlanSlide_(sales, o.combined.name, {title: titles.sales}), mtgSalesPlanSlide_(mtgPlanRanksOnly_(sales, C.sureRanks), o.combined.name,
          {only: C.sureRanks, label: C.sureLabel, title: titles.salesSure}));
      } else {
        // 案件の今年度の売上の合計を、受注月の月に置いて積み上げる（想定売上の積み上げと同じ形）
        const sales = mtgMergePlans_(pick(summaries.map(s => s.planSalesByMonth || mtgSalesPlanByMonthOf_([], s.name, now, {}))));
        sales.year = sales.year || mtgFiscalYearKey_(now);
        const fy = 'FY' + sales.year;
        const measure = fy + '売上';
        // FY26〜28 のページ（outlookPage）：今年度の四半期と来年度以降の半期を続けて積み上げる2枚（全確度・確定（A＋B））
        if (C.outlookPage) {
          plan.push(mtgCompactPlanYearsSlide_(sales, summaries, now), mtgCompactPlanYearsSlide_(sales, summaries, now, {only: C.sureRanks, label: C.sureLabel}));
        } else plan.push(mtgPlanSlide_(sales, now, o.combined.name, {measure, titleNote: measure, title: titles.sales}),
          mtgPlanSlide_(mtgPlanRanksOnly_(sales, C.sureRanks), now, o.combined.name, {only: C.sureRanks, label: C.sureLabel, measure, titleNote: measure,
            title: titles.salesSure}));
      }
    }
    if (merged && amountOn) {
      plan.push(mtgPlanSlide_(merged, now, o.combined.name, {measure: C.measure, title: titles.amount}),
        mtgPlanSlide_(mtgPlanRanksOnly_(merged, C.sureRanks), now, o.combined.name, {only: C.sureRanks, label: C.sureLabel, measure: C.measure,
          title: titles.amountSure}));
    }
    // 状況・計画進捗：今年度の売上（と見込（C）を含む）のあとに、想定売上（同じ2枚。amountPages のとき）
    // 売上のページ：FY26〜28 のページ（outlookPage。2026/11/12 から）か、前の今年度の売上の2枚（FY26売上・見込（C）を含む）
    const status = [].concat(!C.salesPage ? [] : C.outlookPage ? [mtgCompactOutlookSlide_(summaries)]
      : [mtgCompactOverviewSlide_(summaries, {sales: true}), mtgCompactOverviewSlide_(summaries, {sales: true, withC: true})],
      amountOn ? [mtgCompactOverviewSlide_(summaries), mtgCompactOverviewSlide_(summaries, {withC: true})] : []);
    return mtgNumberSlides_(mtgCombinedFrame_(summaries, meeting, final),
      [mtgCombinedCover_(summaries, meeting, final, now, C.coverWord)].concat(status, plan));
  });
}

/**
 * 全サブインダストリーの状況・計画進捗（FY26〜28売上）（2026/11/12）：今年度の売上の状況・計画進捗の2枚のあとのページ。
 * 横棒は今年度の売上（受注月が今年度の案件）を確度 A〜D（compact.outlookRanks）で積み上げ（棒の中に値。右に合計）、計画の位置に点線（compact.planLine）。
 * 右の列：計画・進捗率（A＋B）・確定（A＋B）・見込（C）（FY26）｜売上（FY27）・売上（FY28）（受注月・確度によらない）｜全体合計（FY26 の A〜D ＋ FY27 ＋ FY28）。
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
  // 来年度以降（FY27・FY28）は受注月・確度によらない売上（受注月・営業ステータスは今年度の受注を見込んだ列のため。2026/11/12）
  const lead = fy(year) + ' の' + sureWord + 'は ' + mtgAmount_(sumOf(total.p, year, sure)) + '百万円（計画比 ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') +
    '）、' + range + ' の合計は ' + mtgAmount_(sumOf(total.p, year, keys)) + '百万円。' + fy(next) + ' ' + mtgAmount_(laterOf(total.p, next)) + '・' + fy(last) + ' ' +
    mtgAmount_(laterOf(total.p, last)) + '百万円（失注除く）。\n' +
    '棒は ' + fy(year) + '売上（' + mtgPlanMonthName_() + 'が ' + fy(year) + '）を ' + range + ' で積み上げ。' + fy(next) + '・' + fy(last) + ' は' + mtgPlanMonthName_() +
    '・確度によらない。';
  // 全体合計の数え方は出典の行に（リード文を2行に収めるため）
  const sumNote = '※ 全体合計 ＝ ' + fy(year) + '（' + range + '）＋' + fy(next) + '＋' + fy(last);
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 480, 20, fy(year) + ' の売上の確度別（' + keys[0] + '〜' + keys[keys.length - 1] + '）と計画（百万円）',
    {size: 10.5, bold: true, color: c.ink})];
  const nameW = 96;
  // 列：今年度（計画・進捗率・確定（A＋B）・見込（C））｜来年度・再来年度の売上（受注月・確度によらない）｜全体合計
  const cols = [{key: 'plan', label: '計画', period: '（' + fy(year) + '）', w: 40}, {key: 'rate', label: '進捗率（A＋B）', period: '（' + fy(year) + '）', w: 56},
    {key: 'ab0', label: sureWord, period: '（' + fy(year) + '）', w: 56}, {key: 'c0', label: '見込（C）', period: '（' + fy(year) + '）', w: 50},
    {key: 'y1', label: '売上', period: '（' + fy(next) + '）', w: 48}, {key: 'y2', label: '売上', period: '（' + fy(last) + '）', w: 48},
    {key: 'all', label: '全体合計', period: '（' + fy(year) + '〜' + last + '）', w: 56}];
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
  const light = {C: true, D: true};   // 薄い色の確度は、棒の中の値を濃い緑の字に
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
      c0: amount(sumOf(p, year, ['C'])), y1: amount(laterOf(p, next)), y2: amount(laterOf(p, last)), all: amount(p.all || 0)};
    cols.forEach(col => {
      const text = values[col.key];
      const gap = text === G;
      const strong = col.key === 'rate' || col.key === 'ab0' || col.key === 'all';
      el.push(mtgBox_(col.x - S.padX / 2, ry, col.w + S.padX, rowH, text,
        {size: gap ? 7.5 : col.key === 'rate' ? 9 : 8, bold: !gap && (!!item.total || strong),
          color: gap ? c.muted : col.key === 'rate' && rate != null && rate >= 100 ? c.greenDark : c.ink, align: 'END', role: 'compact-value'}));
    });
  });
  return {title: MTG_TITLES.compactOverview + '（' + fy(year) + '〜' + last + '売上）', lead, elements: el,
    note: sumNote + (marked ? '。※ のある合計は「' + G + '」のシートを除く' : '')};
}

/**
 * 計画に対する積み上げ（FY26〜28売上）（2026/11/12）：左に今年度の売上を受注月の四半期（Q1〜Q4）ごとに確度別に積み上げ（前の四半期の上に続ける。
 * 上に「+増えた分」）、今年度の合計・計画（白抜き）・不足（計画との差。グレーの点線の枠。計画を超えたら「超過」）。区切り線の右に、来年度・再来年度の
 * 半期ごとの売上（受注月・確度によらない。年度ごとに色味を変え、前半期・後半期は同じ色味の濃淡（compact.laterColors）。summary.salesOutlook.halves。確定（A＋B）のページも同じ値）を、今年度の合計の上に
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
  // 確度で分けず、年度・半期ごとの色（compact.laterColors：年度ごとに色味を変え、前半期は濃く・後半期は薄く）。凡例は「FY27前」など
  const colors = C.laterColors || [['#2E75B6', '#9DC3E6'], ['#6F3FA6', '#C0A3E0']];
  const outlooks = summaries.map(sm => sm.salesOutlook).filter(x => x && !x.gap && x.halves);
  const later = [yy(1), yy(2)].reduce((list, y, k) => list.concat(['前', '後'].map((h, n) => {
    const key = 'FY' + y + h;
    const total = outlooks.reduce((t, x) => t + ((x.halves.find(v => v.year === y && v.half === h) || {}).value || 0), 0);
    return {label: 'FY' + y, sub: h + '半期', year: y, key, color: (colors[k] || [])[n] || '#9DC3E6', values: total ? {[key]: total} : {}, total};
  })), []);
  const laterRanks = later.map(l => ({key: l.key, color: l.color}));
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
  // No は3桁が 7.5pt で余裕をもって1行に入る幅（26pt・8pt では、Google スライドで「39／2」と折り返した。2026/11/09）
  const columns = mtgColumns_([30, 120, 130, 80, 66, 130, 112], 5);
  // 金額の列：今年度の売上（salesPage。今の値。前回の売上は記録に無いため）か、想定売上（前回▶今回）
  const bySales = !!o.compact.salesPage;
  const year = mtgFiscalYearKey_(s.now ? new Date(s.now) : new Date());
  const salesOf = ch => (ch.deal && ch.deal.sales ? ch.deal.sales[year] : null);
  const head = mtgHeadRow_(columns, ['No', '得意先／案件名', 'ステータス\n' + word + '▶今回', bySales ? 'FY' + year + '売上\n（今回）' : '想定売上\n' + word + '▶今回',
    '受注月\n' + word + '▶今回', '活動状況（最新）', '活動における課題']);
  const lack = s.gaps || {};
  const noSize = text => {
    let size = 7.5;
    while (mtgTextWidth_(text, size) > (columns[0] - S.cellPadX) * 0.8 && size > 5.5) size -= 0.5;
    return size;
  };
  // 並び：今回のステータスの確度 A → B → … → 創注 → 失注・保留・削除（各サブインダストリーの資料と同じまとまり）、その中は今年度の売上の大きい順
  // （salesPage。想定売上のときは各サブインダストリーの資料と同じ想定売上の大きい順）
  const groups = s.statusGroups || mtgStatusGroups_([]);
  const groupIndex = ch => groups.findIndex(gr => gr.key === mtgStatusGroupOf_(ch));
  const moves = s.changes.filter(ch => ch.category !== 'updated').map((ch, i) => ({ch, i}))
    .sort((a, b) => (groupIndex(a.ch) - groupIndex(b.ch)) || (bySales ? (salesOf(b.ch) || 0) - (salesOf(a.ch) || 0) : 0) || (a.i - b.i)).map(x => x.ch);
  const gone = moves.filter(ch => ch.removed || !ch.deal || mtgExcludedStatus_(ch.statusTo)).length;
  const lead = lack.changes ? mtgChangedLabel_(s) + 'は' + o.gapText + '（' + lack.changes + '）。'
    // 失注・削除の案件も載せる。状況・計画進捗の「変化のある案件」は今ある案件だけを数える（mtgCompactChangedNow_）ので、その分を書き添える
    : moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + (gone ? '（うち失注・削除 ' + gone + '件。状況・計画進捗の件数には含めない）' : '') +
      mtgUpdatedNote_(s) + '。\n' +
      'ステータス・期待値・受注月・活動状況が変わった案件。確度 A から順に、同じ確度は' + (bySales ? 'FY' + year + '売上' : '想定売上') + 'の大きい順。No は資料の通し番号。'
      : mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし' + mtgUpdatedNote_(s) + '。';
  if (!moves.length) {
    return {slides: [{title, lead, elements: mtgEmptyTable_(columns, head, lack.changes ? 'gap' : 'none', lack.changes || '')}], count: 0};
  }
  const rows = moves.map((ch, i) => {
    const status = mtgStatusMove_(ch);
    return mtgRow_(columns, [
      // No：3桁（「392」）でも折り返さないように、小さめの字（7.5pt）で1行。セルの幅の 8割に入らなければさらに小さくする
      // （見積もりぎりぎりだと、Google スライドでは折り返すことがあるため）
      {text: String(start + i + 1), size: noSize(String(start + i + 1)), lines: 1, align: 'CENTER', fill: c.greenSoft, color: c.ink, bold: true},
      mtgDealCell_(ch, columns[1]),
      // ステータスの動きが1行に入らなければ「▶」の前で改行（「B-／クロージング中」のように名前の途中で切らない）
      {text: mtgWrap_(status, columns[2], S.bodySize, 2).length > 1 ? status.replace(' ▶ ', '\n▶ ') : status, lines: 3, fill: c.white},
      // 今年度の売上（今の値。空なら「未入力」、削除した行は「（削除）」）か、想定売上（前回▶今回。変わっていなければ今の値）
      bySales ? {text: ch.removed ? '（削除）' : salesOf(ch) != null ? mtgAmount_(salesOf(ch)) : o.blankText, lines: 1, fill: c.white, align: 'CENTER',
        color: salesOf(ch) != null ? c.body : c.muted}
        : mtgFieldCell_(s, 'amount', mtgValueMove_(ch.base ? ch.base.amount : null, ch.deal ? ch.amount : null, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'activity', ch.activity, {lines: 3, fill: c.white, fit: true, grow: true}),
      mtgFieldCell_(s, 'issue', ch.issue, {lines: 3, fill: c.white, fit: true, grow: true})
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
  const job = {kind: 'compact', final: true, meeting, now: Date.now(), names: mtgTrackedNames_(ss), point: point || null, items: [],
    page: mtgCompactFrontCount_(part), no: 0, deck: null, result: null, folderUrl: '', sourceError: '', overwrite: !!overwrite,
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
        [part === 'changes' ? '表紙と PDF' : '表紙・状況と計画進捗・計画のページと PDF・会議のフォルダへの格納']);
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
