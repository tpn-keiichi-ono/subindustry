/**
 * 統合版（報告用）：全サブインダストリーの資料を、数字の要点と変化のあった案件だけにまとめた資料（RX Business Meeting）。2026/11/09 から。
 * 統合版（createBusinessMtgCombinedSummary()。目次・全サブインダストリーの状況・計画・各サブインダストリーの資料のページ）とは別の資料で、統合版は変えない。
 *
 * ページの並び：
 *   1. 表紙（統合版と同じひな形の表紙。見出しの下は「事前資料」ではなく「報告資料」（compact.coverWord）。mtgCombinedCover_）
 *   2. 全サブインダストリーの状況・計画進捗（mtgCompactOverviewSlide_）：サブインダストリーごとに、今年度の確定（確度 A＋B）の期待値の横棒（A・B を積み上げ、
 *      データラベル付き）と計画線（透明度のあるグレーの点線。上に「計画」）。右に「今年度の計画進捗」（計画・進捗率（A＋B）・確定（A＋B）/件・見込（C）/件）と
 *      「案件の状況（失注除く）」（期待値合計・前回比（比べた日付。増は青・減は赤）・案件数・変化のある案件）。進捗率（A＋B）の高い順、いちばん下に合計
 *   2'. 同じ（見込（C）を含む。棒の右に見込（C）を点線の枠で足す。mtgCompactOverviewSlide_(…, {withC: true})）
 *   3. 計画に対する積み上げ（全体。統合版と同じ mtgPlanSlide_）と、確定（A＋B）だけの積み上げ（mtgPlanRanksOnly_。凡例・見出し・リード文も確定の言葉）
 *   4〜 変化のあった案件：{サブインダストリー}（mtgCompactChangeSlides_。シートの順にサブインダストリーごと）：No.（資料の通し番号）・得意先／案件名・
 *      ステータス・期待値・受注月（前回▶今回）・活動状況・活動における課題。変化が無ければ「該当なし」の1ページ
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

/** 「統合版（報告用）（全サブインダストリー）」（知らせの名前） */
function mtgCompactLabel_() {
  return MTG_OPTIONS.compact.label + '（' + MTG_OPTIONS.combined.name + '）';
}

/** 変化のあった案件のページの前に入るページの数（表紙・全サブインダストリーの状況・計画進捗（見込（C）を含むページも）（・計画に対する積み上げ 2枚）） */
function mtgCompactFrontCount_() {
  return MTG_OPTIONS.plan.enabled ? 5 : 3;
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

/** ファイル名（統合版の名前の後ろに fileSuffix。「20261113_全サブインダストリー_報告用」） */
function mtgCompactFileName_(meeting) {
  return mtgFileName_(MTG_OPTIONS.combined.name, meeting, true) + MTG_OPTIONS.compact.fileSuffix;
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
function mtgCompactFront_(summaries, meeting, final, now) {
  const o = MTG_OPTIONS;
  const C = o.compact;
  return mtgWithAmountCeil_(() => {
    const merged = o.plan.enabled ? mtgMergePlans_(summaries.map(s => s.plan)) : null;
    // 計画に対する積み上げ：全確度（統合版と同じ）のあとに、確定（A＋B）だけ
    const plan = merged ? [mtgPlanSlide_(merged, now, o.combined.name),
      mtgPlanSlide_(mtgPlanRanksOnly_(merged, C.sureRanks), now, o.combined.name, {only: C.sureRanks, label: C.sureLabel})] : [];
    return mtgNumberSlides_(mtgCombinedFrame_(summaries, meeting, final),
      [mtgCombinedCover_(summaries, meeting, final, now, C.coverWord), mtgCompactOverviewSlide_(summaries),
        mtgCompactOverviewSlide_(summaries, {withC: true})].concat(plan));
  });
}

/**
 * 全サブインダストリーの状況・計画進捗（統合版の「計画に対する積み上げ（サブインダストリー別）」の図に、「全サブインダストリーの状況」の数字をまとめたもの）。
 * 横棒は確定（確度 A・B）だけを積み上げる（棒の中に値。入らなければ棒の右に「A 36」）。計画線は計画の位置に、透明度のあるグレーの点線（上に「計画」）。
 * 目盛りは、いちばん大きい計画の planScale 倍（か、それより大きい確定）を棒の幅いっぱいにする。進捗率（A＋B）＝ 確定（A＋B）÷ 計画。
 * 確定・見込・進捗率は今年度の積み上げ（plan.monthColumn の月が今年度の案件。mtgPlanOf_）、期待値合計・前回比・案件数・変化のある案件は失注を除くすべて
 * （統合版の状況のページと同じ）。期待値・月の列が無いシートは、棒の場所に「データなし」と理由。
 * options.withC：次のページ（見込（C）を含む）。棒の確定（A・B）の右に見込（C）を点線の枠で足す（並び・右の数字は同じ。目盛りは A＋B＋C も入るように）
 */
function mtgCompactOverviewSlide_(summaries, options) {
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
    : common === '前回比' && commonOf ? '前回比\n（' + mtgShortDate_(commonOf.previous.date) + '）' : common;
  const sum = (p, keys) => keys.reduce((t, k) => t + (p.byRank[k] || 0), 0);
  const cnt = (p, keys) => keys.reduce((t, k) => t + ((p.byRankCount || {})[k] || 0), 0);
  const rate = p => (p.gap || !(p.plan > 0) ? null : sum(p, ['A', 'B']) / p.plan * 100);
  const order = p => (rate(p) == null ? -1 : rate(p));
  const items = summaries.map(s => ({s, p: s.plan, lack: s.gaps || {}}))
    .sort((a, b) => (order(b.p) - order(a.p)) || (sum(b.p, ['A', 'B']) - sum(a.p, ['A', 'B'])));
  const ok = items.filter(x => !x.p.gap);
  const merged = mtgMergePlans_(ok.map(x => x.p));
  const first = mtgMonthOf_(merged.start);
  const last = mtgMonthOf_(merged.start + 11);
  const period = String(first.y).slice(2) + '/' + first.m + '〜' + String(last.y).slice(2) + '/' + last.m;
  // 合計（作れたシートの分だけ。作れないシートがあれば ※）
  const total = {active: 0, weighted: 0, delta: 0, changed: 0, lack: {}};
  items.forEach(({s, lack}) => {
    const b = s.bridge;
    total.active += b.endCount;
    if (lack.weighted) total.lack.weighted = true;
    else total.weighted += b.end;
    if (lack.weighted || lack.changes) total.lack.delta = true;
    else total.delta += b.end - b.start;
    if (lack.changes) total.lack.changed = true;
    else total.changed += s.kpi.changed;
  });
  const reached = ok.filter(x => rate(x.p) >= 100).length;
  const totalRate = ok.length ? rate(merged) : null;
  const lead = '進捗率（A＋B）は ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') + '（計画に届いているのは ' + ok.length + 'のうち ' + reached + '）。' +
    '案件 ' + total.active + '件・期待値合計 ' + mtgAmount_(total.weighted) + '百万円（失注除く）。\n' +
    '進捗率（A＋B）の高い順。進捗率・確定（A＋B）・見込（C）は今年度（' + period + '）の' + mtgPlanMonthName_() + 'の案件の期待値（進捗率 ＝ A＋B ÷ 計画）。';
  // 見込（C）を含むページのリード文：確定＋見込の計画比
  const withCRate = ok.length && merged.plan > 0 ? Math.round(sum(merged, ['A', 'B', 'C']) / merged.plan * 100) : null;
  const leadC = '確定（A＋B）に見込（C）を足すと ' + mtgAmount_(sum(merged, ['A', 'B', 'C'])) + '百万円で、計画の ' + (withCRate == null ? '—' : withCRate + '%') +
    '（進捗率（A＋B）は ' + (totalRate == null ? '—' : Math.round(totalRate) + '%') + '）。\n' +
    '見込（C）は点線の枠。並び・右の数字は前のページと同じ（進捗率（A＋B）の高い順）。';
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 480, 20, '今年度（' + period + '）の確定（A＋B）' + (withC ? 'と見込（C）' : '') + 'の期待値と計画（百万円）',
    {size: 10.5, bold: true, color: c.ink})];
  // 列：名前・横棒｜今年度の計画進捗（計画・進捗率（A＋B）・確定（A＋B）/件・見込（C）/件）｜案件の状況（期待値合計・前回比・案件数・変化のある案件）
  const nameW = 96;
  const groups = [
    {label: '今年度の計画進捗', cols: [{key: 'plan', label: '計画', w: 36}, {key: 'rate', label: '進捗率\n（A＋B）', w: 44},
      {key: 'ab', label: '確定（A＋B）\n/件', w: 66}, {key: 'c', label: '見込（C）\n/件', w: 60}]},
    {label: '案件の状況（失注除く）', cols: [{key: 'weighted', label: '期待値\n合計', w: 46}, {key: 'delta', label: deltaLabel, w: 46},
      {key: 'active', label: '案件数', w: 34}, {key: 'changed', label: '変化の\nある案件', w: 42}]}
  ];
  const colsW = groups.reduce((t, gr) => t + gr.cols.reduce((u, col) => u + col.w, 0), 0) + 8;
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  // まとまりの見出し（「今年度の計画進捗」など）と下線：上（グラフの見出しとの間）と下（列の見出しとの間）に少し余白
  const groupY = y + 23;
  const headY = groupY + 19;
  const headH = 20;
  const rows = items.concat([{total: true, s: {name: '合計'}, p: merged, lack: total.lack}]);
  const rowH = Math.min(22, (S.bottom - headY - headH - 2) / rows.length);
  // 目盛り：いちばん大きい計画が棒の幅の 1/planScale ほどの位置に来るように（計画線の右に、計画を超えた分を描ける）
  const max = Math.max.apply(null, [1].concat(ok.map(item => Math.max(sum(item.p, ['A', 'B']), item.p.plan * C.planScale))).concat(ok.map(item => sum(item.p, barKeys))));
  const scale = barW / max;
  let cx = barX + barW + 10;
  groups.forEach((gr, gi) => {
    if (gi) cx += 8;
    const gw = gr.cols.reduce((u, col) => u + col.w, 0);
    el.push(mtgBox_(cx - S.padX / 2 + 4, groupY, gw + S.padX - 4, 12, gr.label, {size: 8, bold: true, color: c.greenDark, align: 'CENTER', role: 'compact-group'}));
    el.push({type: 'box', x: cx + 4, y: groupY + 14, w: gw - 4, h: 0.75, fill: c.greenDark, shape: 'RECTANGLE'});
    gr.cols.forEach(col => {
      el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, headH, col.label, {size: 7.5, color: c.sub, align: 'CENTER', role: 'compact-head'}));
      col.x = cx;
      cx += col.w;
    });
  });
  const cols = [].concat.apply([], groups.map(gr => gr.cols));
  // 凡例：グラフの上（列の見出しの行の、棒の左端から）に確度 A・B の色。計画線は線の上に見出し（「計画」）を付けるので凡例に入れない
  let lx = barX;
  ranks.forEach(r => {
    el.push(r.dashed ? {type: 'box', x: lx, y: headY + headH / 2 - 3.5, w: 7, h: 7, fill: c.white, line: r.color, lineWeight: 1, dash: true,
      shape: 'RECTANGLE', role: 'legend'}
      : {type: 'box', x: lx, y: headY + headH / 2 - 3.5, w: 7, h: 7, fill: r.color, shape: 'RECTANGLE', role: 'legend'});
    const bw = mtgTextWidth_(r.key, 8.5) + S.padX + 2;
    el.push(mtgBox_(lx + 10 - S.padX / 2, headY, bw, headH, r.key, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 10 + bw - S.padX + 10;
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
    el.push({type: 'line', x: px, y: groupY + 14, w: 0, h: headY + headH + 3 - groupY - 14, color: line.color, alpha: line.alpha, weight: line.weight,
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
      const outside = [];
      ranks.forEach(r => {
        const v = p.byRank[r.key] || 0;
        if (v <= 0) return;
        const bx = barX + level * scale;
        const bw = Math.max(v * scale, 0.75);
        // 見込（C）は白地に確度の色の点線の枠（確定と見分ける）
        el.push(r.dashed ? {type: 'box', x: bx, y: ry + (rowH - bh) / 2, w: bw, h: bh, fill: c.white, line: r.color, lineWeight: 1, dash: true,
          shape: 'RECTANGLE', role: 'compact-bar'}
          : {type: 'box', x: bx, y: ry + (rowH - bh) / 2, w: bw, h: bh, fill: r.color, line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role: 'compact-bar'});
        // データラベル：棒の中（白い字。見込は緑の字）。入らなければ棒の右に、確度を付けて確度の色の字で（「A 36」）
        const label = mtgAmount_(Math.round(v * 10) / 10);
        if (mtgTextWidth_(label, 7.5) + 4 <= bw) {
          el.push(mtgBox_(bx - S.padX / 2, ry, bw + S.padX, rowH, label, {size: 7.5, bold: true, color: r.dashed ? c.greenDark : c.white, align: 'CENTER',
            role: 'compact-bar-label'}));
        } else {
          outside.push({label: r.key + ' ' + label, color: r.dashed ? c.greenDark : r.color});
        }
        level += v;
      });
      let ox = barX + level * scale + 3;
      outside.forEach(o2 => {
        const tw = mtgTextWidth_(o2.label, 7.5);
        el.push(mtgBox_(ox - S.padX / 2, ry, tw + S.padX + 2, rowH, o2.label, {size: 7.5, bold: true, color: o2.color, role: 'compact-bar-label'}));
        ox += tw + 5;
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
      weighted: items.every(z => z.lack.weighted) ? G : mtgAmount_(total.weighted) + mark('weighted'),
      delta: items.every(z => z.lack.weighted || z.lack.changes) ? G : mtgSigned_(total.delta) + mark('delta'),
      active: total.active + '件', changed: items.every(z => z.lack.changes) ? G : total.changed + '件' + mark('changed')
    } : {
      plan: p.gap ? G : mtgAmount_(p.plan), ab: amountCount(['A', 'B']), c: amountCount(['C']),
      weighted: lack.weighted ? G : mtgAmount_(b.end), delta: lack.weighted || lack.changes ? G : mtgSigned_(b.end - b.start),
      active: b.endCount + '件', changed: lack.changes ? G : s.kpi.changed + '件'
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
  return Object.assign({title: MTG_TITLES.compactOverview + (withC ? '（見込（C）を含む）' : ''), lead: withC ? leadC : lead, elements: el}, marked ? {note: '※ は「' + G + '」のシートを除いた合計'} : {});
}

/**
 * 変化のあった案件：{サブインダストリー}（1つのサブインダストリーのページ。行が入らなければ「（続き）」のページ）。変化のある案件（その他の項目のみの更新を除く。
 * 各サブインダストリーの資料の「前回から変化のある案件」と同じ案件。期待値の大きい順）を、No.（start + 1 からの資料の通し番号）とともに全件。
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
  const columns = mtgColumns_([26, 120, 130, 80, 66, 130, 116], 5);   // No. は3桁が 8pt で1行に入る幅
  const head = mtgHeadRow_(columns, ['No.', '得意先／案件名', 'ステータス\n' + word + '▶今回', '期待値\n' + word + '▶今回',
    '受注月\n' + word + '▶今回', '活動状況（最新）', '活動における課題']);
  const lack = s.gaps || {};
  // 並びは期待値（今の値）の大きい順。同じなら想定売上の大きい順、それも同じならシートの資料と同じ順（今回のステータスの順）。
  // 期待値が空・削除した案件は後ろ
  const value = v => (v == null ? -Infinity : v);
  const moves = s.changes.filter(ch => ch.category !== 'updated').map((ch, i) => ({ch, i}))
    .sort((a, b) => (value(b.ch.weightedTo) - value(a.ch.weightedTo)) || (value(b.ch.amount) - value(a.ch.amount)) || (a.i - b.i)).map(x => x.ch);
  const lead = lack.changes ? mtgChangedLabel_(s) + 'は' + o.gapText + '（' + lack.changes + '）。'
    : moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + mtgUpdatedNote_(s) + '。\n' +
      'ステータス・期待値・受注月・活動状況が変わった案件（期待値の大きい順）。No. は資料の通し番号。'
      : mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし' + mtgUpdatedNote_(s) + '。';
  if (!moves.length) {
    return {slides: [{title, lead, elements: mtgEmptyTable_(columns, head, lack.changes ? 'gap' : 'none', lack.changes || '')}], count: 0};
  }
  const rows = moves.map((ch, i) => {
    const status = mtgStatusMove_(ch);
    return mtgRow_(columns, [
      // No.：3桁（「380」）でも折り返さないように、小さめの字（8pt）で1行。入らなければさらに小さくする
      {text: String(start + i + 1), size: 8, lines: 1, fit: true, minSize: 6, align: 'CENTER', fill: c.greenSoft, color: c.ink, bold: true},
      mtgDealCell_(ch, columns[1]),
      // ステータスの動きが1行に入らなければ「▶」の前で改行（「B-／クロージング中」のように名前の途中で切らない）
      {text: mtgWrap_(status, columns[2], S.bodySize, 2).length > 1 ? status.replace(' ▶ ', '\n▶ ') : status, lines: 3, fill: c.white},
      mtgFieldCell_(s, 'weighted', mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
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
  return {slides: pages.map((p, n) => ({title: title + (n ? '（続き）' : ''), lead, elements: [{type: 'table', x: S.x, y: S.top, columns, rows: p.rows}]})),
    count: moves.length};
}

/* ---------------- 作る（区切りながら） ---------------- */

/** 統合版（報告用）を作れるか（作れなければ {error: 理由}、作れれば {text: 確かめの文}） */
function mtgCompactCheck_(meeting, point) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const o = MTG_OPTIONS;
  if (!o.compact.enabled) return {error: '統合版（報告用）を使わない設定です（MeetingSummary.gs の MTG_OPTIONS.compact.enabled）。'};
  const pending = mtgJobLoad_();
  if (pending) {
    return {error: (mtgDeckJob_(pending) ? mtgDeckJob_(pending).label : 'サブインダストリーの資料') + 'を作っている途中です（自動で続けています）。' +
      '終わってから（showBusinessMtgSummariesResult() で確かめてから）実行してください。'};
  }
  const names = mtgTrackedNames_(ss);
  if (!names.length) return {error: '追跡しているシート（test.gs の DIFF_RULES）が見つかりません。'};
  return {text: '「' + o.combined.name + '」の' + o.compact.label + '（' + names.length + 'サブインダストリーの状況・計画進捗と、変化のあった案件）を、' +
    '今のシートの値で作ります。\n\n' +
    '会議日 ' + mtgDateLabel_(meeting) + ' の資料として作り、会議のフォルダに格納します（会議日の記録・時点コピー・インデックスは変えません）。\n' +
    '比べる時点：' + (point ? mtgChosenText_(point) + '（全シート共通）' : '前回の会議（シートごと）')};
}

/**
 * 統合版（報告用）を作り始める（画面の「作成する」から）。時間がかかれば区切って続ける（mtgCompactRun_。続きは continueBusinessMtgSummaries）。
 * 戻り値は終わり（区切ったとき）の知らせ
 */
function mtgCompactStart_(meeting, point, overwrite) {
  const started = mtgNow_();   // 6分の上限は、この実行を始めたときから
  const check = mtgCompactCheck_(meeting, point);
  if (check.error) throw new Error(check.error);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const job = {kind: 'compact', final: true, meeting, now: Date.now(), names: mtgTrackedNames_(ss), point: point || null, items: [],
    page: mtgCompactFrontCount_(), no: 0, deck: null, result: null, folderUrl: '', sourceError: '', overwrite: !!overwrite};
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
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  job.folderUrl = folder.getUrl();
  const fileName = mtgCompactFileName_(job.meeting);
  if (!job.deck) {
    job.deck = mtgDeckStart_(fileName, folder);
    mtgJobCheckpoint_(job);
  }
  while (job.items.length < job.names.length) {
    if (!fits(R.summaryMs + R.deckMs)) return mtgJobPause_(job);
    const batch = R.readMode === 'api' ? job.names.slice(job.items.length, job.items.length + Math.max(R.sheetsPerRead || 1, 1))
      : job.names.slice(job.items.length);
    const source = mtgJobSource_(ss, job, batch);
    if (!fits(R.summaryMs + R.deckMs)) return mtgJobPause_(job);
    const summaries = mtgSummaries_(source, batch, new Date(job.now), true, job.meeting, job.point).summaries;
    for (let i = 0; i < summaries.length; i++) {
      if (i > 0 && !fits(R.deckMs)) return mtgJobPause_(job);
      const summary = summaries[i];
      const timer = mschTimer_(o.compact.label + 'のスライド（' + summary.name + '）', true);
      const part = mtgCompactChangeSlides_(summary, job.no);
      const slides = mtgNumberSlides_(mtgCombinedFrame_([summary], job.meeting, true), part.slides, job.page);
      mtgDeckAdd_(job.deck, slides, job.page);
      timer.lap('スライドを足す（' + slides.length + '枚）');
      job.items.push({name: summary.name, start: job.page + 1, pages: slides.length, lite: mtgCombinedLite_(summary)});
      job.page += slides.length;
      job.no += part.count;
      mtgJobCheckpoint_(job);   // 資料ごとに進み具合を保存（実行が止められても、この資料の次から。mtgCombinedRecover_）
    }
  }
  if (!job.result) {
    if (!fits(R.combinedFinishMs)) return mtgJobPause_(job);
    const timer = mschTimer_(o.compact.label + 'の仕上げ', true);
    job.deck.finishing = true;   // 仕上げの途中で止められたら、続きからは作れない（作り直す）
    mtgJobCheckpoint_(job);
    mtgDeckAdd_(job.deck, mtgCompactFront_(job.items.map(item => item.lite), job.meeting, true, new Date(job.now)), 0);
    timer.lap('表紙・状況と計画進捗・計画のページ');
    const existing = job.overwrite ? mtgExistingPdf_(fileName, o.combined.name, job.meeting, true, folder) : null;
    const out = mtgDeckFinish_(job.deck, fileName, folder, existing);
    timer.lap('並べて PDF に保存');
    const result = {name: o.combined.name, file: out.pdf, pages: job.page, overwritten: !!out.overwritten};
    // 会議のフォルダに格納する（できなくても止めない）。インデックスの統合版の列には書かない
    const stored = mtgStoreToMeetingFolder_(o.combined.name, job.meeting, out.pdf);
    if (stored.error) result.storeError = stored.error;
    else result.stored = stored;
    job.result = mtgCombinedPlain_(result);
    job.deck = null;
  }
  return true;
}

/** 統合版（報告用）の仕事の、終わり（done）・区切ったときの知らせ */
function mtgCompactMessage_(job, done) {
  const label = mtgCompactLabel_();
  if (!done) {
    const rest = job.names.slice(job.items.length);
    const left = [].concat(rest.length ? ['変化のあった案件のページ ' + rest.length + '件（' + rest.join('・') + '）'] : [],
      ['表紙・状況と計画進捗・計画のページと PDF・会議のフォルダへの格納']);
    return ['時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りは約' + Math.round(MTG_OPTIONS.run.continueAfterMs / 60000) +
      '分後に自動で続けます：' + left.join('・') + '。',
    '続きが終わったら、エディタで showBusinessMtgSummariesResult() を実行すると結果が出ます（実行ログにも出ます）。終わるまで、会議サマリの関数は実行しないでください。',
    '', label + '：' + job.items.length + '/' + job.names.length + 'サブインダストリーのページを作りました。',
    '', '保存先：' + job.folderUrl].join('\n');
  }
  const r = job.result;
  return ['「' + label + '」：' + r.fileName + '（' + r.pages + 'ページ）\n　' + r.fileUrl +
    (r.overwritten ? '（前の資料に上書き。リンクはそのまま）' : '') +
    (r.folderName ? '\n　会議のフォルダ「' + r.folderName + '」' + (r.created ? 'を作って' : 'に') + '格納しました：' + r.folderUrl : '') +
    (r.storeError ? '\n　※ 会議のフォルダに格納できませんでした（' + r.storeError + '）。資料は保存先にあります。' : ''),
  '', '保存先：' + job.folderUrl].concat(mtgSourceNote_(job)).join('\n');
}
