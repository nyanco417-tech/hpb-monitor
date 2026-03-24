/**
 * HPB競合サロン 空き状況モニター（GitHub Actions版）
 *
 * Full Moon / Flower Factories / La.stella の予約カレンダーを
 * スクリーンショットで保存し、前日との差分を分析する
 * 結果をメールで通知する
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// 設定
const SALONS = [
  {
    name: 'fullmoon',
    label: 'Full Moon',
    storeId: 'H000742281',
    couponId: 'CP00000010923330',
  },
  {
    name: 'flower',
    label: 'Flower Factories',
    storeId: 'H000685940',
    couponId: 'CP00000009651621',
  },
  {
    name: 'lastella',
    label: 'La.stella',
    storeId: 'H000715681',
    couponId: 'CP00000010301268',
  },
];

const WEEKS_TO_ADVANCE = 5;

// GitHub Actions上ではリポジトリ内のdataディレクトリに保存
const OUTPUT_BASE = path.join(__dirname, 'data');

const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

function getDateStr() {
  // JSTで日付を取得（GitHub ActionsはUTCなので+9時間）
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const y = jst.getUTCFullYear();
  const m = String(jst.getUTCMonth() + 1).padStart(2, '0');
  const d = String(jst.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getPrevDateStr(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * サロンページから口コミ件数と評価点を取得（JSON-LD構造化データから）
 */
async function scrapeReviewData(browser, salon) {
  console.log(`\n--- ${salon.label} 口コミ取得 ---`);
  const page = await browser.newPage();
  try {
    const salonUrl = `https://beauty.hotpepper.jp/kr/sln${salon.storeId}/`;
    await page.goto(salonUrl, { waitUntil: 'networkidle', timeout: 30000 });

    const reviewData = await page.evaluate(() => {
      // JSON-LDから取得（最も安定）
      const scripts = document.querySelectorAll('script[type="application/ld+json"]');
      for (const script of scripts) {
        try {
          const json = JSON.parse(script.textContent);
          if (json.aggregateRating) {
            return {
              reviewCount: json.aggregateRating.reviewCount || 0,
              rating: json.aggregateRating.ratingValue || 0,
            };
          }
        } catch (e) {}
      }
      // フォールバック: HTMLから取得
      const countEl = document.querySelector('.slnHeaderKuchikomiCount');
      const ratingEl = document.querySelector('.slnHeaderKuchikomiPoint');
      if (countEl) {
        const match = countEl.textContent.match(/(\d+)/);
        return {
          reviewCount: match ? parseInt(match[1], 10) : 0,
          rating: ratingEl ? parseFloat(ratingEl.textContent.trim()) : 0,
        };
      }
      return { reviewCount: 0, rating: 0 };
    });

    console.log(`  ${salon.label}: 口コミ ${reviewData.reviewCount}件 / 評価 ${reviewData.rating}`);
    return reviewData;
  } catch (err) {
    console.error(`  [ERROR] 口コミ取得失敗 ${salon.label}: ${err.message}`);
    return { reviewCount: 0, rating: 0 };
  } finally {
    await page.close();
  }
}

async function expandCalendarTable(page) {
  await page.evaluate(() => {
    const tbl = document.querySelector('#jsRsvCdTbl');
    if (!tbl) return;
    const table = tbl.querySelector('.innerTable');
    const tableWidth = table ? table.scrollWidth + 2 : 3000;
    let el = tbl;
    while (el && el !== document.documentElement) {
      el.style.width = tableWidth + 'px';
      el.style.maxWidth = 'none';
      el.style.overflow = 'visible';
      el = el.parentElement;
    }
  });
}

async function scrapeCalendarData(page) {
  return await page.evaluate(() => {
    const tbl = document.querySelector('#jsRsvCdTbl');
    if (!tbl) return { dates: [], slots: {} };
    const table = tbl.querySelector('table.innerTable') || tbl.querySelector('table');
    if (!table) return { dates: [], slots: {} };

    const MONTHS = {
      Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
      Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
    };

    const dateRow = table.querySelector('tr.dayCellContainer');
    if (!dateRow) return { dates: [], slots: {} };
    const dateThs = dateRow.querySelectorAll('th');
    const dates = [];
    for (const th of dateThs) {
      const text = th.textContent.trim();
      const match = text.match(/(\w{3})\s+(\d{1,2})\s+\d{2}:\d{2}:\d{2}\s+\w+\s+(\d{4})/);
      if (match) {
        const monthNum = MONTHS[match[1]];
        const day = parseInt(match[2], 10);
        const year = parseInt(match[3], 10);
        if (monthNum) {
          dates.push(`${year}-${String(monthNum).padStart(2, '0')}-${String(day).padStart(2, '0')}`);
        } else {
          dates.push(null);
        }
      } else {
        dates.push(null);
      }
    }

    const tbody = table.querySelector('tbody');
    if (!tbody) return { dates: dates.filter(Boolean), slots: {} };
    const bodyThs = tbody.querySelectorAll(':scope > tr > th');
    if (bodyThs.length < 3) return { dates: dates.filter(Boolean), slots: {} };

    const timeTable = bodyThs[0].querySelector('table');
    if (!timeTable) return { dates: dates.filter(Boolean), slots: {} };
    const timeRows = timeTable.querySelectorAll('tr');
    const times = [];
    for (const tr of timeRows) {
      const th = tr.querySelector('th');
      if (th) {
        times.push(th.textContent.trim().replace('：', ':'));
      }
    }

    const slots = {};
    const dataThStart = 1;
    const dataThEnd = bodyThs.length - 1;
    for (let col = dataThStart; col < dataThEnd && (col - dataThStart) < dates.length; col++) {
      const dateStr = dates[col - dataThStart];
      if (!dateStr) continue;
      const colTable = bodyThs[col].querySelector('table');
      if (!colTable) continue;
      const colRows = colTable.querySelectorAll('tr');
      slots[dateStr] = {};
      for (let r = 0; r < colRows.length && r < times.length; r++) {
        const td = colRows[r].querySelector('td');
        if (!td) continue;
        const text = td.textContent.trim();
        let status;
        if (text === '◎' || text === '○') status = '◎';
        else if (text === '△') status = '△';
        else if (text === '×') status = '×';
        else if (text === '－' || text === '-' || text === 'ー') status = '-';
        else status = text || '-';
        slots[dateStr][times[r]] = status;
      }
    }
    return { dates: dates.filter(Boolean), slots };
  });
}

async function captureCalendar(page, salon, outputDir, weekIndex) {
  await expandCalendarTable(page);
  await page.waitForTimeout(300);
  const calendarEl = page.locator('#jsRsvCdTbl');
  const exists = await calendarEl.count();
  if (exists === 0) {
    console.log(`  [WARN] カレンダーが見つかりません`);
    const filename = `${salon.name}_week${weekIndex}.png`;
    await page.screenshot({ path: path.join(outputDir, filename), fullPage: true });
    return null;
  }
  const filename = `${salon.name}_week${weekIndex}.png`;
  await calendarEl.screenshot({ path: path.join(outputDir, filename) });
  console.log(`  スクリーンショット保存: ${filename}`);
  const data = await scrapeCalendarData(page);
  return data;
}

async function processSalon(browser, salon, outputDir) {
  console.log(`\n=== ${salon.label} (${salon.storeId}) ===`);
  const page = await browser.newPage();
  await page.setViewportSize({ width: 2200, height: 900 });
  const mergedSlots = {};

  try {
    const couponUrl = `https://beauty.hotpepper.jp/CSP/kr/reserve/?storeId=${salon.storeId}`;
    console.log(`  クーポンページにアクセス...`);
    await page.goto(couponUrl, { waitUntil: 'networkidle', timeout: 30000 });

    const couponLink = page.locator(`a[href*="couponId=${salon.couponId}"][href*="add=0"]`);
    const linkCount = await couponLink.count();
    if (linkCount === 0) {
      console.log(`  [WARN] 指定クーポンが見つかりません。最初のクーポンを使用...`);
      const firstLink = page.locator('a[href*="afterCoupon"][href*="add=0"]').first();
      if (await firstLink.count() === 0) {
        console.log(`  [ERROR] クーポンが見つかりません。スキップします。`);
        return;
      }
      await firstLink.click();
    } else {
      await couponLink.first().click();
    }

    await page.waitForLoadState('networkidle', { timeout: 30000 });
    console.log(`  日時選択ページ到達`);
    await page.waitForSelector('#jsRsvCdTbl', { timeout: 10000 });
    const data0 = await captureCalendar(page, salon, outputDir, 0);
    if (data0) mergeSlots(mergedSlots, data0.slots);

    for (let w = 1; w <= WEEKS_TO_ADVANCE; w++) {
      const nextLink = page.locator('a.arrowPagingWeekR.jscCalNavi');
      const nextExists = await nextLink.count();
      if (nextExists === 0) {
        console.log(`  [WARN] 「次の一週間」リンクが見つかりません（week ${w}）`);
        break;
      }
      const href = await nextLink.getAttribute('href');
      const nextUrl = href.startsWith('http') ? href : `https://beauty.hotpepper.jp${href}`;
      console.log(`  次の一週間へ (week ${w})...`);
      await sleep(2000);
      await page.goto(nextUrl, { waitUntil: 'networkidle', timeout: 30000 });
      await page.waitForSelector('#jsRsvCdTbl', { timeout: 10000 });
      const dataW = await captureCalendar(page, salon, outputDir, w);
      if (dataW) mergeSlots(mergedSlots, dataW.slots);
    }

    const jsonData = { salon: salon.name, capturedAt: getDateStr(), slots: mergedSlots };
    const jsonPath = path.join(outputDir, `${salon.name}_data.json`);
    fs.writeFileSync(jsonPath, JSON.stringify(jsonData, null, 2), 'utf-8');
    console.log(`  データ保存: ${salon.name}_data.json (${Object.keys(mergedSlots).length}日分)`);
  } catch (err) {
    console.error(`  [ERROR] ${salon.label}: ${err.message}`);
    const errFile = path.join(outputDir, `${salon.name}_error.png`);
    await page.screenshot({ path: errFile, fullPage: true }).catch(() => {});
  } finally {
    await page.close();
  }
}

function mergeSlots(target, source) {
  for (const [dateStr, timeSlots] of Object.entries(source)) {
    if (!target[dateStr]) target[dateStr] = {};
    for (const [time, status] of Object.entries(timeSlots)) {
      target[dateStr][time] = status;
    }
  }
}

function formatShortDate(d) {
  return `${d.getMonth() + 1}/${String(d.getDate()).padStart(2, '0')}`;
}

function formatDateWithDay(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const month = d.getMonth() + 1;
  const day = String(d.getDate()).padStart(2, '0');
  const dayName = DAY_NAMES[d.getDay()];
  return `${month}/${day}(${dayName})`;
}

function isWeekendDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.getDay() === 0 || d.getDay() === 6;
}

function sortByDateTime(a, b) {
  if (a.date !== b.date) return a.date.localeCompare(b.date);
  return a.time.localeCompare(b.time);
}

/**
 * 差分レポート生成（テキスト + HTML）
 */
function generateDiffReport(outputDir, dateStr) {
  const prevDateStr = getPrevDateStr(dateStr);
  const prevDir = path.join(OUTPUT_BASE, prevDateStr);

  if (!fs.existsSync(prevDir)) {
    console.log(`\n前日データ (${prevDateStr}) が見つかりません。差分レポートは生成しません。`);
    const report = `=== HPB空き状況 差分レポート（${dateStr}）===\n\n初回取得のため差分なし。明日以降、前日との比較が自動生成されます。\n`;
    fs.writeFileSync(path.join(outputDir, 'diff_report.txt'), report, 'utf-8');
    fs.writeFileSync(path.join(outputDir, 'diff_report.html'), generateHtmlReport(dateStr, null), 'utf-8');
    return { text: report, hasPrev: false };
  }

  let report = `=== HPB空き状況 差分レポート（${dateStr} vs ${prevDateStr}）===\n`;
  const salonReports = [];

  for (const salon of SALONS) {
    const todayFile = path.join(outputDir, `${salon.name}_data.json`);
    const prevFile = path.join(prevDir, `${salon.name}_data.json`);

    const salonReport = { name: salon.label, filled: [], opened: [], newDates: [], weeklyStats: {} };
    report += `\n■ ${salon.label}\n`;

    if (!fs.existsSync(todayFile)) {
      report += `  今日のデータがありません。\n`;
      salonReports.push(salonReport);
      continue;
    }
    if (!fs.existsSync(prevFile)) {
      report += `  前日のデータがありません。初回取得です。\n`;
      salonReports.push(salonReport);
      continue;
    }

    const todayData = JSON.parse(fs.readFileSync(todayFile, 'utf-8'));
    const prevData = JSON.parse(fs.readFileSync(prevFile, 'utf-8'));

    const filled = [];
    const opened = [];
    const newDates = [];
    const weeklyStats = {};

    for (const [date, timeSlots] of Object.entries(todayData.slots)) {
      const d = new Date(date + 'T00:00:00');
      const dayOfWeek = d.getDay();
      const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      const monday = new Date(d);
      monday.setDate(monday.getDate() + mondayOffset);
      const sunday = new Date(monday);
      sunday.setDate(sunday.getDate() + 6);
      const weekKey = formatShortDate(monday) + '-' + formatShortDate(sunday);

      if (!weeklyStats[weekKey]) {
        weeklyStats[weekKey] = { total: 0, available: 0, sortKey: monday.getTime() };
      }

      const prevTimeSlots = prevData.slots[date];
      if (!prevTimeSlots) newDates.push(date);

      for (const [time, status] of Object.entries(timeSlots)) {
        if (status === '-') continue;
        weeklyStats[weekKey].total++;
        if (status === '◎' || status === '△') weeklyStats[weekKey].available++;
        if (!prevTimeSlots) continue;
        const prevStatus = prevTimeSlots[time];
        if (!prevStatus) continue;
        if ((prevStatus === '◎' || prevStatus === '△') && status === '×') {
          filled.push({ date, time, from: prevStatus, to: status });
        } else if (prevStatus === '×' && (status === '◎' || status === '△')) {
          opened.push({ date, time, from: prevStatus, to: status });
        }
      }
    }

    report += `[新しく埋まった枠] ${filled.length}件\n`;
    for (const item of filled.sort(sortByDateTime)) {
      report += `  ${formatDateWithDay(item.date)} ${item.time}  ${item.from}→${item.to}\n`;
    }
    if (filled.length === 0) report += `  なし\n`;

    report += `\n[キャンセルで空いた枠] ${opened.length}件\n`;
    for (const item of opened.sort(sortByDateTime)) {
      const isWeekend = isWeekendDate(item.date);
      const note = isWeekend ? '  ← 土日の後出し開放？要注目' : '';
      report += `  ${formatDateWithDay(item.date)} ${item.time}  ${item.from}→${item.to}${note}\n`;
    }
    if (opened.length === 0) report += `  なし\n`;

    if (newDates.length > 0) {
      report += `\n[新規表示された日付] ${newDates.length}日\n`;
      for (const date of newDates.sort()) {
        const slots = todayData.slots[date];
        const total = Object.values(slots).filter(s => s !== '-').length;
        const avail = Object.values(slots).filter(s => s === '◎' || s === '△').length;
        report += `  ${formatDateWithDay(date)}  空き${avail}/${total}枠\n`;
      }
    }

    report += `\n[週別サマリー（月曜始まり）]\n`;
    report += `  期間          | 予約済 | 空き  | 合計  | 予約率\n`;
    report += `  --------------|--------|-------|-------|-------\n`;
    const sortedWeeks = Object.entries(weeklyStats).sort((a, b) => a[1].sortKey - b[1].sortKey);
    for (const [weekLabel, stats] of sortedWeeks) {
      if (stats.total === 0) continue;
      const booked = stats.total - stats.available;
      const bookedPct = ((booked / stats.total) * 100).toFixed(1);
      report += `  ${weekLabel.padEnd(14)}| ${String(booked).padStart(4)}枠 | ${String(stats.available).padStart(3)}枠 | ${String(stats.total).padStart(3)}枠 | ${bookedPct}%\n`;
    }

    salonReport.filled = filled;
    salonReport.opened = opened;
    salonReport.newDates = newDates;
    salonReport.weeklyStats = weeklyStats;
    salonReports.push(salonReport);
  }

  // 口コミ増減セクション
  report += generateReviewSection(outputDir, dateStr);

  fs.writeFileSync(path.join(outputDir, 'diff_report.txt'), report, 'utf-8');
  fs.writeFileSync(path.join(outputDir, 'diff_report.html'), generateHtmlReport(dateStr, salonReports), 'utf-8');
  console.log(`\n差分レポート保存: diff_report.txt / diff_report.html`);
  console.log(report);
  return { text: report, html: generateHtmlReport(dateStr, salonReports), hasPrev: true, salonReports };
}

/**
 * 口コミ増減レポートセクション生成
 * 今日のreviews.jsonと過去データを比較し、日次・週次の増減を出す
 */
function generateReviewSection(outputDir, dateStr) {
  const todayReviewsFile = path.join(outputDir, 'reviews.json');
  if (!fs.existsSync(todayReviewsFile)) return '';

  const todayReviews = JSON.parse(fs.readFileSync(todayReviewsFile, 'utf-8'));
  let section = `\n${'='.repeat(60)}\n`;
  section += `■ 口コミ状況（${dateStr}）\n\n`;

  // 今日の口コミ数一覧
  section += `[本日の口コミ数]\n`;
  for (const salon of SALONS) {
    const r = todayReviews.salons[salon.name];
    if (r) {
      section += `  ${salon.label.padEnd(20)} ${r.reviewCount}件  (${r.rating}点)\n`;
    }
  }

  // 前日との比較
  const prevDateStr = getPrevDateStr(dateStr);
  const prevReviewsFile = path.join(OUTPUT_BASE, prevDateStr, 'reviews.json');
  if (fs.existsSync(prevReviewsFile)) {
    const prevReviews = JSON.parse(fs.readFileSync(prevReviewsFile, 'utf-8'));
    section += `\n[前日比（${prevDateStr}→${dateStr}）]\n`;
    for (const salon of SALONS) {
      const today = todayReviews.salons[salon.name];
      const prev = prevReviews.salons[salon.name];
      if (today && prev) {
        const diff = today.reviewCount - prev.reviewCount;
        const sign = diff > 0 ? '+' : '';
        section += `  ${salon.label.padEnd(20)} ${prev.reviewCount}→${today.reviewCount}件 (${sign}${diff})\n`;
      }
    }
  }

  // 週ごとのサマリー（月曜始まり）: 過去のreviews.jsonを遡って集計
  section += `\n[週別口コミ増減サマリー]\n`;
  const weeklyReviews = collectWeeklyReviews();
  if (weeklyReviews.length > 0) {
    section += `  期間          |`;
    for (const salon of SALONS) {
      section += ` ${salon.label.substring(0, 10).padEnd(10)} |`;
    }
    section += `\n`;
    section += `  --------------|`;
    for (const salon of SALONS) {
      section += `------------|`;
    }
    section += `\n`;
    for (const week of weeklyReviews) {
      section += `  ${week.label.padEnd(14)}|`;
      for (const salon of SALONS) {
        const data = week.salons[salon.name];
        if (data) {
          const sign = data.diff > 0 ? '+' : '';
          section += ` ${(sign + data.diff + '件').padStart(10)} |`;
        } else {
          section += `        -   |`;
        }
      }
      section += `\n`;
    }
  } else {
    section += `  データ蓄積中です（1週間後から表示されます）\n`;
  }

  return section;
}

/**
 * 過去のreviews.jsonから週ごとの口コミ増減を集計
 */
function collectWeeklyReviews() {
  const dataDir = OUTPUT_BASE;
  if (!fs.existsSync(dataDir)) return [];

  // 日付ディレクトリを一覧取得
  const dirs = fs.readdirSync(dataDir)
    .filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d))
    .sort();

  // 各日のreviewCountを取得
  const dailyData = {};
  for (const dir of dirs) {
    const reviewsFile = path.join(dataDir, dir, 'reviews.json');
    if (!fs.existsSync(reviewsFile)) continue;
    const data = JSON.parse(fs.readFileSync(reviewsFile, 'utf-8'));
    dailyData[dir] = data.salons;
  }

  const dates = Object.keys(dailyData).sort();
  if (dates.length < 2) return [];

  // 週ごとにグループ化（月曜始まり）
  const weeks = {};
  for (const dateStr of dates) {
    const d = new Date(dateStr + 'T00:00:00');
    const dayOfWeek = d.getDay();
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const monday = new Date(d);
    monday.setDate(monday.getDate() + mondayOffset);
    const sunday = new Date(monday);
    sunday.setDate(sunday.getDate() + 6);
    const weekKey = formatShortDate(monday) + '-' + formatShortDate(sunday);

    if (!weeks[weekKey]) {
      weeks[weekKey] = { sortKey: monday.getTime(), firstDate: null, lastDate: null };
    }
    if (!weeks[weekKey].firstDate || dateStr < weeks[weekKey].firstDate) {
      weeks[weekKey].firstDate = dateStr;
    }
    if (!weeks[weekKey].lastDate || dateStr > weeks[weekKey].lastDate) {
      weeks[weekKey].lastDate = dateStr;
    }
  }

  // 週ごとの増減を計算
  const result = [];
  const sortedWeeks = Object.entries(weeks).sort((a, b) => a[1].sortKey - b[1].sortKey);
  for (const [weekLabel, weekInfo] of sortedWeeks) {
    const entry = { label: weekLabel, salons: {} };
    const first = dailyData[weekInfo.firstDate];
    const last = dailyData[weekInfo.lastDate];
    if (!first || !last) continue;

    for (const salon of SALONS) {
      if (first[salon.name] && last[salon.name]) {
        entry.salons[salon.name] = {
          diff: last[salon.name].reviewCount - first[salon.name].reviewCount,
          start: first[salon.name].reviewCount,
          end: last[salon.name].reviewCount,
        };
      }
    }
    result.push(entry);
  }

  return result;
}

/**
 * HTMLメールレポート生成
 */
function generateHtmlReport(dateStr, salonReports) {
  const prevDateStr = getPrevDateStr(dateStr);

  if (!salonReports) {
    return `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:20px;">
      <h2>📊 HPB空き状況レポート（${dateStr}）</h2>
      <p>初回取得のため差分データはありません。明日以降、前日との比較が届きます。</p>
    </body></html>`;
  }

  let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; padding: 20px; background: #f5f5f5; color: #333; }
    .container { max-width: 700px; margin: 0 auto; background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a1a1a; border-bottom: 2px solid #e91e63; padding-bottom: 8px; }
    h3 { color: #e91e63; margin-top: 24px; }
    .salon-section { margin-bottom: 32px; }
    table { border-collapse: collapse; width: 100%; margin: 8px 0 16px 0; }
    th { background: #f8f8f8; text-align: left; padding: 8px 12px; border-bottom: 2px solid #ddd; font-size: 13px; }
    td { padding: 6px 12px; border-bottom: 1px solid #eee; font-size: 13px; }
    .filled { color: #d32f2f; }
    .opened { color: #388e3c; font-weight: bold; }
    .weekend-note { color: #f57c00; font-size: 11px; }
    .bar-container { background: #eee; border-radius: 4px; height: 20px; width: 100%; position: relative; }
    .bar-fill { border-radius: 4px; height: 20px; transition: width 0.3s; }
    .bar-label { position: absolute; right: 6px; top: 2px; font-size: 11px; color: #555; }
    .stat-row td { padding: 8px 12px; }
    .summary-box { background: #fce4ec; border-radius: 8px; padding: 12px 16px; margin: 12px 0; }
    .summary-box.green { background: #e8f5e9; }
    .count-badge { display: inline-block; background: #e91e63; color: #fff; border-radius: 12px; padding: 2px 10px; font-size: 12px; margin-left: 8px; }
    .count-badge.green { background: #388e3c; }
    .count-badge.gray { background: #999; }
  </style></head><body><div class="container">`;

  html += `<h2>📊 HPB空き状況レポート（${dateStr}）</h2>`;
  html += `<p style="color:#888;font-size:13px;">前日（${prevDateStr}）との比較</p>`;

  for (const salon of salonReports) {
    html += `<div class="salon-section">`;
    html += `<h3>🏠 ${salon.name}</h3>`;

    // 埋まった枠
    html += `<p><strong>新しく埋まった枠</strong> <span class="count-badge">${salon.filled.length}件</span></p>`;
    if (salon.filled.length > 0) {
      html += `<table><tr><th>日付</th><th>時間</th><th>変化</th></tr>`;
      for (const item of salon.filled.sort(sortByDateTime)) {
        html += `<tr class="filled"><td>${formatDateWithDay(item.date)}</td><td>${item.time}</td><td>${item.from} → ${item.to}</td></tr>`;
      }
      html += `</table>`;
    }

    // 空いた枠
    html += `<p><strong>キャンセルで空いた枠</strong> <span class="count-badge green">${salon.opened.length}件</span></p>`;
    if (salon.opened.length > 0) {
      html += `<table><tr><th>日付</th><th>時間</th><th>変化</th><th></th></tr>`;
      for (const item of salon.opened.sort(sortByDateTime)) {
        const isWeekend = isWeekendDate(item.date);
        const note = isWeekend ? '<span class="weekend-note">⚠ 土日の後出し開放？</span>' : '';
        html += `<tr class="opened"><td>${formatDateWithDay(item.date)}</td><td>${item.time}</td><td>${item.from} → ${item.to}</td><td>${note}</td></tr>`;
      }
      html += `</table>`;
    } else {
      html += `<p style="color:#999;font-size:13px;">なし</p>`;
    }

    // 新規日付
    if (salon.newDates.length > 0) {
      html += `<p><strong>新規表示された日付</strong> <span class="count-badge gray">${salon.newDates.length}日</span></p>`;
    }

    // 週別サマリー（予約数バーチャート）
    const sortedWeeks = Object.entries(salon.weeklyStats).sort((a, b) => a[1].sortKey - b[1].sortKey);
    if (sortedWeeks.length > 0) {
      html += `<p><strong>週別 予約状況</strong></p>`;
      html += `<table class="stat-table"><tr><th>期間</th><th>予約状況</th><th>予約数</th><th>予約率</th></tr>`;
      for (const [weekLabel, stats] of sortedWeeks) {
        if (stats.total === 0) continue;
        const booked = stats.total - stats.available;
        const bookedPct = ((booked / stats.total) * 100).toFixed(1);
        const pctNum = parseFloat(bookedPct);
        let barColor = '#388e3c';
        if (pctNum >= 90) barColor = '#d32f2f';
        else if (pctNum >= 70) barColor = '#f57c00';
        else if (pctNum >= 50) barColor = '#fbc02d';
        html += `<tr class="stat-row"><td style="white-space:nowrap;">${weekLabel}</td>`;
        html += `<td><div class="bar-container"><div class="bar-fill" style="width:${bookedPct}%;background:${barColor};"></div><div class="bar-label">${booked}/${stats.total}</div></div></td>`;
        html += `<td style="text-align:center;font-weight:bold;">${booked}枠</td>`;
        html += `<td style="text-align:right;font-weight:bold;color:${barColor}">${bookedPct}%</td></tr>`;
      }
      html += `</table>`;
    }

    html += `</div>`;
  }

  // 口コミセクション（HTML版）
  html += generateHtmlReviewSection(dateStr);

  html += `<p style="color:#aaa;font-size:11px;margin-top:24px;text-align:center;">HPB空き状況モニター（自動送信）</p>`;
  html += `</div></body></html>`;
  return html;
}

/**
 * 口コミセクション（HTML版）
 */
function generateHtmlReviewSection(dateStr) {
  const outputDir = path.join(OUTPUT_BASE, dateStr);
  const todayReviewsFile = path.join(outputDir, 'reviews.json');
  if (!fs.existsSync(todayReviewsFile)) return '';

  const todayReviews = JSON.parse(fs.readFileSync(todayReviewsFile, 'utf-8'));
  let html = `<div style="margin-top:32px;border-top:2px solid #e91e63;padding-top:16px;">`;
  html += `<h3 style="color:#e91e63;">口コミ状況</h3>`;

  // 今日の口コミ数
  html += `<table><tr><th>サロン</th><th>口コミ数</th><th>評価</th><th>前日比</th></tr>`;
  const prevDateStr = getPrevDateStr(dateStr);
  const prevReviewsFile = path.join(OUTPUT_BASE, prevDateStr, 'reviews.json');
  const prevReviews = fs.existsSync(prevReviewsFile) ? JSON.parse(fs.readFileSync(prevReviewsFile, 'utf-8')) : null;

  for (const salon of SALONS) {
    const r = todayReviews.salons[salon.name];
    if (!r) continue;
    let diffHtml = '-';
    if (prevReviews && prevReviews.salons[salon.name]) {
      const diff = r.reviewCount - prevReviews.salons[salon.name].reviewCount;
      if (diff > 0) diffHtml = `<span style="color:#388e3c;font-weight:bold;">+${diff}</span>`;
      else if (diff === 0) diffHtml = `<span style="color:#999;">±0</span>`;
      else diffHtml = `<span style="color:#d32f2f;">${diff}</span>`;
    }
    html += `<tr><td>${salon.label}</td><td style="text-align:center;font-weight:bold;">${r.reviewCount}件</td><td style="text-align:center;">${r.rating}点</td><td style="text-align:center;">${diffHtml}</td></tr>`;
  }
  html += `</table>`;

  // 週別増減
  const weeklyReviews = collectWeeklyReviews();
  if (weeklyReviews.length > 0) {
    html += `<p><strong>週別 口コミ増減</strong></p>`;
    html += `<table><tr><th>期間</th>`;
    for (const salon of SALONS) html += `<th>${salon.label}</th>`;
    html += `</tr>`;
    for (const week of weeklyReviews) {
      html += `<tr><td style="white-space:nowrap;">${week.label}</td>`;
      for (const salon of SALONS) {
        const data = week.salons[salon.name];
        if (data) {
          const color = data.diff > 0 ? '#388e3c' : data.diff === 0 ? '#999' : '#d32f2f';
          const sign = data.diff > 0 ? '+' : '';
          html += `<td style="text-align:center;color:${color};font-weight:bold;">${sign}${data.diff}件</td>`;
        } else {
          html += `<td style="text-align:center;color:#999;">-</td>`;
        }
      }
      html += `</tr>`;
    }
    html += `</table>`;
  }

  html += `</div>`;
  return html;
}

/**
 * メール送信
 */
async function sendEmail(dateStr, reportText, reportHtml, outputDir) {
  const emailUser = process.env.EMAIL_USER;
  const emailPass = process.env.EMAIL_PASS;
  const emailTo = process.env.EMAIL_TO || emailUser;

  if (!emailUser || !emailPass) {
    console.log('\nメール設定がありません。メール送信をスキップします。');
    return;
  }

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: emailUser, pass: emailPass },
  });

  // スクリーンショットを添付
  const attachments = [];
  for (const salon of SALONS) {
    const week0 = path.join(outputDir, `${salon.name}_week0.png`);
    if (fs.existsSync(week0)) {
      attachments.push({
        filename: `${salon.name}_week0.png`,
        path: week0,
      });
    }
  }

  const mailOptions = {
    from: emailUser,
    to: emailTo,
    subject: `📊 HPB空き状況レポート（${dateStr}）`,
    text: reportText,
    html: reportHtml,
    attachments,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`\nメール送信完了: ${emailTo}`);
  } catch (err) {
    console.error(`\nメール送信エラー: ${err.message}`);
  }
}

async function main() {
  const dateStr = getDateStr();
  const outputDir = path.join(OUTPUT_BASE, dateStr);

  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`出力先: ${outputDir}`);
  console.log(`日付: ${dateStr}`);

  const browser = await chromium.launch({ headless: true });

  try {
    for (let i = 0; i < SALONS.length; i++) {
      if (i > 0) {
        console.log(`\nサロン間待機 (3秒)...`);
        await sleep(3000);
      }
      await processSalon(browser, SALONS[i], outputDir);
    }
  } finally {
    await browser.close();
  }

  // 口コミ取得（ブラウザ再起動）
  const browser2 = await chromium.launch({ headless: true });
  const reviews = {};
  try {
    for (let i = 0; i < SALONS.length; i++) {
      if (i > 0) await sleep(2000);
      reviews[SALONS[i].name] = await scrapeReviewData(browser2, SALONS[i]);
    }
  } finally {
    await browser2.close();
  }
  const reviewsPath = path.join(outputDir, 'reviews.json');
  fs.writeFileSync(reviewsPath, JSON.stringify({ date: dateStr, salons: reviews }, null, 2), 'utf-8');
  console.log(`\n口コミデータ保存: reviews.json`);

  const { text, html, hasPrev } = generateDiffReport(outputDir, dateStr);

  // メール送信
  await sendEmail(dateStr, text, html || '', outputDir);

  console.log(`\n完了！`);
}

main().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});
