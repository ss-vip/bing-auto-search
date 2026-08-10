const fs = require('fs');
const path = require('path');

const RSS_URL = 'https://www.digitimes.com.tw/tech/rss/xml/xmlrss_10_60.xml';
const FILE = path.join(__dirname, '..', 'example.json');
const MAX_TOTAL = 50;

function cleanTitle(raw) {
  let s = raw.replace(/\u3000/g, ' ');   // 全形空格 → 半形
  while (/<[^>]*>/.test(s)) s = s.replace(/<[^>]*>/g, '');  // 反覆移除 HTML tag（處理 <<script> 等畸形輸入）
  return s
    .replace(/\s+/g, ' ')      // 壓縮空白
    .trim()
    .replace(/^[【\[（(][^】\]）)]*[】\]）)]\s*/, '');  // 移除系列/專欄前綴，如【動物農莊】（專訪）
}

async function main() {
  const res = await fetch(RSS_URL);
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const xml = await res.text();

  const titles = [];
  for (const item of xml.split('<item>').slice(1)) {
    const m = item.match(/<title><!\[CDATA\[([^\]]*)\]\]><\/title>/);
    if (!m) continue;
    const t = cleanTitle(m[1]);
    if (t.length < 4 || t.length > 40) continue;          // 過短或過長
    if (!/[a-zA-Z\u4e00-\u9fff]/.test(t)) continue;        // 至少含一個中英文字元
    if (!titles.includes(t)) titles.push(t);
  }

  if (titles.length === 0) {
    console.log('RSS 無有效標題，跳過');
    return;
  }

  const data = JSON.parse(fs.readFileSync(FILE, 'utf8'));
  // 直接以最新 RSS 標題取代舊 keywords（不再合併累積）
  data.keywords = titles.slice(0, MAX_TOTAL);

  const next = JSON.stringify(data, null, 2) + '\n';
  if (next !== fs.readFileSync(FILE, 'utf8')) {
    fs.writeFileSync(FILE, next);
    console.log(`更新完成：keywords 已取代為最新 ${data.keywords.length} 筆`);
  } else {
    console.log('無變化，跳過');
  }
}

main().catch(e => {
  console.error('更新失敗:', e.message);
  process.exit(1);
});
