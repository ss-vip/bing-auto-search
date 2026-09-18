'use strict';
/* Bing Auto Search - unit/integration tests (no framework, plain node).
 * Run from repo root:  node tests/test.js
 * Each test builds a fresh vm sandbox via tests/harness.js with stubbed
 * browser APIs, fake timers and a mockable clock.
 */
const { createEnv } = require('./harness');

const TAIPEI = (month, day, h, min = 0) => Date.UTC(2026, month - 1, day, h - 8, min, 0);
const D18_08 = TAIPEI(9, 18, 8, 0);           // 2026-09-18 08:00 Asia/Taipei
const GM_KEY = 'bingAutoSearch';
const seedGM = (env, obj) => env.gm.set(GM_KEY, JSON.stringify(obj));
const yesterdayFull = () => ({ date: '2026-09-17', lastDate: '2026-09-17', pc_count: 45, ph_count: 35, autoStart: true });

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('PASS ' + name); }
  else { fail++; console.log('FAIL ' + name + (extra ? ' :: ' + extra : '')); }
}
async function t(name, fn) {
  try { await fn(); }
  catch (e) { fail++; console.log('FAIL ' + name + ' :: threw: ' + (e && e.stack || e)); }
}

(async () => {
  // ---------- T1: fresh morning auto-start ----------
  await t('T1 fresh morning auto-starts', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, yesterdayFull());
    const { api } = env;
    await api.init();
    await env.flush();
    ok('T1 start scheduled', env.pendingTimers().includes(1500));
    await env.advance(1500); await env.flush();
    const s = api._state();
    ok('T1 running', s.taskStatus === 'running', s.taskStatus);
    ok('T1 loop active', s.timerActive === true);
  });

  // ---------- T2: no login gate - anonymous header starts fine ----------
  await t('T2 anonymous header does not block auto-start', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, yesterdayFull());
    env.qsMap.set('#id_a', { textContent: 'Sign in' }); // logged-out DOM
    const { api } = env;
    await api.init();
    await env.flush();
    await env.advance(1500); await env.flush();
    const s = api._state();
    ok('T2 running while logged out', s.taskStatus === 'running', s.taskStatus);
    ok('T2 loop active', s.timerActive === true);
  });

  // ---------- T3: cross-day reset + resume ----------
  await t('T3 cross-day resets counts and resumes RESTING', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, yesterdayFull());
    const { api } = env;
    api._set({ lastSeenDate: '2026-09-17', taskStatus: 'resting' });
    api.checkAndResetDay();
    const gm = env.gmJson(GM_KEY);
    ok('T3 counts zeroed', gm.pc_count === 0 && gm.ph_count === 0, JSON.stringify(gm));
    ok('T3 date rolled', gm.lastDate === '2026-09-18', gm.lastDate);
    ok('T3 autoStart true', gm.autoStart === true);
    ok('T3 resumed', api._state().taskStatus === 'running');
    ok('T3 loop active', api._state().timerActive === true);
  });

  // ---------- T4: same-day RESTING resume ----------
  await t('T4 same-day RESTING resumes when counts allow', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 0, autoStart: true });
    const { api } = env;
    api._set({ taskStatus: 'resting' });
    api.checkAndResetDay();
    ok('T4 resumed', api._state().taskStatus === 'running');
    ok('T4 loop active', api._state().timerActive === true);
  });

  // ---------- T5: RESTING + maxed stays put ----------
  await t('T5 RESTING with maxed counts does not resume', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 35, autoStart: true });
    const { api } = env;
    api._set({ taskStatus: 'resting' });
    api.checkAndResetDay();
    ok('T5 stays resting', api._state().taskStatus === 'resting');
    ok('T5 no loop', api._state().timerActive === false);
  });

  // ---------- T6: overdue schedule watchdog fires the search ----------
  await t('T6 overdue schedule fires the search', async () => {
    const env = createEnv({ now: D18_08 });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.startSearchLoop();
    const t0 = api._state().nextExecuteTime;
    ok('T6 schedule set', t0 > 0);
    api.stopTimer(); // simulate lost loop handle (e.g. after freeze)
    env.setMockNow(t0 + 1000);
    api.checkScheduledExecution();
    await env.flush();
    await env.advance(300); await env.flush();
    ok('T6 fired once', env.gmJson(GM_KEY).pc_count === 1, JSON.stringify(env.gmJson(GM_KEY)));
    ok('T6 navigated', env.hrefWrites.length === 1, JSON.stringify(env.hrefWrites));
  });

  // ---------- T7: concurrent double-fire counts once ----------
  await t('T7 count lock dedups concurrent performSearch', async () => {
    const env = createEnv({ now: D18_08 });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.performSearch();
    api.performSearch();
    await env.flush();
    await env.advance(300); await env.flush();
    ok('T7 counted once', env.gmJson(GM_KEY).pc_count === 1, JSON.stringify(env.gmJson(GM_KEY)));
    ok('T7 navigated once', env.hrefWrites.length === 1, JSON.stringify(env.hrefWrites));
    ok('T7 loop alive for retry', api._state().timerActive === true);
  });

  // ---------- T8: atMax completes ----------
  await t('T8 maxed counts complete the task', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 0, autoStart: true });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.performSearch();
    await env.flush();
    ok('T8 resting', api._state().taskStatus === 'resting');
    ok('T8 countdown done', env.elements.get('br_countdown').textContent === '完成');
  });

  // ---------- T9: startSearch merged atMax branch ----------
  await t('T9 startSearch reports which side hit max', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 0, autoStart: true });
    const { api } = env;
    api.startSearch();
    const s = api._state();
    ok('T9 resting', s.taskStatus === 'resting');
    ok('T9 pc message', env.elements.get('br_status_text').textContent === '桌面版任務已達標');
  });

  // ---------- T10: english branch uses pool keywords (fix A) ----------
  await t('T10 getEnWordKeyword builds on keywordsPool', async () => {
    const env = createEnv({ now: D18_08, seed: 7 });
    const { api } = env;
    const pool = api._state().keywordsPool;
    const seen = new Set();
    let allContainBase = true;
    for (let i = 0; i < 120; i++) {
      const kw = await api.getEnWordKeyword();
      seen.add(kw);
      if (!pool.some((k) => kw.includes(k))) allContainBase = false;
    }
    ok('T10 base always from pool (no fix-only strings)', allContainBase);
    ok('T10 diverse output', seen.size > 20, 'unique=' + seen.size);
  });

  // ---------- T11: pure keyword helpers ----------
  await t('T11 keyword helpers behave', async () => {
    const env = createEnv();
    const { api } = env;
    ok('T11 dedup words', api.removeDuplicateWords('a a b') === 'a b');
    ok('T11 filter fixes', JSON.stringify(api.filterDuplicateFixes('Python 教學', ['教學', '最新'])) === JSON.stringify(['最新']));
    ok('T11 merge keywords', JSON.stringify(api.mergeAndDeduplicateKeywords(['b', 'a', ''], ['a'])) === JSON.stringify(['a', 'b']));
    ok('T11 escape html', api.escapeHtml('<b>"x"&') === '&lt;b&gt;&quot;x&quot;&amp;');
  });

  // ---------- T12: /search stuck submit triggers forced redirect + recovery ----------
  await t('T12 silent submit failure on /search recovers', async () => {
    const env = createEnv({ now: D18_08, url: 'https://www.bing.com/search?q=old&FORM=HDRS2' });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 0, autoStart: true });
    env.hooks.formSubmit = () => {}; // submit silently does nothing
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.performSearch();
    await env.flush();
    await env.advance(300); await env.flush();
    ok('T12 no premature redirect', env.hrefWrites.length === 0);
    await env.advance(4000); await env.flush();
    ok('T12 forced redirect', env.hrefWrites.some((h) => h.includes('/search?q=') && !h.includes('q=old')), JSON.stringify(env.hrefWrites));
    ok('T12 loop recovered', api._state().timerActive === true);
    ok('T12 fail counted', env.ss.get('bing_redirect_fails') === '1');
  });

  // ---------- T13: successful submit has no extra redirect ----------
  await t('T13 successful submit leaves watchdog quiet', async () => {
    const env = createEnv({ now: D18_08, url: 'https://www.bing.com/' });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.performSearch();
    await env.flush();
    await env.advance(300); await env.flush();
    ok('T13 submitted once', env.hrefWrites.length === 1, JSON.stringify(env.hrefWrites));
    await env.advance(4000); await env.flush();
    ok('T13 no forced redirect', env.hrefWrites.length === 1);
    ok('T13 fails cleared', env.ss.get('bing_redirect_fails') === undefined);
  });

  // ---------- T14: two redirect fails pause the task ----------
  await t('T14 repeated redirect failure pauses', async () => {
    const env = createEnv({ now: D18_08, url: 'https://www.bing.com/search?q=old&FORM=HDRS2' });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 0, autoStart: true });
    env.hooks.formSubmit = () => {};
    env.ss.set('bing_redirect_fails', '2');
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.performSearch();
    await env.flush();
    await env.advance(4300); await env.flush();
    ok('T14 paused', api._state().taskStatus === 'paused');
    ok('T14 countdown reset', env.elements.get('br_countdown').textContent === '--');
    ok('T14 fails cleared', env.ss.get('bing_redirect_fails') === undefined);
  });

  // ---------- T15: updateStatusAfterInit transition ----------
  await t('T15 init-status reconciles RUNNING+maxed to RESTING', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 0, autoStart: true });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.updateStatusAfterInit();
    ok('T15 resting', api._state().taskStatus === 'resting');
    ok('T15 complete text', env.elements.get('br_status_text').textContent === '任務已完成! 等待明日...');
  });

  // ---------- T16: toggle pause/resume ----------
  await t('T16 toggle pauses and resumes', async () => {
    const env = createEnv({ now: D18_08 });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.toggleScript();
    ok('T16 paused', api._state().taskStatus === 'paused');
    ok('T16 resume label', env.elements.get('br_toggle_btn').textContent === '▶ 繼續搜尋');
    api.toggleScript();
    ok('T16 running again', api._state().taskStatus === 'running');
    ok('T16 pause label', env.elements.get('br_toggle_btn').textContent === '⏸ 暫停搜尋');
  });

  // ---------- T17: task ownership across tabs ----------
  await t('T17 claim/heartbeat/owner-expiry across tabs', async () => {
    const envA = createEnv({ now: D18_08 });
    const envB = createEnv({ now: D18_08, sharedLS: envA.ls });
    const envC = createEnv({ now: D18_08, sharedLS: envA.ls });
    envB.api._set({ tabId: 'tabB' });
    envC.api._set({ tabId: 'tabC' });
    ok('T17 A claims', envA.api.claimTask() === true);
    envC.api.releaseTask(); // foreign id -> must not remove A's claim
    ok('T17 foreign release ignored', envA.ls.has('bing_task_owner_pc'));
    ok('T17 B blocked', envB.api.claimTask() === false);
    envA.api.heartbeatTask();
    await envB.advance(20000); await envB.flush();
    ok('T17 B still blocked inside window', envB.api.claimTask() === false);
    await envB.advance(15000); await envB.flush();
    ok('T17 B takes over after expiry', envB.api.claimTask() === true);
    ok('T17 ownership transferred', JSON.parse(envA.ls.get('bing_task_owner_pc')).id === 'tabB');
    envB.api.releaseTask();
    ok('T17 owner release clears', !envA.ls.has('bing_task_owner_pc'));
  });

  // ---------- T18: date handling across midnight Taipei time ----------
  await t('T18 getToday respects Asia/Taipei', async () => {
    const env = createEnv({ now: TAIPEI(9, 17, 23, 59) });
    ok('T18 before midnight', env.api.getToday() === '2026-09-17', env.api.getToday());
    env.setMockNow(TAIPEI(9, 17, 23, 59) + 120000);
    ok('T18 after midnight', env.api.getToday() === '2026-09-18', env.api.getToday());
  });

  // ---------- T19: run quotas per page type ----------
  await t('T19 canRunSearch honors per-type quotas', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 44, ph_count: 35, autoStart: true });
    const { api } = env;
    ok('T19 pc can run', api.canRunSearch(api.getConfig()) === true);
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 35, autoStart: true });
    ok('T19 pc maxed', api.canRunSearch(api.getConfig()) === false);
    env.setUrl('https://m.bing.com/');
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 45, ph_count: 34, autoStart: true });
    ok('T19 ph can run', api.canRunSearch(api.getConfig()) === true);
  });

  // ---------- T20: full loop performs exactly one search ----------
  await t('T20 full timer loop performs one search then waits on navigation', async () => {
    const env = createEnv({ now: D18_08, seed: 99 });
    seedGM(env, yesterdayFull());
    const { api } = env;
    await api.init();
    await env.flush();
    await env.advance(1500); await env.flush();
    ok('T20 running', api._state().taskStatus === 'running');
    await env.advance(125000); await env.flush();
    ok('T20 counted once', env.gmJson(GM_KEY).pc_count === 1, JSON.stringify(env.gmJson(GM_KEY)));
    ok('T20 navigated once', env.hrefWrites.length === 1, JSON.stringify(env.hrefWrites));
    ok('T20 still running', api._state().taskStatus === 'running');
  });

  // ---------- T21: search history capped + split by type ----------
  await t('T21 history capped at 3 and split pc/ph', async () => {
    const env = createEnv({ now: D18_08 });
    const { api } = env;
    api.addSearchHistory('kw1'); api.addSearchHistory('kw2');
    api.addSearchHistory('kw3'); api.addSearchHistory('kw4');
    const h = api.getSearchHistory();
    ok('T21 capped', h.length === 3, String(h.length));
    ok('T21 newest first', h[0].keyword === 'kw4');
    env.setUrl('https://m.bing.com/');
    ok('T21 ph separate', api.getSearchHistory().length === 0);
    api.addSearchHistory('m1');
    ok('T21 ph stored', api.getSearchHistory().length === 1);
  });

  // ---------- T22: cleanCount resets ----------
  await t('T22 cleanCount zeroes counts and pauses', async () => {
    const env = createEnv({ now: D18_08 });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 10, ph_count: 5, autoStart: true });
    const { api } = env;
    api._set({ taskStatus: 'running' });
    api.cleanCount(null);
    const gm = env.gmJson(GM_KEY);
    ok('T22 zeroed', gm.pc_count === 0 && gm.ph_count === 0);
    ok('T22 autoStart off', gm.autoStart === false);
    ok('T22 paused', api._state().taskStatus === 'paused');
  });

  // ---------- T23: mobile page counts ph quota ----------
  await t('T23 mobile page counts ph quota', async () => {
    const env = createEnv({ now: D18_08, url: 'https://m.bing.com/' });
    seedGM(env, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 0, autoStart: true });
    const { api } = env;
    ok('T23 page type ph', api.getBingPageType() === 'ph');
    api._set({ taskStatus: 'running' });
    api.performSearch();
    await env.flush();
    await env.advance(300); await env.flush();
    const gm = env.gmJson(GM_KEY);
    ok('T23 ph counted, pc untouched', gm.ph_count === 1 && gm.pc_count === 0, JSON.stringify(gm));
    ok('T23 navigated on m.bing', env.hrefWrites.length === 1 && env.hrefWrites[0].startsWith('https://m.bing.com/search?q='), JSON.stringify(env.hrefWrites));
    const env2 = createEnv({ now: D18_08, url: 'https://m.bing.com/' });
    seedGM(env2, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 35, autoStart: true });
    env2.api._set({ taskStatus: 'running' });
    env2.api.performSearch();
    await env2.flush();
    ok('T23 ph maxed completes', env2.api._state().taskStatus === 'resting');
  });

  // ---------- T24: UA-based type detection ----------
  await t('T24 mobile UA detected as ph on www', async () => {
    const env = createEnv({ ua: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Mobile/15E148' });
    ok('T24 iphone UA is ph', env.api.getBingPageType() === 'ph');
    const env2 = createEnv();
    ok('T24 desktop UA is pc', env2.api.getBingPageType() === 'pc');
  });

  // ---------- T25: pc + ph tabs share GM without clobbering ----------
  await t('T25 pc+ph tabs coexist without clobbering counts', async () => {
    const sharedGM = new Map(), sharedLS = new Map();
    const envA = createEnv({ now: D18_08, url: 'https://www.bing.com/', sharedGM, sharedLS, seed: 11 });
    const envB = createEnv({ now: D18_08, url: 'https://m.bing.com/', sharedGM, sharedLS, seed: 22 });
    seedGM(envA, { date: '2026-09-18', lastDate: '2026-09-18', pc_count: 0, ph_count: 0, autoStart: true });
    envA.api._set({ taskStatus: 'running' });
    envB.api._set({ taskStatus: 'running' });
    envA.api.performSearch();
    envB.api.performSearch(); // collides on the global count lock -> reschedules
    await envA.flush(); await envB.flush();
    await envA.advance(500); await envA.flush();
    const mid = envA.gmJson(GM_KEY);
    ok('T25 one counted, none lost', (mid.pc_count + mid.ph_count) === 1, JSON.stringify(mid));
    await envB.advance(125000); await envB.flush();
    const end = envA.gmJson(GM_KEY);
    ok('T25 both quotas reach 1', end.pc_count === 1 && end.ph_count === 1, JSON.stringify(end));
  });

  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
