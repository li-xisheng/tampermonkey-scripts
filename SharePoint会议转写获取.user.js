// ==UserScript==
// @name         SharePoint / Stream Transcript Grabber (トランスクリプト取得)
// @name:zh-CN   SharePoint / Stream 会议转写获取
// @namespace    https://github.com/local/sharepoint-stream-transcript
// @version      1.3.1
// @description  从 Microsoft Stream（SharePoint）的会议录像页面，或 Teams 会议“まとめ”页的转写面板获取完整的会议转写（自动滚动收集被虚拟滚动隐藏的内容），可选 .txt / .tsv / .json 格式保存
// @author       you
// @match        *://*.sharepoint.com/*
// @match        *://*.stream.office.com/*
// @match        *://*.microsoftstream.com/*
// @match        *://*.teams.microsoft.com/*
// @match        *://teams.cloud.microsoft/*
// @match        *://*.teams.cloud.microsoft/*
// @match        *://*.office.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

// Teams 的“まとめ”页里，转写面板嵌在 iframe 中，所以不加 @noframes，
// 只有页面里真的出现转写条目时才显示按钮

(function () {
  'use strict';

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---- 选择器（基于 id 前缀，不随版本构建变化） ----------------------
  // 转写条目由 OneTranscript（React 应用）渲染，带有语义化的 id：
  // entry-N（整条）、sub-entry-N（正文）、timestampSpeakerAriaLabel-N（隐藏的“发言人 时间”标签）。
  // 这些 id 比 Fluent UI mergeStyles 生成的类名哈希（如 baseEntry-447）稳定得多，后者每次构建都会变。
  const entryEls = () =>
    Array.from(document.querySelectorAll('[id^="entry-"]')).filter((el) =>
      /^entry-\d+$/.test(el.id)
    );

  function findScroller(node) {
    let n = node;
    while (n && n !== document.body) {
      const s = getComputedStyle(n);
      if (
        (s.overflowY === 'auto' || s.overflowY === 'scroll') &&
        n.scrollHeight > n.clientHeight + 20
      ) {
        return n;
      }
      n = n.parentElement;
    }
    return null;
  }

  // ---- 采集 ---------------------------------------------------------------
  function harvest(store) {
    for (const el of entryEls()) {
      const n = +el.id.slice('entry-'.length);
      const lblEl = document.getElementById('timestampSpeakerAriaLabel-' + n);
      const label = ((lblEl ? lblEl.textContent : '') || el.getAttribute('aria-label') || '').trim();
      const txtEl = document.getElementById('sub-entry-' + n);
      const text = txtEl ? txtEl.textContent.trim() : '';
      store[n] = { n, label, text };
    }
  }

  function sortedKeys(store) {
    return Object.keys(store).map(Number).sort((a, b) => a - b);
  }

  function gapsIn(store, hi) {
    const g = [];
    for (let k = 0; k <= hi; k++) if (!(k in store)) g.push(k);
    return g;
  }

  // ---- 主流程 -------------------------------------------------------------
  async function grab(onProgress) {
    const first = entryEls()[0];
    if (!first) throw new Error('找不到转写内容。请先打开左侧的“转写”（トランスクリプト）面板再执行。');
    const scroller = findScroller(first);
    if (!scroller) throw new Error('找不到可滚动的转写区域。');

    // 先暂停视频，避免转写随播放自动滚动，与脚本的滚动冲突而跳过一段条目。
    const video = document.querySelector('video');
    if (video) { try { video.pause(); } catch (e) {} }

    // 获取期间锁住转写面板，屏蔽页面自己的滚动（跟随播放位置），取完恢复
    const lock = lockScroll(scroller);
    try {
      return await sweep(scroller, lock.setTop, onProgress);
    } finally {
      lock.release();
    }
  }

  /**
   * 锁住滚动容器：页面对它的 scrollTop / scrollTo 等写入一律忽略，
   * 容器内元素的 scrollIntoView 也不生效。只有返回的 setTop 能真正滚动
   */
  function lockScroll(scroller) {
    const desc = Object.getOwnPropertyDescriptor(Element.prototype, 'scrollTop');
    const origIntoView = Element.prototype.scrollIntoView;
    const noop = () => {};
    Object.defineProperty(scroller, 'scrollTop', {
      configurable: true,
      get() { return desc.get.call(this); },
      set: noop,
    });
    ['scrollTo', 'scroll', 'scrollBy'].forEach(k => { scroller[k] = noop; });
    Element.prototype.scrollIntoView = function (...args) {
      if (scroller.contains(this)) return undefined;
      return origIntoView.apply(this, args);
    };
    return {
      setTop: v => desc.set.call(scroller, v),
      release() {
        ['scrollTop', 'scrollTo', 'scroll', 'scrollBy'].forEach(k => { delete scroller[k]; });
        Element.prototype.scrollIntoView = origIntoView;
      },
    };
  }

  async function sweep(scroller, setTop, onProgress) {
    const store = {};

    // 第 1 阶段：从上到下带重叠地扫一遍。
    // 位置由脚本自己记住（pos），万一仍被拉走也能回到原位，不会从头重扫
    setTop(0);
    await sleep(350);
    harvest(store);
    const step = Math.max(250, Math.floor(scroller.clientHeight * 0.5));
    let pos = 0;
    for (let i = 0; i < 400; i++) {
      const maxTop = scroller.scrollHeight - scroller.clientHeight;
      const atBottom = pos >= maxTop - 2;
      pos = Math.min(pos + step, maxTop);
      setTop(pos);
      await sleep(150);
      if (Math.abs(scroller.scrollTop - pos) > step) {
        setTop(pos);
        await sleep(150);
      }
      harvest(store);
      if (onProgress) onProgress(sortedKeys(store).length, null, '获取中');
      if (atBottom) { await sleep(250); harvest(store); break; }
    }

    // 预期总数：取正文条目上的 aria-setsize，没有时用见到的最大序号。
    let keys = sortedKeys(store);
    let maxN = keys[keys.length - 1];
    const sizedEl = document.querySelector('[id^="sub-entry-"][aria-setsize]');
    const setsize = sizedEl ? +sizedEl.getAttribute('aria-setsize') : 0;
    const hi = Math.max(maxN, setsize ? setsize - 1 : maxN);
    const total = hi + 1;

    // 第 2 阶段：按估算的像素位置跳到每个缺口来补齐（这一步跳到上方是正常的）。
    // 虚拟列表会在当前位置附近渲染一大段缓冲（约 180 条），跳准一次就能补上一整段连续的缺口。
    // 列表高度在第 1 阶段中可能变化，所以每次都按当前高度估算位置
    let gaps = gapsIn(store, hi);
    for (let iter = 0; iter < 80 && gaps.length; iter++) {
      const H = scroller.scrollHeight;
      const mid = gaps[Math.floor(gaps.length / 2)];
      const px = Math.floor(H * (mid / total));
      setTop(px); await sleep(240); harvest(store);
      setTop(Math.max(0, px - 900)); await sleep(160); harvest(store);
      setTop(Math.min(H, px + 900)); await sleep(160); harvest(store);
      gaps = gapsIn(store, hi);
      if (onProgress) onProgress(sortedKeys(store).length, total, '补齐中');
    }

    return { store, total, remainingGaps: gaps };
  }

  // ---- 格式化 ----------------------------------------------------------------
  // label -> {speaker, mmss}。“分”一定存在，整分钟时省略“秒”。
  // 发言人名里可能有数字（如“WeWork-602”），所以只匹配末尾的时间，前面的全部当作发言人。
  const TIME_RE = /\s*(?:(\d+)\s*時間)?\s*(\d+)\s*分間?\s*(?:(\d+)\s*秒間?)?\s*$/;
  function parseLabel(label) {
    const m = label.match(TIME_RE);
    if (!m) return { speaker: label.trim(), t: '' };
    const h = +(m[1] || 0), mm = +(m[2] || 0), ss = +(m[3] || 0);
    const speaker = label.slice(0, m.index).trim();
    const total = h * 3600 + mm * 60 + ss;
    const hh = Math.floor(total / 3600);
    const rm = Math.floor((total % 3600) / 60);
    const rs = total % 60;
    const pad = (x) => String(x).padStart(2, '0');
    const t = (hh > 0 ? pad(hh) + ':' : '') + pad(rm) + ':' + pad(rs);
    return { speaker, t };
  }

  function toRows(store) {
    const rows = [];
    for (const k of sortedKeys(store)) {
      const e = store[k];
      const text = (e.text || '').trim();
      const label = (e.label || '').trim();
      if (!label && /トランスクリプトを(開始|中止)しました/.test(text)) {
        rows.push({ n: e.n, kind: 'event', speaker: '', t: '', text });
        continue;
      }
      const { speaker, t } = parseLabel(label);
      rows.push({ n: e.n, kind: 'line', speaker, t, text });
    }
    return rows;
  }

  function buildTxt(rows, meta) {
    let out = meta.title + '\n' + meta.url + '\n' +
      '总条数：' + rows.length + '\n' + '='.repeat(80) + '\n\n';
    for (const r of rows) {
      if (r.kind === 'event') { out += '--- ' + r.text + ' ---\n\n'; continue; }
      out += '[' + r.t + '] ' + r.speaker + '\n' + r.text + '\n\n';
    }
    return out;
  }
  function buildTsv(rows) {
    let out = '';
    for (const r of rows) {
      if (r.kind === 'event') { out += '# ' + r.text + '\n'; continue; }
      out += r.t + '\t' + r.speaker + '\t' + r.text + '\n';
    }
    return out;
  }

  // 没有意义的页面标题（Teams 的 iframe 里常见），不能拿来当文件名
  const GENERIC_TITLE_RE = /^(|microsoft teams|teams|microsoft stream|stream|onetranscript|transcript|トランスクリプト)$/i;

  /** 找录像文件名：先看页面标题，再看页面地址里的 .mp4 路径（Teams iframe 的地址里带着它） */
  function recordingName() {
    const title = (document.title || '').trim();
    if (!GENERIC_TITLE_RE.test(title)) return title;
    let href = location.href;
    try { href = decodeURIComponent(href); } catch (e) { /* 解码失败就用原样 */ }
    const m = /([^/?&=#]+)\.mp4/i.exec(href);
    return m ? m[1] : '';
  }

  /**
   * 文件名：“会议名_日期”。
   * 录像名形如“会议名-20261026_090102UTC-Meeting Recording.mp4”，去掉多余部分；
   * 拿不到会议名时用“会议转写_导出时间”
   */
  function safeName() {
    const pad = (x) => String(x).padStart(2, '0');
    const now = new Date();
    const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
    const raw = recordingName()
      .replace(/\.mp4$/i, '')
      .replace(/[-_ ]*Meeting Recording$/i, '')
      .trim();
    const dated = /^(.*?)[-_ ]*(\d{8})_\d{6}(?:UTC)?$/i.exec(raw);
    const base = dated ? `${dated[1]}_${dated[2]}` : (raw || `会议转写_${stamp}`);
    return base.replace(/[\\/:*?"<>|\x00-\x1F]+/g, '_').trim().slice(0, 120) || `会议转写_${stamp}`;
  }

  function download(name, text, type) {
    const blob = new Blob([text], { type: type + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  // ---- UI --------------------------------------------------------------------
  // 获取比较耗时，所以先保留结果，之后再选择保存格式
  const FORMATS = [
    { key: 'txt', label: 'TXT', type: 'text/plain', build: (r) => buildTxt(r.rows, r.meta) },
    { key: 'tsv', label: 'TSV', type: 'text/tab-separated-values', build: (r) => buildTsv(r.rows) },
    { key: 'json', label: 'JSON', type: 'application/json',
      build: (r) => JSON.stringify({ meta: r.meta, total: r.total, rows: r.rows }, null, 2) },
  ];
  const PANEL_ID = '__tx_grab_panel';
  const BASE_FONT = '600 13px/1.2 "Segoe UI",system-ui,sans-serif';

  function styledButton(text, bg, color) {
    const b = document.createElement('button');
    b.textContent = text;
    Object.assign(b.style, {
      padding: '8px 12px', background: bg, color, border: 'none', borderRadius: '6px',
      font: BASE_FONT, cursor: 'pointer'
    });
    return b;
  }

  function closePanel() {
    const p = document.getElementById(PANEL_ID);
    if (p) p.remove();
  }

  function showFormatPanel(result) {
    closePanel();
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    Object.assign(panel.style, {
      position: 'fixed', zIndex: 2147483647, right: '16px', bottom: '64px',
      padding: '12px', background: '#fff', color: '#222', borderRadius: '8px',
      boxShadow: '0 4px 14px rgba(0,0,0,.3)', font: BASE_FONT, maxWidth: '320px'
    });

    const info = document.createElement('div');
    info.textContent = '已获取 ' + result.rows.length + ' 条';
    info.style.marginBottom = '8px';
    panel.appendChild(info);

    if (result.remainingGaps.length) {
      const warn = document.createElement('div');
      warn.textContent = '⚠ ' + result.remainingGaps.length + ' 条未取到，可重新获取补全';
      Object.assign(warn.style, { color: '#c62828', marginBottom: '8px', fontWeight: '400' });
      panel.appendChild(warn);
    }

    const row = document.createElement('div');
    Object.assign(row.style, { display: 'flex', gap: '6px', flexWrap: 'wrap' });
    for (const f of FORMATS) {
      const b = styledButton(f.label, '#5b2be6', '#fff');
      b.addEventListener('click', () => {
        download(result.base + '.' + f.key, f.build(result), f.type);
        b.textContent = '✓ ' + f.label;
      });
      row.appendChild(b);
    }
    const close = styledButton('关闭', '#eee', '#333');
    close.addEventListener('click', closePanel);
    row.appendChild(close);
    panel.appendChild(row);

    document.body.appendChild(panel);
  }

  function makeButton() {
    if (document.getElementById('__tx_grab_btn')) return;
    const btn = document.createElement('button');
    btn.id = '__tx_grab_btn';
    btn.textContent = '📝 获取会议转写';
    Object.assign(btn.style, {
      position: 'fixed', zIndex: 2147483647, right: '16px', bottom: '16px',
      padding: '10px 14px', background: '#5b2be6', color: '#fff', border: 'none',
      borderRadius: '8px', font: BASE_FONT,
      boxShadow: '0 4px 14px rgba(0,0,0,.3)', cursor: 'pointer'
    });
    btn.addEventListener('click', async () => {
      const original = btn.textContent;
      btn.disabled = true;
      closePanel();
      try {
        const meta = { title: recordingName() || document.title, url: location.href };
        const base = safeName();
        const { store, total, remainingGaps } = await grab((got, exp, phase) => {
          btn.textContent = (phase || '获取中') + '… ' + got + (exp ? '/' + exp : '');
        });
        // 获取过程中切换到了别的视频时，结果会混在一起，不允许保存
        if (location.href !== meta.url) throw new Error('获取过程中页面已切换，请重新执行。');
        showFormatPanel({ rows: toRows(store), meta, total, remainingGaps, base });
      } catch (e) {
        console.error('[TranscriptGrabber]', e);
        alert('获取失败：\n' + (e && e.message ? e.message : e));
      } finally {
        btn.disabled = false;
        btn.textContent = original;
      }
    });
    document.body.appendChild(btn);
  }

  // 切换到别的视频后，关闭上一个视频的保存面板（避免保存旧内容）
  let panelUrl = location.href;
  setInterval(() => {
    if (location.href !== panelUrl) { panelUrl = location.href; closePanel(); }
  }, 1000);

  // 页面是单页应用，转写面板随时可能打开或关闭，所以定期检查：
  // 有转写条目时显示按钮，没有时隐藏（获取中除外）
  function syncButton() {
    const has = entryEls().length > 0;
    const btn = document.getElementById('__tx_grab_btn');
    if (has && !btn) makeButton();
    if (btn && !btn.disabled) btn.style.display = has ? '' : 'none';
  }
  setInterval(syncButton, 1500);
  syncButton();
})();
