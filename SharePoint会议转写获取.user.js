// ==UserScript==
// @name         SharePoint / Stream Transcript Grabber (トランスクリプト取得)
// @name:zh-CN   SharePoint / Stream 会议转写获取
// @namespace    https://github.com/local/sharepoint-stream-transcript
// @version      1.2.1
// @description  从 Microsoft Stream（SharePoint）的会议录像页面获取完整的会议转写（自动滚动收集被虚拟滚动隐藏的内容），可选 .txt / .tsv / .json 格式保存
// @author       you
// @match        *://*.sharepoint.com/*
// @match        *://*.stream.office.com/*
// @match        *://*.microsoftstream.com/*
// @run-at       document-idle
// @grant        none
// @noframes
// ==/UserScript==

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

    const store = {};
    const H = scroller.scrollHeight;

    // 第 1 阶段：从上到下带重叠地扫一遍。
    scroller.scrollTop = 0;
    await sleep(350);
    harvest(store);
    const step = Math.max(250, Math.floor(scroller.clientHeight * 0.5));
    for (let i = 0; i < 400; i++) {
      const atBottom = scroller.scrollTop >= scroller.scrollHeight - scroller.clientHeight - 2;
      scroller.scrollTop = Math.min(scroller.scrollTop + step, scroller.scrollHeight);
      await sleep(150);
      harvest(store);
      if (onProgress) onProgress(sortedKeys(store).length, null);
      if (atBottom) { await sleep(250); harvest(store); break; }
    }

    // 预期总数：取正文条目上的 aria-setsize，没有时用见到的最大序号。
    let keys = sortedKeys(store);
    let maxN = keys[keys.length - 1];
    const sizedEl = document.querySelector('[id^="sub-entry-"][aria-setsize]');
    const setsize = sizedEl ? +sizedEl.getAttribute('aria-setsize') : 0;
    const hi = Math.max(maxN, setsize ? setsize - 1 : maxN);
    const total = hi + 1;

    // 第 2 阶段：按估算的像素位置跳到每个缺口来补齐。
    // 虚拟列表会在当前位置附近渲染一大段缓冲（约 180 条），跳准一次就能补上一整段连续的缺口。
    let gaps = gapsIn(store, hi);
    for (let iter = 0; iter < 80 && gaps.length; iter++) {
      const mid = gaps[Math.floor(gaps.length / 2)];
      const px = Math.floor(H * (mid / total));
      scroller.scrollTop = px; await sleep(240); harvest(store);
      scroller.scrollTop = Math.max(0, px - 900); await sleep(160); harvest(store);
      scroller.scrollTop = Math.min(H, px + 900); await sleep(160); harvest(store);
      gaps = gapsIn(store, hi);
      if (onProgress) onProgress(sortedKeys(store).length, total);
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

  function safeName() {
    const base = (document.title || 'transcript')
      .replace(/\.mp4$/i, '')
      .replace(/[\\/:*?"<>|]+/g, '_')
      .slice(0, 120);
    return base || 'transcript';
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
        const meta = { title: document.title, url: location.href };
        const base = safeName();
        const { store, total, remainingGaps } = await grab((got, exp) => {
          btn.textContent = '获取中… ' + got + (exp ? '/' + exp : '');
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

  // Stream 页面是单页应用，需要反复尝试挂上按钮。
  const iv = setInterval(makeButton, 1500);
  makeButton();
  // 2 分钟后停止尝试（按钮挂上后会一直保留）
  setTimeout(() => clearInterval(iv), 120000);
})();
