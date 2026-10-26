// ==UserScript==
// @name         SharePoint / Stream Transcript Grabber (トランスクリプト取得)
// @namespace    https://github.com/local/sharepoint-stream-transcript
// @version      1.0.0
// @description  Microsoft Stream (SharePoint) の会議録画ページから、仮想スクロールで消えるトランスクリプトを全件取得して .txt / .json / .tsv で保存する。ネットワーク応答は暗号化されているため DOM を自動スクロールして収集する。
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

  // ---- selectors (id-prefix based = build-independent) ----------------------
  // Transcript entries are rendered by the OneTranscript React app with semantic
  // ids: entry-N (the group), sub-entry-N (the text), timestampSpeakerAriaLabel-N
  // (hidden "<speaker> <time>" label). These are far more stable than Fluent UI
  // mergeStyles class hashes (e.g. baseEntry-447), whose suffixes change per build.
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

  // ---- harvest ---------------------------------------------------------------
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

  // ---- main grab -------------------------------------------------------------
  async function grab(onProgress) {
    const first = entryEls()[0];
    if (!first) throw new Error('トランスクリプトのエントリが見つかりません。左側の「トランスクリプト」パネルを開いてから実行してください。');
    const scroller = findScroller(first);
    if (!scroller) throw new Error('スクロール可能なトランスクリプト コンテナが見つかりませんでした。');

    // Pause the video so its auto-scroll (transcript follows playback) does not
    // fight our scrolling and skip a band of entries.
    const video = document.querySelector('video');
    if (video) { try { video.pause(); } catch (e) {} }

    const store = {};
    const H = scroller.scrollHeight;

    // Phase 1: top -> bottom sweep with overlap.
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

    // Expected total: aria-setsize on a text entry, else max index seen.
    let keys = sortedKeys(store);
    let maxN = keys[keys.length - 1];
    const sizedEl = document.querySelector('[id^="sub-entry-"][aria-setsize]');
    const setsize = sizedEl ? +sizedEl.getAttribute('aria-setsize') : 0;
    const hi = Math.max(maxN, setsize ? setsize - 1 : maxN);
    const total = hi + 1;

    // Phase 2: fill gaps by jumping to each gap's estimated pixel position.
    // The virtualizer renders a large buffer (~180 entries) around any position,
    // so one well-placed jump fills a whole contiguous hole.
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

  // ---- format ----------------------------------------------------------------
  // label -> {speaker, mmss}. minutes are always present; seconds omitted on an
  // exact minute. Speaker may contain digits (e.g. "WeWork-602"), so match the
  // TRAILING time token and treat everything before it as the speaker.
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
      '総エントリ数: ' + rows.length + '\n' + '='.repeat(80) + '\n\n';
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
  function makeButton() {
    if (document.getElementById('__tx_grab_btn')) return;
    const btn = document.createElement('button');
    btn.id = '__tx_grab_btn';
    btn.textContent = '📝 トランスクリプト取得';
    Object.assign(btn.style, {
      position: 'fixed', zIndex: 2147483647, right: '16px', bottom: '16px',
      padding: '10px 14px', background: '#5b2be6', color: '#fff', border: 'none',
      borderRadius: '8px', font: '600 13px/1.2 "Segoe UI",system-ui,sans-serif',
      boxShadow: '0 4px 14px rgba(0,0,0,.3)', cursor: 'pointer'
    });
    btn.addEventListener('click', async () => {
      const original = btn.textContent;
      btn.disabled = true;
      try {
        const meta = { title: document.title, url: location.href };
        const { store, total, remainingGaps } = await grab((got, exp) => {
          btn.textContent = '取得中… ' + got + (exp ? '/' + exp : '');
        });
        const rows = toRows(store);
        const base = safeName();
        download(base + '.txt', buildTxt(rows, meta), 'text/plain');
        download(base + '.json', JSON.stringify({ meta, total, rows }, null, 2), 'application/json');
        download(base + '.tsv', buildTsv(rows), 'text/tab-separated-values');
        const warn = remainingGaps && remainingGaps.length
          ? '（未取得 ' + remainingGaps.length + ' 件。もう一度実行すると補完されます）'
          : '';
        btn.textContent = '✅ ' + rows.length + ' 件 保存' + warn;
      } catch (e) {
        console.error('[TranscriptGrabber]', e);
        alert('取得に失敗しました:\n' + (e && e.message ? e.message : e));
        btn.textContent = original;
      } finally {
        btn.disabled = false;
        setTimeout(() => { btn.textContent = original; }, 6000);
      }
    });
    document.body.appendChild(btn);
  }

  // The Stream page is a SPA; keep trying to attach the button.
  const iv = setInterval(makeButton, 1500);
  makeButton();
  // stop after 2 min of attaching (button persists once added)
  setTimeout(() => clearInterval(iv), 120000);
})();
