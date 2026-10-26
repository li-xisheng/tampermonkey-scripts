// ==UserScript==
// @name         Teams Chat Exporter
// @name:zh-CN   Teams 聊天导出
// @namespace    http://tampermonkey.net/
// @version      2026.10.26.10
// @description  按“今天”“本周”等时间范围自动获取 Teams 对话，保存为 HTML（图片内嵌）或 Markdown（含图片和附件的 ZIP），也支持多个对话合并保存
// @author       You
// @match        https://teams.cloud.microsoft/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=microsoft.com
// @grant        GM_xmlhttpRequest
// @connect      teams.microsoft.com
// @connect      asm.skype.com
// @connect      sharepoint.com
// @connect      office.com
// @connect      live.com
// @connect      1drv.ms
// @noframes
// @license      MIT
// ==/UserScript==

// 原版是 Greasy Fork 上的 Teams Chat Exporter（script 537462）。
// 已大幅改写，为避免被自动更新覆盖，删除了 @updateURL / @downloadURL。

(() => {
    'use strict';

    // Teams 启用了 Trusted Types CSP，禁止给 innerHTML 赋值，也不能解析 HTML 字符串。
    // 所以界面只用 DOM API 搭建，HTML / Markdown 在收集时直接从 DOM 生成

    const LIST_ID = 'chat-pane-list';
    const UI_ID = 'tce-root';
    const PREFS_KEY = 'tce-prefs';
    const COLLECT_THROTTLE_MS = 300;
    const WATCH_INTERVAL_MS = 1500;
    const AUTO_SCROLL = { waitMs: 1500, idleRoundsToStop: 4, maxRounds: 400 };
    const FETCH_CONCURRENCY = 4;
    const WEEKDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
    const FILE_LINK_RE = /sharepoint\.com|onedrive\.|1drv\.ms|\.(docx?|xlsx?|pptx?|pdf|zip|csv|txt)(?:[?#]|$)/i;
    const ATTACHMENT_LABEL_RE = /^([^\n]+)\n(https:\/\/\S+)/;
    // 允许获取图片和附件的域名。限定范围，避免带着登录状态去请求消息里的任意地址
    const ASSET_HOST_RE = /(^|\.)(teams\.microsoft\.com|teams\.cloud\.microsoft|asm\.skype\.com|sharepoint\.com|office\.com|onedrive\.live\.com|1drv\.ms)$/i;
    const IMG_MARK = '\u0001'; // Markdown 中图片的占位标记。正文文字里会去掉这个字符，所以不会冲突
    const IMAGE_EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/bmp': 'bmp' };
    const FORMATS = {
        html: { label: 'HTML' },
        md: { label: 'Markdown (ZIP)' },
    };

    const state = {
        messages: new Map(), // 当前对话：消息 ID -> 快照
        chat: { id: '', title: '' },
        chatList: [],        // 左侧对话列表：{ id, title, time }
        selected: new Set(), // 勾选的对话 ID
        batchRun: 0,         // 批量导出的轮次编号，中止时加 1
        batchLabel: '',      // 批量导出进度（如“2/5 · 对话名”）
        collectRun: 0,       // 每次自动获取的编号。只用本次看到的消息来判断“是否已翻到开始日期”
        observer: null,
        observedList: null,
        collecting: false,
        autoScrolling: false,
        scrollRun: 0,        // 自动获取的轮次编号，用来确保旧的循环会停下
        scrollWarning: '',   // 最近一次自动获取可能有遗漏时的提示文字
        busy: false,
        notice: { text: '', isError: false },
    };
    let ui = null;
    let collectTimer = null;

    // ---- 通用 -----------------------------------------------------------------

    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const pad = (n, w = 2) => String(n).padStart(w, '0');
    const fmtDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    const fmtDay = d => `${fmtDate(d)} (${WEEKDAYS[d.getDay()]})`;
    const fmtTime = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    const fmtDateTime = d => `${fmtDate(d)} ${fmtTime(d)}`;
    const compactDate = d => fmtDate(d).replace(/-/g, '');
    const stamp = d => `${compactDate(d)}-${pad(d.getHours())}${pad(d.getMinutes())}`;
    const esc = s => String(s).replace(/[&<>"']/g, c => (
        { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
    const safeName = (s, max = 60) =>
        String(s || '').replace(/[<>:"/\\|?*#%\x00-\x1F]+/g, '_').replace(/\s+/g, ' ').trim()
            .replace(/^\.+/, '_').slice(0, max) || 'untitled';
    /** 作为 Markdown 普通文字安全输出（转义语法字符，并去掉换行） */
    const mdText = s => String(s).replace(/[\u0000-\u001F]/g, ' ').replace(/([\\`*_[\]<>#!|])/g, '\\$1');
    /** Markdown 链接地址。默认用 <...> 包裹，http(s)、mailto 和相对路径以外的都丢弃 */
    const mdUrl = u => {
        const s = String(u || '');
        if (/^[a-z][\w+.-]*:/i.test(s) && !/^(https?|mailto):/i.test(s)) return '';
        return s.replace(/[<>\s]/g, c => encodeURIComponent(c));
    };
    const encodeText = s => new TextEncoder().encode(s);
    const errorText = e => (e && e.message ? e.message : String(e));

    function uniqueName(used, name) {
        const dot = name.lastIndexOf('.');
        let candidate = name;
        for (let i = 2; used.has(candidate); i++) {
            candidate = dot > 0 ? `${name.slice(0, dot)} (${i})${name.slice(dot)}` : `${name} (${i})`;
        }
        used.add(candidate);
        return candidate;
    }

    async function mapPool(items, limit, fn) {
        let next = 0;
        const worker = async () => {
            while (next < items.length) {
                const i = next++;
                await fn(items[i], i);
            }
        };
        await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
    }

    function loadPrefs() {
        try {
            return { mode: 'quick', format: 'html', withFiles: false, ...JSON.parse(localStorage.getItem(PREFS_KEY) || '{}') };
        } catch (_) {
            return { mode: 'quick', format: 'html', withFiles: false };
        }
    }

    function savePrefs() {
        try {
            localStorage.setItem(PREFS_KEY, JSON.stringify({ mode: ui.mode, format: ui.format.value, withFiles: ui.withFiles.checked }));
        } catch (_) { /* 保存失败也不影响使用 */ }
    }

    // ---- ZIP（STORE 方式、无依赖，与微信公众号转Markdown 相同） --------------

    const CRC_TABLE = (() => {
        const t = new Uint32Array(256);
        for (let i = 0; i < 256; i++) {
            let c = i;
            for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            t[i] = c;
        }
        return t;
    })();

    function crc32(data) {
        let crc = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    /** files: Array<{ name: string, data: Uint8Array }> */
    function buildZipBlob(files) {
        const parts = [];
        const central = [];
        let offset = 0;
        for (const file of files) {
            const name = encodeText(file.name);
            const crc = crc32(file.data);
            const size = file.data.length;
            const local = new Uint8Array(30 + name.length);
            const lv = new DataView(local.buffer);
            lv.setUint32(0, 0x04034b50, true);
            lv.setUint16(4, 20, true);
            lv.setUint16(6, 0x0800, true); // UTF-8 文件名
            lv.setUint16(12, 0x0021, true);
            lv.setUint32(14, crc, true);
            lv.setUint32(18, size, true);
            lv.setUint32(22, size, true);
            lv.setUint16(26, name.length, true);
            local.set(name, 30);
            parts.push(local, file.data);

            const c = new Uint8Array(46 + name.length);
            const cv = new DataView(c.buffer);
            cv.setUint32(0, 0x02014b50, true);
            cv.setUint16(4, 20, true);
            cv.setUint16(6, 20, true);
            cv.setUint16(8, 0x0800, true);
            cv.setUint16(14, 0x0021, true);
            cv.setUint32(16, crc, true);
            cv.setUint32(20, size, true);
            cv.setUint32(24, size, true);
            cv.setUint16(28, name.length, true);
            cv.setUint32(42, offset, true);
            c.set(name, 46);
            central.push(c);
            offset += local.length + size;
        }
        const centralSize = central.reduce((n, c) => n + c.length, 0);
        const eocd = new Uint8Array(22);
        const ev = new DataView(eocd.buffer);
        ev.setUint32(0, 0x06054b50, true);
        ev.setUint16(8, files.length, true);
        ev.setUint16(10, files.length, true);
        ev.setUint32(12, centralSize, true);
        ev.setUint32(16, offset, true);
        return new Blob([...parts, ...central, eocd], { type: 'application/zip' });
    }

    // ---- 获取图片和附件 -------------------------------------------------------
    // 图片优先取原图（data-gallery-src），取不到再用页面上的缩略图（blob:）。
    // blob: 地址可能会被 Teams 释放，所以在收集时就先读进来

    const assetCache = new Map(); // URL -> Promise<Blob | null>
    const isHtmlBlob = b => /text\/html/i.test(b.type || '');

    function gmFetch(url) {
        if (typeof GM_xmlhttpRequest !== 'function') return Promise.resolve(null);
        return new Promise(resolve => GM_xmlhttpRequest({
            method: 'GET', url, responseType: 'blob', timeout: 30000,
            onload: r => resolve(r.status >= 200 && r.status < 300 && r.response && !isHtmlBlob(r.response) ? r.response : null),
            onerror: () => resolve(null),
            ontimeout: () => resolve(null),
        }));
    }

    function isAllowedAsset(url) {
        if (/^(blob|data):/.test(url)) return true;
        try {
            const u = new URL(url);
            return u.protocol === 'https:' && ASSET_HOST_RE.test(u.hostname);
        } catch (_) {
            return false;
        }
    }

    async function loadAsset(url) {
        if (!isAllowedAsset(url)) return null;
        try {
            if (/^(blob|data):/.test(url)) {
                const r = await fetch(url);
                return r.ok ? await r.blob() : null;
            }
            try {
                const r = await fetch(url, { credentials: 'include' });
                if (r.ok) {
                    const b = await r.blob();
                    if (!isHtmlBlob(b)) return b;
                }
            } catch (_) { /* 因 CORS 等原因失败时，改用 GM_xmlhttpRequest 重试 */ }
            return await gmFetch(url);
        } catch (e) {
            console.warn('[TeamsExporter] 取得失敗', url, e);
            return null;
        }
    }

    function fetchAsset(url) {
        if (!assetCache.has(url)) {
            // 失败可能只是暂时的，所以不缓存失败结果，下次导出时会重试
            const pending = loadAsset(url).then(b => { if (!b) assetCache.delete(url); return b; });
            assetCache.set(url, pending);
        }
        return assetCache.get(url);
    }

    async function firstAsset(urls) {
        for (const u of urls) {
            const b = await fetchAsset(u);
            if (b) return b;
        }
        return null;
    }

    /** SharePoint 的共享链接加上 download=1 就会返回文件本体 */
    function downloadUrl(url) {
        if (!/sharepoint\.com/i.test(url)) return url;
        try {
            const u = new URL(url);
            u.searchParams.set('download', '1');
            return u.toString();
        } catch (_) {
            return url;
        }
    }

    function imageExt(blob, url) {
        if (IMAGE_EXT[blob.type]) return IMAGE_EXT[blob.type];
        const m = /\.(png|jpe?g|gif|webp|svg|bmp)(?:[?#]|$)/i.exec(url || '');
        return m ? m[1].toLowerCase().replace('jpeg', 'jpg') : 'png';
    }

    const blobToDataUrl = blob => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(blob);
    });

    // ---- 提取消息 --------------------------------------------------
    // 不直接修改 Teams 的 DOM，一律在副本上处理

    function cleanBody(bodyEl, id) {
        const clone = bodyEl.cloneNode(true);
        clone.querySelectorAll('img[itemtype*="Emoji"]').forEach(img =>
            img.replaceWith(document.createTextNode(img.alt || '')));
        clone.querySelectorAll('[aria-label="Animated GIF"]').forEach(el =>
            el.replaceWith(document.createTextNode('[GIF]')));
        clone.querySelectorAll('div[aria-label*="Mention"]').forEach(div => {
            const span = document.createElement('span');
            span.className = 'mention';
            span.append(...div.childNodes);
            div.replaceWith(span);
        });
        clone.querySelectorAll('div[data-track-module-name="messageQuotedReply"]').forEach(div => {
            const quote = document.createElement('blockquote');
            quote.append(...div.childNodes);
            div.replaceWith(quote);
        });
        // 把 src 换成占位标记，方便导出时替换成实际图片
        const images = [];
        clone.querySelectorAll('img').forEach((img, i) => {
            const key = `${id}-${i}`;
            const urls = [img.getAttribute('data-gallery-src'), img.src].filter(Boolean);
            urls.filter(u => u.startsWith('blob:')).forEach(fetchAsset);
            images.push({ key, urls, alt: img.alt || '' });
            // 原图地址等导出文件里用不到的引用一律去掉
            ['src', 'srcset', 'data-gallery-src'].forEach(a => img.removeAttribute(a));
            img.setAttribute('data-tce-img', key);
        });
        return { el: clone, images };
    }

    function fileNameFromUrl(url) {
        try {
            const u = new URL(url);
            return decodeURIComponent(u.searchParams.get('file') || u.pathname.split('/').pop() || '') || 'file';
        } catch (_) {
            return 'file';
        }
    }

    function extractFiles(node) {
        const found = new Map(); // 链接（没有链接时用 name:文件名）-> { name, url }
        const add = (name, url) => {
            if (name && !found.has(url || `name:${name}`)) found.set(url || `name:${name}`, { name, url });
        };
        // 附件卡片：卡片内部元素的 aria-label 是“文件名\n共享链接”
        // （外层 file-attachment-grid 的 aria-label 只是“The message has an attachment.”这类说明，不能当文件名）
        const grid = node.querySelector('[data-tid="file-attachment-grid"]');
        if (grid) {
            grid.querySelectorAll('[aria-label]').forEach(el => {
                const m = ATTACHMENT_LABEL_RE.exec(el.getAttribute('aria-label'));
                if (m) add(m[1].trim(), m[2]);
            });
            // 取不到链接的卡片，至少保留文件名（role=group 的 aria-label 就是文件名）
            grid.querySelectorAll('[role="group"][aria-label]').forEach(el => {
                const name = el.getAttribute('aria-label').trim();
                if (![...found.values()].some(f => f.name === name)) add(name, '');
            });
        }
        // 正文里贴的文件链接
        node.querySelectorAll('[id^="content-"] a[href]').forEach(a => {
            if (!FILE_LINK_RE.test(a.href)) return;
            const name = (a.getAttribute('title') || a.textContent || '').trim();
            add(name || fileNameFromUrl(a.href), a.href);
        });
        return [...found.values()];
    }

    const MD_BLOCK = new Set(['p', 'div', 'section', 'tr']);

    function mdTable(table, walk) {
        const rows = [...table.querySelectorAll('tr')].map(tr =>
            [...tr.children].map(td => walk(td).replace(/\s*\n\s*/g, ' ').replace(/\|/g, '\\|').trim()));
        if (!rows.length) return '';
        const width = Math.max(...rows.map(r => r.length));
        const line = r => `| ${Array.from({ length: width }, (_, i) => r[i] || '').join(' | ')} |`;
        return `\n${[line(rows[0]), line(Array(width).fill('---')), ...rows.slice(1).map(line)].join('\n')}\n`;
    }

    /** 把正文 DOM 转成 Markdown。图片先输出为 IMG_MARK 包裹的键，导出时再替换 */
    function toMarkdown(root) {
        const walk = n => {
            // 转义正文，避免被当成 Markdown 语法或 HTML
            if (n.nodeType === Node.TEXT_NODE) return mdText(n.textContent.replace(/\s+/g, ' '));
            if (n.nodeType !== Node.ELEMENT_NODE) return '';
            const tag = n.tagName.toLowerCase();
            const inner = () => [...n.childNodes].map(walk).join('');
            const wrap = mark => { const t = inner().trim(); return t ? `${mark}${t}${mark}` : ''; };
            if (MD_BLOCK.has(tag)) return `\n${inner()}\n`;
            switch (tag) {
                case 'br': return '\n';
                case 'strong': case 'b': return wrap('**');
                case 'em': case 'i': return wrap('*');
                case 's': case 'del': case 'strike': return wrap('~~');
                case 'pre': return `\n\`\`\`\n${n.textContent.replace(/```/g, '`\u200b``').replace(/\n$/, '')}\n\`\`\`\n`;
                case 'code': return `\`${n.textContent.replace(/`/g, '\u200b`')}\``;
                case 'a': {
                    const text = inner().trim();
                    const href = mdUrl(n.getAttribute('href'));
                    return href && text ? `[${text}](<${href}>)` : text;
                }
                case 'img': {
                    const key = n.getAttribute('data-tce-img');
                    return key ? `![${mdText(n.alt || '图片')}](${IMG_MARK}${key}${IMG_MARK})` : mdText(n.alt || '');
                }
                case 'ul': case 'ol':
                    return `\n${[...n.children].map((li, i) =>
                        `${tag === 'ol' ? `${i + 1}.` : '-'} ${walk(li).trim().replace(/\n+/g, '\n  ')}`).join('\n')}\n`;
                case 'blockquote':
                    return `\n${inner().trim().split('\n').map(l => `> ${l}`).join('\n')}\n`;
                case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
                    return `\n**${inner().trim()}**\n`;
                case 'table': return mdTable(n, walk);
                default: return inner();
            }
        };
        return walk(root).replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    }

    /** 只在正文外面找，避免取到引用回复里的发言人和时间 */
    const outside = (node, selector, bodyWrap) =>
        [...node.querySelectorAll(selector)].find(el => !bodyWrap.contains(el)) || null;

    function snapshot(node) {
        const bodyWrap = node.querySelector('[id^="message-body-"]');
        if (!bodyWrap) return null;
        const id = parseInt(bodyWrap.id.slice('message-body-'.length), 10);
        if (!Number.isFinite(id)) return null;
        const bodyEl = bodyWrap.querySelector('[id^="content-"]') || bodyWrap;
        const authorEl = outside(node, '[data-tid="message-author-name"]', bodyWrap);
        const timeEl = outside(node, '[id^="timestamp-"]', bodyWrap);
        const time = timeEl ? new Date(timeEl.getAttribute('datetime')) : null;
        const body = cleanBody(bodyEl, id);
        return {
            id,
            author: authorEl ? authorEl.textContent.trim() : '',
            time: time && !isNaN(time) ? time : null,
            html: body.el.innerHTML,
            md: toMarkdown(body.el),
            images: body.images,
            files: extractFiles(node),
        };
    }

    // ---- 收集 -----------------------------------------------------------------

    /** 识别当前打开的对话。ID 取自对话列表里选中的项 */
    function currentChat() {
        const header = document.querySelector('[data-tid="chat-title"]');
        const fromDoc = document.title.replace(/^\(\d+\)\s*/, '').split('|').map(s => s.trim())
            .filter(p => p && !/^(Microsoft Teams|Teams|Chat|チャット|聊天)$/i.test(p))[0];
        const title = (header && header.textContent.trim()) || fromDoc || 'Teams';
        const item = document.querySelector('[role="treeitem"][tabindex="0"] [id^="title-chat-list-item_"]');
        const id = item && item.textContent.trim() === title ? item.id.slice('title-chat-list-item_'.length) : '';
        return { id, title };
    }

    /** 两边都有 ID 时比较 ID，任一边没有 ID 时比较标题 */
    const sameChat = (a, b) => (a.id && b.id ? a.id === b.id : a.title === b.title);
    const chatKey = c => c.id || c.title;

    /** 收集当前显示的消息，返回新增的条数 */
    function collectVisible() {
        const list = document.getElementById(LIST_ID);
        if (!list) return 0;
        let added = 0;
        for (const child of list.children) {
            const msg = snapshot(child);
            if (!msg) continue;
            const prev = state.messages.get(msg.id);
            if (!prev) added++;
            // 重新渲染后发言人或时间缺失时，保留之前的值
            state.messages.set(msg.id, {
                ...msg,
                author: msg.author || (prev && prev.author) || '',
                time: msg.time || (prev && prev.time) || null,
                run: state.collectRun,
            });
        }
        return added;
    }

    /** 对话已切换时停止收集。可以继续时返回 true */
    function ensureSameChat() {
        if (sameChat(currentChat(), state.chat)) return true;
        stopCollect();
        setNotice('已切换到其他对话，收集已停止', true);
        return false;
    }

    function observeList(list) {
        if (state.observer) state.observer.disconnect();
        state.observedList = list;
        // 即使变化不断，也按固定间隔收集（用防抖可能会一直被推迟）
        state.observer = new MutationObserver(() => {
            if (collectTimer) return;
            collectTimer = setTimeout(() => {
                collectTimer = null;
                if (state.collecting && ensureSameChat()) { collectVisible(); render(); }
            }, COLLECT_THROTTLE_MS);
        });
        state.observer.observe(list, { childList: true, subtree: true, characterData: true });
        collectVisible();
    }

    /** 开始收集。同一对话已收集的内容会保留（只有“重置”才会清空） */
    function startCollect() {
        const list = document.getElementById(LIST_ID);
        if (!list) {
            setNotice('请先打开一个对话', true);
            return false;
        }
        const chat = currentChat();
        if (!sameChat(chat, state.chat)) state.messages = new Map();
        state.chat = { id: chat.id || state.chat.id, title: chat.title };
        state.collecting = true;
        observeList(list);
        setNotice('向上滚动可收集更早的消息');
        return true;
    }

    function stopCollect() {
        if (state.observer) state.observer.disconnect();
        clearTimeout(collectTimer);
        collectTimer = null;
        state.observer = null;
        state.observedList = null;
        state.collecting = false;
        stopAutoScroll();
    }

    function stopAutoScroll() {
        state.scrollRun++;
        state.autoScrolling = false;
        render();
    }

    /** 找到包住消息列表、能上下滚动的容器 */
    function findScroller(node) {
        for (let n = node.parentElement; n && n !== document.body; n = n.parentElement) {
            const { overflowY } = getComputedStyle(n);
            if ((overflowY === 'auto' || overflowY === 'scroll') && n.scrollHeight > n.clientHeight) return n;
        }
        return null;
    }

    /** 指定 run 时，只统计该次自动获取中看到的消息 */
    function oldestTime(run) {
        let oldest = null;
        for (const m of state.messages.values()) {
            if (run !== undefined && m.run !== run) continue;
            if (m.time && (!oldest || m.time < oldest)) oldest = m.time;
        }
        return oldest;
    }

    /**
     * 滚动到当前最早的一条消息，让 Teams 继续加载更早的内容，直到开始日期。
     * 翻到开始日期（或对话开头）时返回 true，中止或失败时返回 false
     */
    async function autoScroll() {
        if (!state.collecting && !startCollect()) return false;
        const run = ++state.scrollRun;
        const alive = () => run === state.scrollRun && state.collecting;
        // 之前收集的部分中间可能有缺口，所以只用这次看到的范围来判断
        const collectRun = ++state.collectRun;
        collectVisible();
        state.autoScrolling = true;
        setNotice('正在自动向上翻…');
        let idle = 0;
        let result = { done: false, reason: '已达到最大次数，已停止' };
        for (let round = 0; round < AUTO_SCROLL.maxRounds; round++) {
            const { from } = readRange(); // 中途修改日期也会生效
            const oldest = oldestTime(collectRun);
            if (from && oldest && oldest < from) { result = { done: true, reason: '已取到开始日期' }; break; }
            if (idle >= AUTO_SCROLL.idleRoundsToStop) {
                // 也可能只是加载慢，没翻到开始日期时给出提醒
                result = from
                    ? { done: true, reason: '没有翻到开始日期之前的消息就停止了。如果加载较慢，可能有遗漏', warn: true }
                    : { done: true, reason: '已取到对话开头' };
                break;
            }
            const first = document.getElementById(LIST_ID)?.firstElementChild;
            if (!first) { result = { done: false, reason: '找不到消息列表，已停止' }; break; }
            // 已经在最顶端时滚动不会触发加载，先往下挪一点再滚回去
            const scroller = findScroller(first);
            if (scroller && scroller.scrollTop <= 1) {
                scroller.scrollTop = 60;
                await sleep(50);
            }
            first.scrollIntoView({ block: 'start' });
            await sleep(AUTO_SCROLL.waitMs);
            if (!alive()) return false;
            if (!ensureSameChat()) return false;
            idle = collectVisible() ? 0 : idle + 1;
            render();
        }
        if (run === state.scrollRun) state.autoScrolling = false;
        state.scrollWarning = result.warn ? result.reason : '';
        setNotice(result.reason, !result.done || !!result.warn);
        return result.done;
    }

    // ---- 时间范围 -----------------------------------------------------------------

    const startOfToday = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
    const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
    const startOfWeek = d => addDays(d, -((d.getDay() + 6) % 7)); // 一周从周一开始
    const isWeekend = d => d.getDay() === 0 || d.getDay() === 6;   // 只跳过周六日，节假日不处理

    /** 往前找最近的工作日（不含 d 当天） */
    function previousWorkday(d) {
        let x = addDays(d, -1);
        while (isWeekend(x)) x = addDays(x, -1);
        return x;
    }

    /** 往回数 n 个工作日（含 d 当天，若 d 是工作日）的起始日 */
    function workdaysBack(d, n) {
        let x = d;
        let count = isWeekend(d) ? 0 : 1;
        while (count < n) {
            x = addDays(x, -1);
            if (!isWeekend(x)) count++;
        }
        return x;
    }

    // Teams 以工作日为主：按工作日取范围，夹在中间的周末消息作为附加内容一并导出

    /** 快速模式：全部截止到今天，点一下直接导出 */
    const QUICK_PRESETS = [
        ['今天', () => { const t = startOfToday(); return [t, t]; }],
        ['上个工作日起', () => { const t = startOfToday(); return [previousWorkday(t), t]; }],
        ['本周', () => { const t = startOfToday(); return [startOfWeek(t), t]; }],
        ['最近5个工作日', () => { const t = startOfToday(); return [workdaysBack(t, 5), t]; }],
        ['最近一个月', () => { const t = startOfToday(); return [oneMonthBefore(t), t]; }],
        ['全部', () => [null, null]],
    ];

    /** 上个月的同一天。上个月没有这一天时（如 3/31）取上个月的最后一天 */
    function oneMonthBefore(d) {
        const lastDayPrev = new Date(d.getFullYear(), d.getMonth(), 0).getDate();
        return new Date(d.getFullYear(), d.getMonth() - 1, Math.min(d.getDate(), lastDayPrev));
    }

    /** 详细模式：填入完整的日期范围，之后还可以手动调整 */
    const DETAIL_PRESETS = [
        ['今天', () => { const t = startOfToday(); return [t, t]; }],
        ['上个工作日', () => { const t = startOfToday(); return [previousWorkday(t), addDays(t, -1)]; }],
        ['本周', () => { const s = startOfWeek(startOfToday()); return [s, addDays(s, 6)]; }],
        ['上周', () => { const s = addDays(startOfWeek(startOfToday()), -7); return [s, addDays(s, 6)]; }],
        ['最近5个工作日', () => { const t = startOfToday(); return [workdaysBack(t, 5), t]; }],
        ['上月', () => { const t = startOfToday(); return [new Date(t.getFullYear(), t.getMonth() - 1, 1), new Date(t.getFullYear(), t.getMonth(), 0)]; }],
        ['全部', () => [null, null]],
    ];

    function setRange([from, to]) {
        ui.from.value = from ? fmtDate(from) : '';
        ui.to.value = to ? fmtDate(to) : '';
        render();
    }

    function readRange() {
        const parse = (value, endOfDay) => {
            if (!value) return null;
            const [y, m, d] = value.split('-').map(Number);
            return endOfDay ? new Date(y, m - 1, d, 23, 59, 59, 999) : new Date(y, m - 1, d);
        };
        return {
            from: parse(ui && ui.from.value, false),
            to: parse(ui && ui.to.value, true),
        };
    }

    /** 按 ID 排序，缺失的发言人和时间用前一条消息补上 */
    function orderedMessages() {
        const sorted = [...state.messages.values()].sort((a, b) => a.id - b.id);
        let lastAuthor = '';
        let lastTime = null;
        return sorted.map(m => {
            const author = m.author || lastAuthor;
            const time = m.time || lastTime;
            lastAuthor = author;
            lastTime = time;
            return { ...m, author, time };
        });
    }

    function messagesInRange() {
        const { from, to } = readRange();
        const all = orderedMessages();
        if (!from && !to) return all;
        return all.filter(m => m.time && (!from || m.time >= from) && (!to || m.time <= to));
    }

    function describeSpan(msgs) {
        const times = msgs.map(m => m.time).filter(Boolean);
        return times.length ? { first: times[0], last: times[times.length - 1] } : null;
    }

    function chatForExport() {
        const msgs = messagesInRange();
        if (!msgs.length) return null;
        const span = describeSpan(msgs);
        const { from, to } = readRange();
        const title = state.chat.title || currentChat().title;
        return {
            key: chatKey(state.chat.title ? state.chat : currentChat()),
            title,
            msgs,
            meta: {
                title,
                count: msgs.length,
                span: span ? `${fmtDateTime(span.first)} 〜 ${fmtDateTime(span.last)}` : '未知',
                filter: from || to ? `${from ? fmtDate(from) : '最早'} 〜 ${to ? fmtDate(to) : '最新'}` : '未指定',
                exportedAt: fmtDateTime(new Date()),
                fileSpan: span ? `${compactDate(span.first)}-${compactDate(span.last)}` : 'unknown',
            },
        };
    }

    // ---- 生成文档 ---------------------------------------------------------
    // links: { image(key) -> 图片地址, file(file) -> 链接地址 }

    const HTML_STYLE = `
body { font-family: "Segoe UI", "Microsoft YaHei", "Yu Gothic UI", sans-serif; margin: 2em auto; max-width: 960px; padding: 0 1em; color: #222; }
header { border-bottom: 2px solid #5b5fc7; margin-bottom: 1em; padding-bottom: .5em; }
header dl { display: grid; grid-template-columns: auto 1fr; gap: .2em 1em; margin: 0; }
header dt { color: #666; }
h2.day { text-align: center; font-size: .95em; color: #666; border-top: 1px solid #ddd; padding-top: .8em; }
.author { font-weight: bold; margin-top: 1em; }
.msg { display: flex; gap: .8em; }
.time { color: #888; font-size: .85em; min-width: 3em; padding-top: .15em; }
section { flex: 1; min-width: 0; padding-left: .8em; border-left: 1px solid #ddd; }
blockquote { margin: .3em 0; padding: .3em .8em; background: #f4f4f8; border-left: 3px solid #bbb; }
.mention { font-weight: bold; }
.files { margin: .3em 0; padding-left: 1.2em; font-size: .9em; }
img { max-width: 100%; }
table { border-collapse: collapse; } td, th { border: 1px solid #ccc; padding: .3em .6em; }`;

    function htmlPage(title, headerHtml, bodyHtml) {
        return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: blob: https:; style-src 'unsafe-inline'">
<title>${esc(title)}</title><style>${HTML_STYLE}</style></head><body>
<header><h1>${esc(title)}</h1>${headerHtml}</header>
${bodyHtml}
</body></html>`;
    }

    function messageHtml(m, links) {
        const body = m.html.replace(/data-tce-img="([^"]+)"/g, (_, key) => `src="${esc(links.image(key))}"`);
        const files = m.files.length
            ? `<ul class="files">${m.files.map(f => {
                const href = links.file(f);
                return `<li>📎 ${href ? `<a href="${esc(href)}">${esc(f.name)}</a>` : esc(f.name)}</li>`;
            }).join('')}</ul>`
            : '';
        return `<section>${body}${files}</section>`;
    }

    /** 依次排列消息，同时判断日期分隔和发言人切换 */
    function eachWithHeadings(msgs, fn) {
        let lastDay = null;
        let lastAuthor = null;
        return msgs.map(m => {
            const day = m.time ? fmtDay(m.time) : '日期未知';
            const newDay = day !== lastDay;
            const newAuthor = newDay || m.author !== lastAuthor;
            lastDay = day;
            lastAuthor = m.author;
            return fn(m, { day, newDay, newAuthor, time: m.time ? fmtTime(m.time) : '--:--' });
        });
    }

    function buildHtml(chat, links) {
        const { meta } = chat;
        const body = eachWithHeadings(chat.msgs, (m, h) =>
            (h.newDay ? `<h2 class="day">${esc(h.day)}</h2>` : '')
            + (h.newAuthor ? `<div class="author">${esc(m.author || '（未知）')}</div>` : '')
            + `<div class="msg"><span class="time">${h.time}</span>${messageHtml(m, links)}</div>`).join('\n');
        const header = `<dl>
<dt>消息数</dt><dd>${meta.count}</dd>
<dt>时间范围</dt><dd>${esc(meta.span)}</dd>
<dt>筛选</dt><dd>${esc(meta.filter)}</dd>
<dt>导出时间</dt><dd>${esc(meta.exportedAt)}</dd>
</dl>`;
        return htmlPage(meta.title, header, body);
    }

    function buildMarkdown(chat, links) {
        const { meta } = chat;
        const head = [
            `# ${mdText(meta.title)}`, '',
            `- 消息数：${meta.count}`,
            `- 时间范围：${meta.span}`,
            `- 筛选：${meta.filter}`,
            `- 导出时间：${meta.exportedAt}`,
        ];
        const body = eachWithHeadings(chat.msgs, (m, h) => {
            const lines = [];
            if (h.newDay) lines.push(`## ${h.day}`, '');
            if (h.newAuthor) lines.push(`### ${mdText(m.author || '（未知）')}`, '');
            // 用 <...> 包裹后，含空格或括号的路径、地址也能直接作为链接
            const markRe = new RegExp(`\\(${IMG_MARK}([^${IMG_MARK}]+)${IMG_MARK}\\)`, 'g');
            const text = m.md.replace(markRe, (_, key) => `(<${mdUrl(links.image(key))}>)`);
            // 只有附件的消息不写“无正文”，直接列出附件
            lines.push(`\`${h.time}\``);
            if (text || !m.files.length) lines.push(text || '（无正文）');
            for (const f of m.files) {
                const href = mdUrl(links.file(f));
                lines.push(href ? `📎 [${mdText(f.name)}](<${href}>)` : `📎 ${mdText(f.name)}`);
            }
            return lines.join('\n');
        });
        return `${head.join('\n')}\n\n---\n\n${body.join('\n\n')}\n`;
    }

    function buildIndexHtml(entries) {
        const rows = entries.map(e => `<tr><td><a href="${esc(encodeURI(e.path))}">${esc(e.meta.title)}</a></td>`
            + `<td>${e.meta.count}</td><td>${esc(e.meta.span)}</td><td>${esc(e.statsText)}</td></tr>`).join('\n');
        const table = `<table><tr><th>对话</th><th>消息数</th><th>时间范围</th><th>图片 / 附件</th></tr>\n${rows}</table>`;
        return htmlPage('Teams 聊天导出', `<p>导出时间：${esc(fmtDateTime(new Date()))}</p>`, table);
    }

    // ---- 导出 ---------------------------------------------------------------

    const fallbackUrl = img => img.urls.find(u => !u.startsWith('blob:')) || '';

    /** 获取每张图片，用 toSrc 转成引用地址；取不到时保留原地址 */
    async function resolveImages(msgs, toSrc, onProgress) {
        const images = msgs.flatMap(m => m.images);
        const srcByKey = new Map();
        let ok = 0;
        let done = 0;
        await mapPool(images, FETCH_CONCURRENCY, async (img, i) => {
            const blob = await firstAsset(img.urls);
            const src = blob ? await toSrc(blob, img, i) : null;
            if (src) ok++;
            srcByKey.set(img.key, src || fallbackUrl(img));
            onProgress(`正在获取图片 ${++done}/${images.length}`);
        });
        return { srcByKey, total: images.length, ok };
    }

    async function resolveFiles(msgs, toHref, onProgress) {
        const files = [...new Map(msgs.flatMap(m => m.files).filter(f => f.url).map(f => [f.url, f])).values()];
        const hrefByUrl = new Map();
        let done = 0;
        await mapPool(files, FETCH_CONCURRENCY, async f => {
            const blob = await fetchAsset(downloadUrl(f.url));
            if (blob) hrefByUrl.set(f.url, await toHref(blob, f));
            onProgress(`正在获取附件 ${++done}/${files.length}`);
        });
        return { hrefByUrl, total: files.length, ok: hrefByUrl.size };
    }

    const statsText = (imgs, files) => `图片 ${imgs.ok}/${imgs.total}`
        + (files ? `，附件 ${files.ok}/${files.total}` : '');

    /** HTML：用 data URL 内嵌图片的单个文件 */
    async function renderHtmlFile(chat, onProgress) {
        const imgs = await resolveImages(chat.msgs, blob => blobToDataUrl(blob), onProgress);
        const links = { image: key => imgs.srcByKey.get(key) || '', file: f => f.url };
        return { text: buildHtml(chat, links), stats: statsText(imgs) };
    }

    /** Markdown：在 dir（ZIP 内的文件夹，以 / 结尾）下放 chat.md、images/ 和 files/ */
    async function renderMarkdownFiles(chat, dir, withFiles, onProgress) {
        const used = new Set();
        const out = [];
        const put = async (path, blob) => {
            out.push({ name: dir + path, data: new Uint8Array(await blob.arrayBuffer()) });
            return path;
        };
        const imgs = await resolveImages(chat.msgs, (blob, img, i) =>
            put(uniqueName(used, `images/${pad(i + 1, 3)}.${imageExt(blob, img.urls[0])}`), blob), onProgress);
        const files = withFiles
            ? await resolveFiles(chat.msgs, (blob, f) => put(uniqueName(used, `files/${safeName(f.name, 100)}`), blob), onProgress)
            : null;
        const links = {
            image: key => imgs.srcByKey.get(key) || '',
            file: f => (files && files.hrefByUrl.get(f.url)) || f.url,
        };
        out.push({ name: `${dir}chat.md`, data: encodeText(buildMarkdown(chat, links)) });
        return { files: out, stats: statsText(imgs, files) };
    }

    function saveBlob(blob, fileName) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 10000);
    }

    /** 保存一个对话：HTML 单文件，或 Markdown ZIP。返回完成提示 */
    async function saveOne(chat, onProgress) {
        const base = `Teams_${safeName(chat.title)}_${chat.meta.fileSpan}`;
        if (ui.format.value === 'md') {
            const r = await renderMarkdownFiles(chat, '', ui.withFiles.checked, onProgress);
            saveBlob(buildZipBlob(r.files), `${base}.zip`);
            return `已保存：${base}.zip（${chat.msgs.length} 条，${r.stats}）`;
        }
        const r = await renderHtmlFile(chat, onProgress);
        saveBlob(new Blob([r.text], { type: 'text/html;charset=utf-8' }), `${base}.html`);
        return `已保存：${base}.html（${chat.msgs.length} 条，${r.stats}）`;
    }

    /** 一个对话时存单个文件，多个对话时合并成一个 ZIP（附带 index.html 目录）。返回完成提示 */
    async function saveChats(chats, onProgress) {
        if (chats.length === 1) return saveOne(chats[0], onProgress);
        const all = [];
        const entries = [];
        for (const [i, chat] of chats.entries()) {
            const name = `${pad(i + 1)}_${safeName(chat.title, 40)}`;
            const progress = text => onProgress(`(${i + 1}/${chats.length}) ${chat.title}：${text}`);
            if (ui.format.value === 'md') {
                const r = await renderMarkdownFiles(chat, `${name}/`, ui.withFiles.checked, progress);
                all.push(...r.files);
                entries.push({ path: `${name}/chat.md`, meta: chat.meta, statsText: r.stats });
            } else {
                const r = await renderHtmlFile(chat, progress);
                all.push({ name: `${name}.html`, data: encodeText(r.text) });
                entries.push({ path: `${name}.html`, meta: chat.meta, statsText: r.stats });
            }
        }
        all.unshift({ name: 'index.html', data: encodeText(buildIndexHtml(entries)) });
        const fileName = `Teams_合并_${chats.length}个对话_${stamp(new Date())}.zip`;
        saveBlob(buildZipBlob(all), fileName);
        return `已保存：${fileName}`;
    }

    /** 快速模式：确定时间范围后，自动翻到开始日期并直接保存当前对话 */
    async function quickExport(preset) {
        setRange(preset());
        if (!(await autoScroll())) return;
        const warning = state.scrollWarning;
        stopCollect();
        const chat = chatForExport();
        if (!chat) return setNotice('所选时间范围内没有消息', true);
        const done = await saveChats([chat], text => setNotice(text));
        return setNotice(warning ? `${done}\n⚠ ${warning}` : done, !!warning);
    }

    // ---- 批量导出（详细模式） ------------------------------------------------

    const OPEN_CHAT = { waitMs: 250, maxWaits: 60, settleMs: 1000 };

    /** 读取左侧对话列表。列表只渲染可见部分，所以与之前读到的合并，不删除 */
    function loadChatList() {
        const found = new Map(state.chatList.map(c => [c.id, c]));
        document.querySelectorAll('[id^="title-chat-list-item_"]').forEach(span => {
            const id = span.id.slice('title-chat-list-item_'.length);
            if (!span.closest('[role="treeitem"]')) return;
            const time = document.getElementById(`time-chat-list-item_${id}`);
            found.set(id, { id, title: span.textContent.trim(), time: time ? time.textContent.trim() : '' });
        });
        state.chatList = [...found.values()];
        renderChatList();
        if (!state.chatList.length) setNotice('找不到对话列表，请先切到“聊天”', true);
        else render();
    }

    /** 当前打开的对话的特征：顶部标题和第一条消息的 ID */
    function openedChatSignature() {
        const first = document.getElementById(LIST_ID)?.querySelector('[id^="message-body-"]');
        return { title: currentChat().title, first: first ? first.id : '' };
    }

    /**
     * 点击左侧列表打开对话，等到切换完成为止。
     * 左侧列表的名字和顶部标题不一定相同（如群聊），所以用“标题和消息都变了”来判断已切换
     */
    async function openChat(target, alive) {
        const before = openedChatSignature();
        const item = document.getElementById(`title-chat-list-item_${target.id}`)?.closest('[role="treeitem"]');
        // 已经是这个对话（左侧列表里当前打开的一项 tabindex 为 0）
        if (before.title === target.title || (item && item.getAttribute('tabindex') === '0')) return;
        if (!item) throw new Error('左侧列表里找不到，请滚动左侧列表后刷新');
        item.scrollIntoView({ block: 'nearest' });
        item.click();
        for (let i = 0; i < OPEN_CHAT.maxWaits; i++) {
            await sleep(OPEN_CHAT.waitMs);
            if (!alive()) return;
            const now = openedChatSignature();
            if (document.getElementById(LIST_ID) && now.title !== before.title && now.first !== before.first) {
                await sleep(OPEN_CHAT.settleMs); // 等消息渲染出来
                return;
            }
        }
        throw new Error('打开超时');
    }

    /** 详细模式：依次打开勾选的对话，各自翻到开始日期后收集，最后一起保存 */
    async function exportSelected() {
        const targets = state.chatList.filter(c => state.selected.has(c.id));
        if (!targets.length) return setNotice('请先勾选对话', true);
        const batch = ++state.batchRun;
        const alive = () => batch === state.batchRun;
        const results = [];
        const warnings = [];
        const failed = [];
        let empty = 0;
        try {
            for (const [i, target] of targets.entries()) {
                state.batchLabel = `${i + 1}/${targets.length} · ${target.title}`;
                // 单个对话出错时跳过，继续处理后面的对话
                try {
                    await openChat(target, alive);
                    if (!alive()) return setNotice('已中止');
                    state.messages = new Map();
                    state.chat = { id: '', title: '' };
                    const done = await autoScroll();
                    if (!alive()) return setNotice('已中止');
                    if (!done) throw new Error(state.notice.text);
                    if (state.scrollWarning) warnings.push(target.title);
                    stopCollect();
                    const chat = chatForExport();
                    if (chat) results.push(chat); else empty++;
                } catch (e) {
                    stopCollect();
                    failed.push(`${target.title}（${errorText(e)}）`);
                }
            }
        } finally {
            state.batchLabel = '';
            stopCollect();
        }
        const extra = [
            empty ? `${empty} 个对话在时间范围内没有消息` : '',
            failed.length ? `⚠ 已跳过：${failed.join('、')}` : '',
            warnings.length ? `⚠ 可能有遗漏：${warnings.join('、')}` : '',
        ].filter(Boolean);
        if (!results.length) return setNotice(['所选时间范围内没有消息', ...extra].join('\n'), true);
        const done = await saveChats(results, text => setNotice(text));
        return setNotice([done, ...extra].join('\n'), failed.length + warnings.length > 0);
    }

    function abortAll() {
        state.batchRun++;
        stopCollect();
        setNotice('已中止');
    }

    /** 导出类操作同一时间只执行一个，失败时在界面上提示 */
    async function runExclusive(task) {
        if (state.busy) return;
        state.busy = true;
        render();
        try {
            await task();
        } catch (e) {
            console.error('[TeamsExporter]', e);
            setNotice(`失败：${errorText(e)}`, true);
        } finally {
            state.busy = false;
            render();
        }
    }

    // ---- UI -------------------------------------------------------------------

    function el(tag, css, text) {
        const node = document.createElement(tag);
        if (css) node.style.cssText = css;
        if (text !== undefined) node.textContent = text;
        return node;
    }

    // 使用 Teams（Fluent UI）的配色变量，浅色和深色主题下都能融入
    const THEME = {
        bg: 'var(--colorNeutralBackground1, #fff)',
        bgSubtle: 'var(--colorNeutralBackground3, #f5f5f5)',
        bgHover: 'var(--colorSubtleBackgroundHover, rgba(0,0,0,.05))',
        fg: 'var(--colorNeutralForeground1, #242424)',
        fgSub: 'var(--colorNeutralForeground2, #424242)',
        fgHint: 'var(--colorNeutralForeground3, #616161)',
        stroke: 'var(--colorNeutralStroke1, #d1d1d1)',
        divider: 'var(--colorNeutralStroke2, #e0e0e0)',
        brand: 'var(--colorBrandBackground, #5b5fc7)',
        brandFg: 'var(--colorBrandForeground1, #5b5fc7)',
        onBrand: 'var(--colorNeutralForegroundOnBrand, #fff)',
        danger: 'var(--colorPaletteRedForeground1, #c62828)',
    };
    const RAIL_WIDTH = 68; // 左侧应用栏的宽度

    const buttonColors = primary => (primary
        ? { background: THEME.brand, color: THEME.onBrand, borderColor: THEME.brand }
        : { background: THEME.bg, color: THEME.fg, borderColor: THEME.stroke });

    function button(text, primary, onClick) {
        const b = el('button', 'padding:5px 10px;border-radius:4px;cursor:pointer;font:inherit;border:1px solid;', text);
        Object.assign(b.style, buttonColors(primary));
        b.addEventListener('click', onClick);
        return b;
    }

    const row = (css = '') => el('div', `display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:8px;${css}`);
    const fieldCss = `font:inherit;color:inherit;background:${THEME.bg};border:1px solid ${THEME.stroke};border-radius:4px;padding:2px 4px;`;

    /** 与应用栏其他图标同样线条风格的“导出”图标 */
    function exportIcon() {
        const ns = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(ns, 'svg');
        [['width', '24'], ['height', '24'], ['viewBox', '0 0 24 24'], ['fill', 'none'], ['stroke', 'currentColor'],
            ['stroke-width', '1.5'], ['stroke-linecap', 'round'], ['stroke-linejoin', 'round'], ['aria-hidden', 'true']]
            .forEach(([k, v]) => svg.setAttribute(k, v));
        ['M12 3.5v10.5', 'M7.5 9.5 12 14l4.5-4.5', 'M4.5 16v2A2.5 2.5 0 0 0 7 20.5h10a2.5 2.5 0 0 0 2.5-2.5v-2'].forEach(d => {
            const path = document.createElementNS(ns, 'path');
            path.setAttribute('d', d);
            svg.appendChild(path);
        });
        return svg;
    }

    /** 放在左下角（应用栏最下方）的启动按钮。不插入 Teams 的 DOM，而是叠加显示 */
    function buildLauncher() {
        const b = el('button', `position:fixed;left:0;bottom:12px;width:${RAIL_WIDTH}px;height:48px;z-index:10000;`
            + 'display:flex;align-items:center;justify-content:center;background:transparent;border:none;border-radius:4px;cursor:pointer;');
        b.title = '导出对话';
        b.setAttribute('aria-label', b.title);
        const badge = el('span', `position:absolute;top:10px;right:20px;width:8px;height:8px;border-radius:50%;background:${THEME.brandFg};display:none;`);
        b.append(exportIcon(), badge);
        b.addEventListener('mouseenter', () => { b.style.background = THEME.bgHover; });
        b.addEventListener('mouseleave', () => { b.style.background = 'transparent'; });
        b.addEventListener('click', () => togglePanel(ui.panel.style.display === 'none'));
        return { launcher: b, badge };
    }

    function setNotice(text, isError = false) {
        state.notice = { text, isError };
        render();
    }

    function presetButtons(presets, onPick) {
        const r = row();
        for (const [label, range] of presets) {
            const b = button(label, false, () => onPick(range));
            b.style.padding = '3px 9px';
            r.appendChild(b);
        }
        return r;
    }

    function buildQuickSection() {
        const box = el('div');
        box.append(presetButtons(QUICK_PRESETS, range => runExclusive(() => quickExport(range))));
        return box;
    }

    function buildDetailSection() {
        const box = el('div');
        const from = el('input', fieldCss);
        const to = el('input', fieldCss);
        from.type = 'date';
        to.type = 'date';
        from.addEventListener('change', render);
        to.addEventListener('change', render);
        const dates = row();
        dates.append(from, el('span', '', '–'), to);

        const refreshBtn = button('刷新列表', false, loadChatList);
        const selectAll = el('input');
        selectAll.type = 'checkbox';
        selectAll.addEventListener('change', () => {
            state.selected = selectAll.checked ? new Set(state.chatList.map(c => c.id)) : new Set();
            renderChatList();
            render();
        });
        const selectAllLabel = el('label', 'display:flex;align-items:center;gap:4px;cursor:pointer;margin-left:auto;');
        selectAllLabel.append(selectAll, '全选');
        const listTools = row();
        listTools.append(refreshBtn, selectAllLabel);
        const chatList = el('div', `max-height:240px;overflow:auto;border:1px solid ${THEME.divider};border-radius:4px;margin-bottom:8px;`);
        const exportBtn = button('', true, () => runExclusive(exportSelected));

        box.append(presetButtons(DETAIL_PRESETS, range => setRange(range())), dates, listTools, chatList, exportBtn);
        return { box, from, to, selectAll, chatList, exportBtn };
    }

    /** 重建对话清单（只在读取列表或全选时调用，避免定时刷新打断勾选） */
    function renderChatList() {
        if (!ui) return;
        ui.chatList.replaceChildren(...state.chatList.map(c => {
            const line = el('label', 'display:flex;align-items:center;gap:6px;padding:4px 8px;cursor:pointer;');
            const box = el('input');
            box.type = 'checkbox';
            box.checked = state.selected.has(c.id);
            box.addEventListener('change', () => {
                const next = new Set(state.selected);
                if (box.checked) next.add(c.id); else next.delete(c.id);
                state.selected = next;
                render();
            });
            const name = el('span', 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;', c.title);
            name.title = c.title;
            line.append(box, name, el('span', `color:${THEME.fgHint};font-size:12px;`, c.time));
            return line;
        }));
    }

    function buildPanel(prefs) {
        // 在启动按钮右侧（紧挨应用栏）弹出，底边对齐
        const panel = el('div', `display:none;position:fixed;left:${RAIL_WIDTH + 8}px;bottom:12px;z-index:10000;`
            + 'width:350px;max-height:calc(100vh - 80px);overflow:auto;padding:12px;border-radius:8px;'
            + `background:${THEME.bg};color:${THEME.fg};border:1px solid ${THEME.divider};`
            + 'box-shadow:var(--shadow16, 0 8px 16px rgba(0,0,0,.14));font:13px/1.5 "Segoe UI","Microsoft YaHei","Yu Gothic UI",sans-serif;');
        const header = row('justify-content:space-between;');
        const tabs = row('margin:0;gap:0;');
        const quickTab = button('快速', false, () => setMode('quick'));
        const detailTab = button('详细', false, () => setMode('detail'));
        quickTab.style.borderRadius = '4px 0 0 4px';
        detailTab.style.borderRadius = '0 4px 4px 0';
        tabs.append(quickTab, detailTab);
        header.append(tabs, button('✕', false, () => togglePanel(false)));

        const format = el('select', `${fieldCss}max-width:100%;`);
        for (const [value, f] of Object.entries(FORMATS)) {
            const opt = el('option', '', f.label);
            opt.value = value;
            format.appendChild(opt);
        }
        format.value = FORMATS[prefs.format] ? prefs.format : 'html';
        const withFiles = el('input');
        withFiles.type = 'checkbox';
        withFiles.checked = !!prefs.withFiles;
        const filesLabel = el('label', 'display:flex;align-items:center;gap:4px;cursor:pointer;');
        filesLabel.title = '同时打包附件文件，取不到时只保留文件名和链接';
        filesLabel.append(withFiles, '附件');
        [format, withFiles].forEach(c => c.addEventListener('change', () => { savePrefs(); render(); }));
        const formatRow = row();
        formatRow.append(format, filesLabel);

        const quick = buildQuickSection();
        const detail = buildDetailSection();
        const status = el('div', `background:${THEME.bgSubtle};border-radius:4px;padding:6px 8px;margin:8px 0;white-space:pre-line;`);
        const abortBtn = button('中止', false, abortAll);
        const notice = el('div', 'font-size:12px;word-break:break-all;white-space:pre-line;');

        panel.append(header, formatRow, quick, detail.box, status, abortBtn, notice);
        return { panel, quickTab, detailTab, format, withFiles, filesLabel, quick, status, abortBtn, notice, ...detail, detailBox: detail.box };
    }

    function setMode(mode) {
        ui.mode = mode;
        savePrefs();
        if (mode === 'detail') loadChatList();
        render();
    }

    function togglePanel(open) {
        ui.panel.style.display = open ? 'block' : 'none';
        if (open && ui.mode === 'detail') loadChatList();
        render();
    }

    function createUI() {
        const prefs = loadPrefs();
        const root = el('div'); // 子元素各自 fixed 定位，root 只用来切换显示和隐藏
        root.id = UI_ID;
        const { launcher, badge } = buildLauncher();
        const parts = buildPanel(prefs);
        root.append(launcher, parts.panel);
        document.body.appendChild(root);
        ui = { root, launcher, badge, mode: prefs.mode === 'detail' ? 'detail' : 'quick', ...parts };
        render();
    }

    /** 状态栏：对话名、条数、时间范围。没有内容时返回空串（状态栏隐藏） */
    function statusText() {
        const head = state.batchLabel ? [state.batchLabel] : [];
        if (!state.messages.size) return [...head, state.collecting ? `${state.chat.title} · 0 条` : ''].filter(Boolean).join('\n');
        const all = orderedMessages();
        const span = describeSpan(all);
        const inRange = messagesInRange();
        const { from, to } = readRange();
        const count = from || to ? `${inRange.length} / ${all.length} 条` : `${all.length} 条`;
        const lines = [...head, `${state.chat.title} · ${count}`];
        if (span) lines.push(`${fmtDateTime(span.first)} – ${fmtDateTime(span.last)}`);
        const images = inRange.reduce((n, m) => n + m.images.length, 0);
        const files = inRange.reduce((n, m) => n + m.files.length, 0);
        if (images || files) lines.push(`图片 ${images} · 附件 ${files}`);
        return lines.join('\n');
    }

    function render() {
        if (!ui) return;
        const quickMode = ui.mode === 'quick';
        const { from, to } = readRange();
        const badRange = from && to && from > to;
        [[ui.quickTab, quickMode], [ui.detailTab, !quickMode]].forEach(([tab, active]) => {
            Object.assign(tab.style, buttonColors(active));
        });
        ui.quick.style.display = quickMode ? 'block' : 'none';
        ui.detailBox.style.display = quickMode ? 'none' : 'block';
        ui.filesLabel.style.display = ui.format.value === 'md' ? 'flex' : 'none';
        ui.quick.querySelectorAll('button').forEach(b => { b.disabled = state.busy; });
        const status = statusText();
        ui.status.textContent = status;
        ui.status.style.display = status ? 'block' : 'none';
        ui.abortBtn.style.display = state.autoScrolling || state.batchLabel ? 'inline-block' : 'none';
        const selectedCount = state.chatList.filter(c => state.selected.has(c.id)).length;
        ui.chatList.style.display = state.chatList.length ? 'block' : 'none';
        ui.selectAll.checked = selectedCount > 0 && selectedCount === state.chatList.length;
        ui.exportBtn.textContent = `导出 ${selectedCount} 个对话`;
        ui.exportBtn.disabled = state.busy || badRange || !selectedCount;
        const notice = badRange ? { text: '开始日期晚于结束日期', isError: true } : state.notice;
        ui.notice.textContent = notice.text;
        ui.notice.style.color = notice.isError ? THEME.danger : THEME.fgSub;
        // 面板打开或有任务时图标用主题色，有任务时再显示一个小圆点
        const panelOpen = ui.panel.style.display !== 'none';
        const working = state.collecting || state.busy;
        ui.launcher.style.color = panelOpen || working ? THEME.brandFg : THEME.fgHint;
        ui.badge.style.display = working ? 'block' : 'none';
        // 不在对话页面时隐藏启动按钮（有任务或面板打开时保留）
        const visible = document.getElementById(LIST_ID) || state.collecting || state.busy || panelOpen;
        ui.root.style.display = visible ? 'block' : 'none';
    }

    /** 跟随单页应用的页面切换：重建界面，消息列表元素被替换时重新监听 */
    function watch() {
        if (!document.getElementById(UI_ID)) createUI();
        const list = document.getElementById(LIST_ID);
        if (state.collecting && list && list !== state.observedList && ensureSameChat()) observeList(list);
        render();
    }

    // 按 Esc 关闭面板（预览对话框打开时除外）
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape' && ui && ui.panel.style.display !== 'none' && !document.querySelector('dialog[open]')) togglePanel(false);
    });
    setInterval(watch, WATCH_INTERVAL_MS);
    watch();
})();
