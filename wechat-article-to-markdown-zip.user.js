// ==UserScript==
// @name         微信公众号文章转Markdown+图片ZIP
// @namespace    http://tampermonkey.net/
// @version      2.3.0
// @description  将微信公众号文章转换为Markdown，并连同图片一起打包下载为ZIP
// @author       ChatGPT
// @match        https://mp.weixin.qq.com/s/*
// @match        https://mp.weixin.qq.com/s?*
// @match        https://mp.weixin.qq.com/cgi-bin/readtemplate*
// @grant        GM_xmlhttpRequest
// @grant        GM_download
// @connect      mp.weixin.qq.com
// @connect      qpic.cn
// @connect      qlogo.cn
// @connect      mmbiz.qpic.cn
// @connect      mmbiz.qlogo.cn
// @connect      res.wx.qq.com
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    function log(...args) {
        console.log('[微信转MD+ZIP]', ...args);
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // ========== 轻量 ZIP 打包（STORE 模式，零依赖） ==========
    const _crc32Table = (() => {
        const table = new Uint32Array(256);
        for (let i = 0; i < 256; i++) {
            let c = i;
            for (let j = 0; j < 8; j++) {
                c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            }
            table[i] = c;
        }
        return table;
    })();

    function crc32(data) {
        let crc = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) {
            crc = _crc32Table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    function buildZipBlob(files) {
        // files: Array<{ name: string, data: Uint8Array }>
        const encoder = new TextEncoder();
        const localHeaders = [];
        const centralHeaders = [];
        let offset = 0;

        for (const file of files) {
            const nameBytes = encoder.encode(file.name);
            const data = file.data;
            const crc = crc32(data);
            const size = data.length;

            // Local file header (30 + nameLen + data)
            const local = new Uint8Array(30 + nameBytes.length + size);
            const lv = new DataView(local.buffer);
            lv.setUint32(0, 0x04034b50, true);   // signature
            lv.setUint16(4, 20, true);            // version needed
            lv.setUint16(6, 0x0800, true);        // flags: UTF-8
            lv.setUint16(8, 0, true);             // compression: STORE
            lv.setUint16(10, 0, true);            // mod time
            lv.setUint16(12, 0x0021, true);       // mod date (2000-01-01)
            lv.setUint32(14, crc, true);          // crc-32
            lv.setUint32(18, size, true);         // compressed size
            lv.setUint32(22, size, true);         // uncompressed size
            lv.setUint16(26, nameBytes.length, true);
            lv.setUint16(28, 0, true);            // extra field length
            local.set(nameBytes, 30);
            local.set(data, 30 + nameBytes.length);
            localHeaders.push(local);

            // Central directory header (46 + nameLen)
            const central = new Uint8Array(46 + nameBytes.length);
            const cv = new DataView(central.buffer);
            cv.setUint32(0, 0x02014b50, true);   // signature
            cv.setUint16(4, 20, true);            // version made by
            cv.setUint16(6, 20, true);            // version needed
            cv.setUint16(8, 0x0800, true);        // flags: UTF-8
            cv.setUint16(10, 0, true);            // compression: STORE
            cv.setUint16(12, 0, true);            // mod time
            cv.setUint16(14, 0x0021, true);       // mod date
            cv.setUint32(16, crc, true);
            cv.setUint32(20, size, true);
            cv.setUint32(24, size, true);
            cv.setUint16(28, nameBytes.length, true);
            cv.setUint16(30, 0, true);            // extra field length
            cv.setUint16(32, 0, true);            // file comment length
            cv.setUint16(34, 0, true);            // disk number
            cv.setUint16(36, 0, true);            // internal attrs
            cv.setUint32(38, 0, true);            // external attrs
            cv.setUint32(42, offset, true);       // local header offset
            central.set(nameBytes, 46);
            centralHeaders.push(central);

            offset += local.length;
        }

        const centralDirOffset = offset;
        let centralDirSize = 0;
        for (const c of centralHeaders) centralDirSize += c.length;

        // End of central directory (22 bytes)
        const eocd = new Uint8Array(22);
        const ev = new DataView(eocd.buffer);
        ev.setUint32(0, 0x06054b50, true);
        ev.setUint16(4, 0, true);
        ev.setUint16(6, 0, true);
        ev.setUint16(8, files.length, true);
        ev.setUint16(10, files.length, true);
        ev.setUint32(12, centralDirSize, true);
        ev.setUint32(16, centralDirOffset, true);
        ev.setUint16(20, 0, true);

        return new Blob([...localHeaders, ...centralHeaders, eocd], { type: 'application/zip' });
    }
    // ========== ZIP 打包结束 ==========

    function sanitizeFileName(name) {
        return (name || '微信文章')
            .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function createDownloadButton() {
        if (document.getElementById('markdown-zip-download-btn')) return;

        const button = document.createElement('button');
        button.id = 'markdown-zip-download-btn';
        button.innerHTML = '📦 下载 Markdown+图片';
        button.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 999999;
            background: #07c160;
            color: white;
            border: none;
            padding: 12px 18px;
            border-radius: 8px;
            font-size: 14px;
            cursor: pointer;
            box-shadow: 0 2px 12px rgba(0,0,0,0.25);
            transition: all 0.2s ease;
        `;

        button.addEventListener('mouseenter', () => {
            button.style.background = '#06ad56';
            button.style.transform = 'translateY(-1px)';
        });

        button.addEventListener('mouseleave', () => {
            button.style.background = '#07c160';
            button.style.transform = 'translateY(0)';
        });

        button.addEventListener('click', async () => {
            button.disabled = true;
            const oldText = button.innerHTML;
            button.innerHTML = '⏳ 正在处理...';
            try {
                await convertToMarkdownZip((msg) => {
                    button.innerHTML = msg;
                });
            } catch (err) {
                console.error(err);
                alert('处理失败：' + (err?.message || err));
            } finally {
                button.disabled = false;
                button.innerHTML = oldText;
            }
        });

        document.body.appendChild(button);
        log('按钮创建成功');
    }

    function showMessage(message, isError = false) {
        const div = document.createElement('div');
        div.textContent = message;
        div.style.cssText = `
            position: fixed;
            top: 80px;
            right: 20px;
            z-index: 1000000;
            background: ${isError ? '#e53935' : '#4caf50'};
            color: white;
            padding: 12px 16px;
            border-radius: 6px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            font-size: 14px;
            max-width: 360px;
            word-break: break-word;
        `;
        document.body.appendChild(div);
        setTimeout(() => div.remove(), 4000);
    }

    function getArticleTitle() {
        const el = document.querySelector('#activity-name, .rich_media_title, h1');
        return el ? el.textContent.trim() : '微信文章';
    }

    function getArticleAuthor() {
        const el = document.querySelector('#js_name, .profile_nickname');
        return el ? el.textContent.trim() : '';
    }

    function getPublishTime() {
        const el = document.querySelector('#publish_time');
        if (!el) return '';
        return el.textContent.trim();
    }

    function getContentElement() {
        const selectors = [
            '#js_content',
            '.rich_media_content',
            '.article-content',
            '[data-role="content"]',
            '.rich_media_area_primary',
            '.rich_media_area_primary_inner',
            '.rich_media_wrp',
            '#page-content',
            '.page-content',
            '.content',
            '.post-content',
            '.article-body',
            '.entry-content',
            'main',
            '[role="main"]',
        ];

        for (const selector of selectors) {
            const el = document.querySelector(selector);
            if (el && el.textContent && el.textContent.trim().length > 100) {
                return el;
            }
        }

        return null;
    }

    async function expandHiddenContent(container) {
        const expandButtons = container.querySelectorAll(
            '[data-action="expand"], .expand-btn, .show-more, .unfold, .js_unfold, .js_show_more, .js_expand'
        );

        expandButtons.forEach(btn => {
            try { btn.click(); } catch (_) {}
        });

        const hiddenElements = container.querySelectorAll(
            '[style*="display: none"], [style*="display:none"], .hidden, .collapse, .fold, .js_fold'
        );

        hiddenElements.forEach(el => {
            try {
                el.style.display = 'block';
                el.style.visibility = 'visible';
                el.style.opacity = '1';
                el.style.height = 'auto';
                el.style.maxHeight = 'none';
                el.classList.remove('hidden', 'collapse', 'fold', 'js_fold');
            } catch (_) {}
        });

        await sleep(800);
    }

    function absoluteUrl(url) {
        if (!url) return '';
        let u = url.trim();
        if (u.startsWith('//')) u = location.protocol + u;
        // 微信图片偶见 http://，在 https 页面属混合内容会被拦截，统一升级
        if (u.startsWith('http://')) u = 'https://' + u.slice(7);
        return u;
    }

    function getImageUrl(img) {
        const attrs = [
            'data-src',
            'data-original',
            'data-lazy-src',
            'data-url',
            'src',
            'data-original-src',
            'data-actualsrc'
        ];

        for (const attr of attrs) {
            const val = img.getAttribute(attr);
            if (val && val.trim() && !val.startsWith('data:image/svg') && !val.includes('loading')) {
                return absoluteUrl(val.trim());
            }
        }

        const parent = img.parentElement;
        if (parent) {
            for (const attr of attrs) {
                const val = parent.getAttribute(attr);
                if (val && val.trim() && !val.startsWith('data:image/svg')) {
                    return absoluteUrl(val.trim());
                }
            }
        }

        return '';
    }

    function guessExtFromUrl(url) {
        try {
            // 微信图片 URL 多为 .../0?wx_fmt=png 形式，扩展名藏在 wx_fmt 参数里
            const fmtMatch = url.match(/[?&]wx_fmt=([a-z0-9]+)/i);
            if (fmtMatch) {
                const fmt = fmtMatch[1].toLowerCase();
                if (fmt === 'jpeg') return 'jpg';
                if (['png', 'gif', 'webp', 'jpg', 'bmp'].includes(fmt)) return fmt;
            }
            const clean = url.split('?')[0].toLowerCase();
            if (clean.endsWith('.png')) return 'png';
            if (clean.endsWith('.gif')) return 'gif';
            if (clean.endsWith('.webp')) return 'webp';
            if (clean.endsWith('.jpeg')) return 'jpg';
            if (clean.endsWith('.jpg')) return 'jpg';
        } catch (_) {}
        return 'jpg';
    }

    function guessExtFromContentType(contentType, fallback = 'jpg') {
        if (!contentType) return fallback;
        const ct = contentType.toLowerCase();
        if (ct.includes('png')) return 'png';
        if (ct.includes('gif')) return 'gif';
        if (ct.includes('webp')) return 'webp';
        if (ct.includes('jpeg') || ct.includes('jpg')) return 'jpg';
        return fallback;
    }

    function processTable(table, imageCollector) {
        // 使用 :scope 避免选中嵌套表格的行和单元格
        const rows = Array.from(table.querySelectorAll(':scope > tr, :scope > thead > tr, :scope > tbody > tr'));
        if (!rows.length) return '';

        // 单行单列表格：微信常用作"文本框"样式，转为引用块
        if (rows.length === 1) {
            const directCells = rows[0].querySelectorAll(':scope > th, :scope > td');
            if (directCells.length === 1) {
                const cellContent = htmlToMarkdown(directCells[0], imageCollector).trim();
                if (cellContent) {
                    return '\n' + cellContent.split('\n')
                        .filter(l => l.trim())
                        .map(line => `> ${line}`).join('\n') + '\n\n';
                }
                return '';
            }
        }

        let md = '\n';
        let firstRow = true;

        rows.forEach(row => {
            const cells = row.querySelectorAll(':scope > th, :scope > td');
            if (!cells.length) return;
            const values = Array.from(cells).map(cell => {
                // 保留单元格内的行内格式（加粗、链接等）
                let cellMd = htmlToMarkdown(cell, imageCollector).trim();
                // 表格单元格必须单行，压平换行
                cellMd = cellMd.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
                cellMd = cellMd.replace(/\|/g, '\\|');
                return cellMd;
            });
            md += '| ' + values.join(' | ') + ' |\n';
            if (firstRow) {
                md += '| ' + values.map(() => '---').join(' | ') + ' |\n';
                firstRow = false;
            }
        });

        return md + '\n';
    }

    function extractCodeText(element) {
        // 微信代码片段：<pre> 中每行一个 <code>，需逐个提取并用换行拼接
        const codeChildren = element.querySelectorAll(':scope > code');
        if (codeChildren.length > 1) {
            return Array.from(codeChildren)
                .map(c => c.textContent || '')
                .join('\n')
                .replace(/\r\n/g, '\n')
                .replace(/\r/g, '\n')
                .replace(/\t/g, '    ');
        }
        return (element.innerText || element.textContent || '')
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .replace(/\t/g, '    ');
    }

    function codeFence(codeText, lang) {
        // 如果代码内容含 ```，用更长的围栏避免冲突
        let fence = '```';
        const backtickRuns = codeText.match(/`{3,}/g) || [];
        for (const run of backtickRuns) {
            if (run.length >= fence.length) {
                fence = '`'.repeat(run.length + 1);
            }
        }
        return `\n${fence}${lang}\n${codeText}\n${fence}\n\n`;
    }

    // svg/path 等 SVG 元素的 className 是 SVGAnimatedString 对象而非字符串，
    // 直接调用字符串方法会抛 TypeError（曾导致含视频文章整篇转换失败）
    function getElementClass(el) {
        const cls = el.className;
        if (typeof cls === 'string') return cls;
        return (typeof el.getAttribute === 'function' && el.getAttribute('class')) || '';
    }

    // 装饰性/交互性元素，不属于正文内容
    const SKIP_TAGS = new Set([
        'svg', 'path', 'style', 'script', 'noscript',
        'button', 'input', 'textarea', 'select', 'canvas',
        'mp-style-type',
    ]);

    // 微信视频容器：播放器 UI 含大量控件文字（倍速/全屏/进度条等），需整体替换为占位符
    const VIDEO_CONTAINER_RE = /(?:^|\s)(?:js_page_video|page_video|video_iframe|js_video_channel_card|channel_video_card)/;

    function htmlToMarkdown(node, imageCollector) {
        let markdown = '';

        for (const child of node.childNodes) {
            if (child.nodeType === Node.TEXT_NODE) {
                markdown += child.textContent.replace(/\s+/g, ' ');
                continue;
            }

            if (child.nodeType !== Node.ELEMENT_NODE) continue;

            const tag = child.tagName.toLowerCase();

            if (SKIP_TAGS.has(tag)) continue;

            try {
                markdown += elementToMarkdown(child, tag, imageCollector);
            } catch (err) {
                // 单个节点失败降级为纯文本，不中断整篇转换
                log(`节点转换失败 <${tag}>，已降级为纯文本:`, err?.message || err);
                markdown += (child.textContent || '').replace(/\s+/g, ' ');
            }
        }

        return markdown;
    }

    function elementToMarkdown(child, tag, imageCollector) {
        let markdown = '';

        switch (tag) {
            case 'h1':
                markdown += `\n# ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'h2':
                markdown += `\n## ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'h3':
                markdown += `\n### ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'h4':
                markdown += `\n#### ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'h5':
                markdown += `\n##### ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'h6':
                markdown += `\n###### ${htmlToMarkdown(child, imageCollector).trim()}\n\n`;
                break;
            case 'p': {
                const text = htmlToMarkdown(child, imageCollector).trim();
                if (text) markdown += `\n${text}\n\n`;
                break;
            }
            case 'br':
                markdown += '  \n';
                break;
            case 'strong':
            case 'b': {
                const inner = htmlToMarkdown(child, imageCollector).trim();
                if (!inner) break;
                // 避免嵌套加粗产生 ****
                if (inner.startsWith('**') && inner.endsWith('**')) {
                    markdown += inner;
                } else {
                    markdown += `**${inner}**`;
                }
                break;
            }
            case 'em':
            case 'i': {
                const inner = htmlToMarkdown(child, imageCollector).trim();
                if (!inner) break;
                if (inner.startsWith('*') && inner.endsWith('*') && !inner.startsWith('**')) {
                    markdown += inner;
                } else {
                    markdown += `*${inner}*`;
                }
                break;
            }
            case 'del':
            case 's': {
                const inner = htmlToMarkdown(child, imageCollector).trim();
                if (!inner) break;
                if (inner.startsWith('~~') && inner.endsWith('~~')) {
                    markdown += inner;
                } else {
                    markdown += `~~${inner}~~`;
                }
                break;
            }
            case 'code':
                if (child.parentElement && child.parentElement.tagName.toLowerCase() === 'pre') {
                    break;
                }
                markdown += `\`${(child.textContent || '').trim()}\``;
                break;
            case 'pre': {
                const codeEl = child.querySelector('code');
                const className = (codeEl?.className || child.className || '');
                const langMatch = className.match(/language-(\w+)/);
                const dataLang = child.getAttribute('data-lang') || '';
                const lang = dataLang || (langMatch ? langMatch[1] : '');
                // 用 pre 元素提取文本（微信每行一个 <code>，只取第一个会丢内容）
                const codeText = extractCodeText(child);
                markdown += codeFence(codeText, lang);
                break;
            }
            case 'a': {
                const href = child.href || '';
                const text = htmlToMarkdown(child, imageCollector).trim() || href;
                markdown += `[${text}](${href})`;
                break;
            }
            case 'img': {
                const url = getImageUrl(child);
                if (url) {
                    const alt = child.getAttribute('alt') || child.getAttribute('title') || '图片';
                    const index = imageCollector.length + 1;
                    const ext = guessExtFromUrl(url);
                    const localName = `images/${String(index).padStart(3, '0')}.${ext}`;
                    imageCollector.push({
                        originalUrl: url,
                        alt,
                        localName,
                        ext,
                    });
                    markdown += `![${alt}](${localName})\n\n`;
                }
                break;
            }
            case 'video': {
                const src = absoluteUrl(child.getAttribute('src') || child.getAttribute('data-src') || '');
                markdown += src
                    ? `\n> 🎬 视频：[点击播放](${src})（链接可能有时效，失效请看原文）\n\n`
                    : '\n> 🎬 此处有视频，请查看原文\n\n';
                break;
            }
            case 'iframe': {
                const src = absoluteUrl(child.getAttribute('data-src') || child.getAttribute('src') || '');
                if (src) markdown += `\n> 🎬 嵌入内容：[点击查看](${src})\n\n`;
                break;
            }
            case 'mp-common-profile': {
                // 公众号名片：textContent 为空，信息在 data-* 属性里
                const nickname = child.getAttribute('data-nickname') || '';
                const alias = child.getAttribute('data-alias') || '';
                if (nickname) {
                    markdown += `\n> 📇 公众号名片：**${nickname}**${alias ? `（${alias}）` : ''}\n\n`;
                }
                break;
            }
            case 'table':
                markdown += processTable(child, imageCollector);
                break;
            case 'blockquote': {
                const content = htmlToMarkdown(child, imageCollector).trim();
                if (content) {
                    markdown += '\n' + content.split('\n').map(line => `> ${line}`).join('\n') + '\n\n';
                }
                break;
            }
            case 'ul': {
                // 代码块行号列表不属于正文
                if (getElementClass(child).includes('line-index')) break;
                const ulItems = child.querySelectorAll(':scope > li');
                markdown += '\n';
                for (const li of ulItems) {
                    markdown += `- ${htmlToMarkdown(li, imageCollector).trim()}\n`;
                }
                markdown += '\n';
                break;
            }
            case 'ol': {
                const olItems = child.querySelectorAll(':scope > li');
                markdown += '\n';
                for (let idx = 0; idx < olItems.length; idx++) {
                    markdown += `${idx + 1}. ${htmlToMarkdown(olItems[idx], imageCollector).trim()}\n`;
                }
                markdown += '\n';
                break;
            }
            case 'li':
                markdown += htmlToMarkdown(child, imageCollector);
                break;
            case 'hr':
                markdown += '\n---\n\n';
                break;
            case 'div':
            case 'span':
            case 'section':
            default: {
                const cls = getElementClass(child);
                // 视频容器整体替换为占位符，避免播放器控件文字混入正文
                if (VIDEO_CONTAINER_RE.test(cls)) {
                    const videoEl = child.querySelector('video');
                    const src = videoEl
                        ? absoluteUrl(videoEl.getAttribute('src') || videoEl.getAttribute('data-src') || '')
                        : '';
                    markdown += src
                        ? `\n> 🎬 视频：[点击播放](${src})（链接可能有时效，失效请看原文）\n\n`
                        : '\n> 🎬 此处有视频，请查看原文\n\n';
                    break;
                }
                // 微信代码片段特殊处理：跳过行号 <ul>，直接提取 <pre><code>
                if (cls.includes('code-snippet') || child.querySelector?.('pre.code-snippet__js, pre[data-lang], ul.code-snippet__line-index')) {
                    const pre = child.querySelector('pre');
                    if (pre) {
                        const langAttr = pre.getAttribute('data-lang') || '';
                        const codeEl = pre.querySelector('code');
                        const classLang = (codeEl?.className || pre.className || '').match(/language-(\w+)/);
                        const lang = langAttr || (classLang ? classLang[1] : '');
                        // 用 pre 元素提取全部文本
                        const codeText = extractCodeText(pre);
                        if (codeText.trim()) {
                            markdown += codeFence(codeText, lang);
                            break;
                        }
                    }
                }
                // section 段落间距处理：叶子 section（不含块级子元素）当作 <p> 处理
                if (tag === 'section') {
                    const hasBlockChild = child.querySelector(
                        'section, p, div, table, pre, ul, ol, blockquote, h1, h2, h3, h4, h5, h6, hr'
                    );
                    if (!hasBlockChild) {
                        const text = htmlToMarkdown(child, imageCollector).trim();
                        if (text) markdown += `\n${text}\n\n`;
                        break;
                    }
                }
                markdown += htmlToMarkdown(child, imageCollector);
                break;
            }
        }

        return markdown;
    }

    function cleanMarkdown(markdown) {
        const codeBlocks = [];
        let idx = 0;

        // 匹配 ``` 及更长围栏（如 ````、````` 等）
        let text = markdown.replace(/(`{3,})[\s\S]*?\1/g, (m) => {
            const key = `__CODE_BLOCK_${idx}__`;
            codeBlocks.push(m);
            idx++;
            return key;
        });

        text = text
            .replace(/\n{3,}/g, '\n\n')
            .replace(/[ \t]+$/gm, '')
            .replace(/^\n+/g, '')
            .replace(/\n+$/g, '');

        codeBlocks.forEach((block, i) => {
            text = text.replace(`__CODE_BLOCK_${i}__`, block);
        });

        return text + '\n';
    }

    function gmRequestBlob(url, timeoutMs = 30000) {
        return new Promise((resolve, reject) => {
            let settled = false;
            const timer = setTimeout(() => {
                if (!settled) {
                    settled = true;
                    reject(new Error(`下载超时(${timeoutMs / 1000}s): ${url}`));
                }
            }, timeoutMs);

            GM_xmlhttpRequest({
                method: 'GET',
                url,
                responseType: 'blob',
                timeout: timeoutMs,
                headers: {
                    Referer: location.href,
                },
                onload: function (res) {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    if (res.status >= 200 && res.status < 300) {
                        resolve({
                            blob: res.response,
                            contentType: res.responseHeaders || '',
                        });
                    } else {
                        reject(new Error(`下载失败: ${res.status} ${url}`));
                    }
                },
                onerror: function (err) {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    reject(new Error(`网络错误: ${url}`));
                },
                ontimeout: function () {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    reject(new Error(`GM超时: ${url}`));
                }
            });
        });
    }

    function parseContentType(headersText) {
        if (!headersText) return '';
        const lines = headersText.split('\n');
        for (const line of lines) {
            const parts = line.split(':');
            if (parts.length >= 2 && parts[0].trim().toLowerCase() === 'content-type') {
                return parts.slice(1).join(':').trim();
            }
        }
        return '';
    }

    const IMAGE_CONCURRENCY = 4;

    async function downloadImages(imageList, onProgress) {
        const uniqueMap = new Map();
        const finalList = [];

        for (const item of imageList) {
            if (!uniqueMap.has(item.originalUrl)) {
                uniqueMap.set(item.originalUrl, item);
                finalList.push(item);
            }
        }

        const results = new Array(finalList.length).fill(null);
        let nextIndex = 0;
        let doneCount = 0;

        async function worker() {
            while (nextIndex < finalList.length) {
                const i = nextIndex++;
                const item = finalList[i];
                log(`下载图片 ${i + 1}/${finalList.length}`, item.originalUrl.substring(0, 80));
                try {
                    const res = await gmRequestBlob(item.originalUrl);
                    const contentType = parseContentType(res.contentType);
                    const realExt = guessExtFromContentType(contentType, item.ext);
                    const localName = `images/${String(i + 1).padStart(3, '0')}.${realExt}`;

                    item.finalLocalName = localName;

                    const arrayBuf = await res.blob.arrayBuffer();
                    results[i] = { name: localName, data: new Uint8Array(arrayBuf) };
                } catch (err) {
                    // 下载失败回退为原始 URL，Markdown 中仍可在线查看
                    log(`图片 ${i + 1} 下载失败，回退为原始链接:`, err.message);
                    item.finalLocalName = item.originalUrl;
                }
                doneCount++;
                onProgress?.(`🖼️ 下载图片 ${doneCount}/${finalList.length}`);
            }
        }

        const workerCount = Math.min(IMAGE_CONCURRENCY, finalList.length);
        await Promise.all(Array.from({ length: workerCount }, () => worker()));

        // 正文中重复出现的同一张图，共享首次出现的下载结果
        for (const item of imageList) {
            const canonical = uniqueMap.get(item.originalUrl);
            if (canonical && canonical !== item) {
                item.finalLocalName = canonical.finalLocalName || canonical.localName;
            }
        }

        const files = results.filter(Boolean);
        log(`图片下载完成，成功 ${files.length}/${finalList.length}`);

        return { imageList, files };
    }

    function replaceMarkdownImagePaths(markdown, imageList) {
        let output = markdown;
        for (const item of imageList) {
            const safeOld = item.localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const target = item.finalLocalName || item.localName;
            // 函数式替换，避免 URL 中的 $ 被 replace 当作特殊替换符
            output = output.replace(new RegExp(`\\]\\(${safeOld}\\)`, 'g'), () => `](${target})`);
        }
        return output;
    }

    async function downloadZip(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }

    async function convertToMarkdownZip(onProgress) {
        const title = sanitizeFileName(getArticleTitle());
        const author = getArticleAuthor();
        const publishTime = getPublishTime();

        const contentElement = getContentElement();
        if (!contentElement) {
            throw new Error('未找到文章内容，请确认页面已完整加载。');
        }

        // const bodyText = document.body.innerText || '';
        // if (bodyText.includes('环境异常') || bodyText.includes('安全验证') || bodyText.includes('验证')) {
        //     throw new Error('检测到微信安全验证页面，请先手动通过验证。');
        // }

        onProgress?.('🔍 展开内容...');
        await expandHiddenContent(contentElement);

        const imageCollector = [];

        onProgress?.('📝 转换 Markdown...');
        let markdown = `# ${title}\n\n`;

        if (author) {
            markdown += `**作者：** ${author}\n\n`;
        }
        if (publishTime) {
            markdown += `**发布时间：** ${publishTime}\n\n`;
        }

        markdown += `**原文链接：** ${location.href}\n\n---\n\n`;
        markdown += htmlToMarkdown(contentElement, imageCollector);
        markdown = cleanMarkdown(markdown);

        log('提取图片数量:', imageCollector.length);

        onProgress?.('📥 下载图片...');
        const { imageList: finalImages, files: imageFiles } = await downloadImages(imageCollector, onProgress);

        markdown = replaceMarkdownImagePaths(markdown, finalImages);

        const encoder = new TextEncoder();
        const mdFile = { name: `${title}.md`, data: encoder.encode(markdown) };

        onProgress?.('📦 生成 ZIP...');
        log('开始生成 ZIP...');
        const zipBlob = buildZipBlob([mdFile, ...imageFiles]);
        log(`ZIP 生成完成, 大小: ${(zipBlob.size / 1024).toFixed(1)} KB`);

        onProgress?.('⬇️ 下载 ZIP...');
        log('触发下载...');
        await downloadZip(zipBlob, `${title}.zip`);

        showMessage(`已下载：${title}.zip`);
    }

    function init() {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', createDownloadButton);
        } else {
            createDownloadButton();
        }
    }

    init();
})();