const cheerio = require('cheerio');
const fetch = require('node-fetch');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        Table, TableRow, TableCell, WidthType, BorderStyle, ExternalHyperlink,
        ImageRun, PageBreak } = require('docx');

const HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
};

const MAX_PAGES = 30;       // Max pages to crawl
const FETCH_TIMEOUT = 15000; // 15s per page
const MAX_TOTAL_TIME = 50000; // 50s total (Vercel limit is 60s)

// ── Fetch page ──
async function fetchPage(url) {
    if (!url.startsWith('http://') && !url.startsWith('https://')) url = 'https://' + url;
    const resp = await fetch(url, { headers: HEADERS, timeout: FETCH_TIMEOUT, redirect: 'follow' });
    if (!resp.ok) throw new Error(`Website returned status ${resp.status}`);
    const html = await resp.text();
    return { html, finalUrl: resp.url || url };
}

// ── Get base domain for same-site check ──
function getBaseDomain(url) {
    try {
        const u = new URL(url);
        return u.hostname;
    } catch { return ''; }
}

// ── Discover all internal links from a page ──
function discoverInternalLinks(html, baseUrl) {
    const $ = cheerio.load(html);
    const baseDomain = getBaseDomain(baseUrl);
    const links = new Set();

    $('a[href]').each((_, a) => {
        let href = $(a).attr('href') || '';
        // Skip anchors, javascript, mailto, tel
        if (href.startsWith('#') || href.startsWith('javascript:') || href.startsWith('mailto:') || href.startsWith('tel:')) return;
        try {
            const fullUrl = new URL(href, baseUrl).href;
            const urlDomain = getBaseDomain(fullUrl);
            // Only follow same-domain links
            if (urlDomain === baseDomain) {
                // Clean URL - remove hash fragments
                const cleanUrl = fullUrl.split('#')[0];
                // Skip file downloads, images, etc.
                if (/\.(pdf|zip|png|jpg|jpeg|gif|svg|mp4|mp3|doc|docx|xls|xlsx|ppt|css|js|json|xml|ico|woff|ttf|eot)$/i.test(cleanUrl)) return;
                links.add(cleanUrl);
            }
        } catch {}
    });

    return [...links];
}

// ── Extract content from a single page ──
function extractPageContent(html, pageUrl) {
    const $ = cheerio.load(html);

    // Remove unwanted tags
    $('script, style, nav, footer, noscript, iframe, svg, header').remove();

    const result = {
        title: $('title').text().trim() || '',
        metaDescription: $('meta[name="description"]').attr('content') || '',
        sections: [],
        images: [],
        url: pageUrl,
    };

    // Find main content area
    const main = $('main').length ? $('main') :
                 $('article').length ? $('article') :
                 $('[role="main"]').length ? $('[role="main"]') :
                 $('body');

    let currentSection = { heading: '', level: 2, paragraphs: [] };

    main.find('h1, h2, h3, h4, h5, h6, p, ul, ol, blockquote, pre, table').each((_, el) => {
        const tag = $(el).prop('tagName').toLowerCase();

        if (/^h[1-6]$/.test(tag)) {
            if (currentSection.paragraphs.length) result.sections.push(currentSection);
            const level = parseInt(tag[1]);
            currentSection = { heading: $(el).text().trim(), level: Math.max(level, 2), paragraphs: [] };
        } else if (tag === 'p') {
            const text = $(el).text().trim();
            if (text.length > 5) currentSection.paragraphs.push({ type: 'text', content: text });
        } else if (tag === 'ul' || tag === 'ol') {
            const items = [];
            $(el).children('li').each((_, li) => {
                const t = $(li).text().trim();
                if (t) items.push(t);
            });
            if (items.length) currentSection.paragraphs.push({ type: 'list', listType: tag, items });
        } else if (tag === 'blockquote') {
            const text = $(el).text().trim();
            if (text) currentSection.paragraphs.push({ type: 'quote', content: text });
        } else if (tag === 'pre') {
            const text = $(el).text().trim();
            if (text) currentSection.paragraphs.push({ type: 'code', content: text });
        } else if (tag === 'table') {
            const rows = [];
            $(el).find('tr').each((_, tr) => {
                const cells = [];
                $(tr).find('td, th').each((_, td) => cells.push($(td).text().trim()));
                if (cells.length) rows.push(cells);
            });
            if (rows.length) currentSection.paragraphs.push({ type: 'table', rows });
        }
    });

    if (currentSection.paragraphs.length) result.sections.push(currentSection);

    // Fallback: grab all text if no sections
    if (!result.sections.length) {
        const allText = main.text().replace(/\s+/g, ' ').trim();
        if (allText.length > 30) {
            result.sections.push({
                heading: result.title || 'Content',
                level: 2,
                paragraphs: [{ type: 'text', content: allText }]
            });
        }
    }

    // Images
    const seenImgs = new Set();
    main.find('img[src]').each((_, img) => {
        let src = $(img).attr('src') || '';
        if (src.startsWith('data:') || seenImgs.has(src)) return;
        try { src = new URL(src, pageUrl).href; } catch { return; }
        seenImgs.add(src);
        result.images.push({ url: src, alt: ($(img).attr('alt') || '').trim() });
    });

    return result;
}

// ── Crawl entire website ──
async function crawlWebsite(startUrl, maxPages) {
    const startTime = Date.now();
    const visited = new Set();
    const queue = [startUrl];
    const pages = [];
    const allImages = [];
    const allLinks = new Set();
    let siteTitle = '';

    while (queue.length > 0 && visited.size < maxPages) {
        // Check time limit
        if (Date.now() - startTime > MAX_TOTAL_TIME) break;

        const url = queue.shift();
        const cleanUrl = url.split('#')[0].replace(/\/$/, ''); // Normalize

        if (visited.has(cleanUrl)) continue;
        visited.add(cleanUrl);

        try {
            const { html, finalUrl } = await fetchPage(url);

            // Extract content from this page
            const pageContent = extractPageContent(html, finalUrl);

            // Use first page title as site title
            if (!siteTitle && pageContent.title) siteTitle = pageContent.title;

            // Only add page if it has meaningful content
            if (pageContent.sections.length > 0) {
                pages.push(pageContent);
            }

            // Collect images
            allImages.push(...pageContent.images);

            // Discover new internal links
            const newLinks = discoverInternalLinks(html, finalUrl);
            for (const link of newLinks) {
                const cleanLink = link.split('#')[0].replace(/\/$/, '');
                allLinks.add(link);
                if (!visited.has(cleanLink)) {
                    queue.push(link);
                }
            }
        } catch (err) {
            // Skip failed pages, continue crawling
            continue;
        }
    }

    return {
        siteTitle,
        pages,
        totalImages: allImages.length,
        totalLinks: allLinks.size,
        pagesCrawled: visited.size,
        images: allImages,
        allLinks: [...allLinks],
    };
}

// ── Download image ──
async function downloadImage(url) {
    try {
        const resp = await fetch(url, { headers: HEADERS, timeout: 8000 });
        if (!resp.ok) return null;
        const ct = resp.headers.get('content-type') || '';
        if (!ct.includes('image') && !url.match(/\.(png|jpg|jpeg|gif|webp|bmp)$/i)) return null;
        const buffer = await resp.buffer();
        if (buffer.length < 500) return null;
        return buffer;
    } catch { return null; }
}

// ── Build DOCX from crawled site ──
async function buildDocx(crawlResult, includeImages, includeLinks) {
    const children = [];

    // ── Cover page ──
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 600 },
        children: [new TextRun({
            text: crawlResult.siteTitle || 'Website Content Export',
            bold: true, size: 56, color: '1a1a2e', font: 'Calibri',
        })],
    }));

    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [
            new TextRun({ text: `Full Website Extraction`, size: 24, color: '4444aa', font: 'Calibri', bold: true }),
        ],
    }));

    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: `Pages crawled: ${crawlResult.pagesCrawled} | `, size: 20, color: '666666', font: 'Calibri' }),
            new TextRun({ text: `Images found: ${crawlResult.totalImages} | `, size: 20, color: '666666', font: 'Calibri' }),
            new TextRun({ text: `Links found: ${crawlResult.totalLinks}`, size: 20, color: '666666', font: 'Calibri' }),
        ],
    }));

    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: `Extracted: ${new Date().toLocaleString()}`, size: 20, color: '888888', font: 'Calibri' }),
        ],
    }));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── Table of Contents (by page) ──
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: 'Table of Contents', font: 'Calibri' })],
    }));

    for (let i = 0; i < crawlResult.pages.length; i++) {
        const page = crawlResult.pages[i];
        const pageTitle = page.title || page.url;
        children.push(new Paragraph({
            spacing: { after: 60 },
            children: [
                new TextRun({ text: `${i + 1}. `, bold: true, size: 20, color: '333333', font: 'Calibri' }),
                new TextRun({ text: pageTitle.substring(0, 80), size: 20, color: '4444aa', font: 'Calibri' }),
            ],
        }));
        // Show sub-sections
        for (const sec of page.sections.slice(0, 5)) {
            if (sec.heading) {
                children.push(new Paragraph({
                    spacing: { after: 30 },
                    children: [new TextRun({ text: `      • ${sec.heading}`, size: 18, color: '777777', font: 'Calibri' })],
                }));
            }
        }
    }

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── Content from each page ──
    for (let i = 0; i < crawlResult.pages.length; i++) {
        const page = crawlResult.pages[i];

        // Page title as H1
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 300, after: 60 },
            children: [new TextRun({ text: page.title || `Page ${i + 1}`, font: 'Calibri' })],
        }));

        // Page URL
        children.push(new Paragraph({
            spacing: { after: 160 },
            children: [new TextRun({ text: `Source: ${page.url}`, size: 18, color: '4444aa', font: 'Calibri', italics: true })],
        }));

        // Meta description
        if (page.metaDescription) {
            children.push(new Paragraph({
                spacing: { after: 120 },
                children: [new TextRun({ text: page.metaDescription, size: 20, color: '555555', font: 'Calibri', italics: true })],
            }));
        }

        // Sections
        for (const section of page.sections) {
            if (section.heading) {
                const headingLevel = [HeadingLevel.HEADING_2, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4][Math.min(section.level - 1, 4)];
                children.push(new Paragraph({
                    heading: headingLevel,
                    spacing: { before: 200, after: 100 },
                    children: [new TextRun({ text: section.heading, font: 'Calibri' })],
                }));
            }

            for (const para of section.paragraphs) {
                if (para.type === 'text') {
                    children.push(new Paragraph({
                        spacing: { after: 100 },
                        children: [new TextRun({ text: para.content, size: 22, font: 'Calibri', color: '333333' })],
                    }));
                } else if (para.type === 'list') {
                    for (const item of para.items) {
                        children.push(new Paragraph({
                            bullet: { level: 0 },
                            spacing: { after: 50 },
                            children: [new TextRun({ text: item, size: 22, font: 'Calibri', color: '333333' })],
                        }));
                    }
                } else if (para.type === 'quote') {
                    children.push(new Paragraph({
                        indent: { left: 720 },
                        spacing: { after: 100 },
                        children: [new TextRun({ text: `"${para.content}"`, italics: true, size: 22, color: '555555', font: 'Calibri' })],
                    }));
                } else if (para.type === 'code') {
                    children.push(new Paragraph({
                        indent: { left: 360 },
                        spacing: { after: 100 },
                        children: [new TextRun({ text: para.content, size: 18, font: 'Consolas', color: '2d2d2d' })],
                    }));
                } else if (para.type === 'table' && para.rows.length) {
                    const maxCols = Math.max(...para.rows.map(r => r.length));
                    const tableRows = para.rows.map((row, ri) =>
                        new TableRow({
                            children: Array.from({ length: maxCols }, (_, ci) =>
                                new TableCell({
                                    width: { size: Math.floor(9000 / maxCols), type: WidthType.DXA },
                                    children: [new Paragraph({
                                        children: [new TextRun({ text: row[ci] || '', bold: ri === 0, size: 20, font: 'Calibri' })],
                                    })],
                                })
                            ),
                        })
                    );
                    children.push(new Table({ rows: tableRows }));
                    children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
                }
            }
        }

        // Page separator (except last page)
        if (i < crawlResult.pages.length - 1) {
            children.push(new Paragraph({ children: [new PageBreak()] }));
        }
    }

    // ── Images section ──
    if (includeImages && crawlResult.images.length) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Images Found Across Website', font: 'Calibri' })],
        }));

        // Deduplicate images
        const seenImgUrls = new Set();
        const uniqueImages = crawlResult.images.filter(img => {
            if (seenImgUrls.has(img.url)) return false;
            seenImgUrls.add(img.url);
            return true;
        });

        let imgCount = 0;
        for (const img of uniqueImages.slice(0, 15)) {
            const buf = await downloadImage(img.url);
            if (buf) {
                try {
                    children.push(new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 160, after: 60 },
                        children: [new ImageRun({ data: buf, transformation: { width: 450, height: 300 }, type: 'jpg' })],
                    }));
                    if (img.alt) {
                        children.push(new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: img.alt, italics: true, size: 18, color: '777777', font: 'Calibri' })],
                        }));
                    }
                    imgCount++;
                } catch {}
            }
            if (imgCount >= 10) break;
        }

        if (uniqueImages.length > 15) {
            children.push(new Paragraph({
                children: [new TextRun({ text: `... and ${uniqueImages.length - 15} more images across the site`, size: 20, color: '888888', font: 'Calibri' })],
            }));
        }
    }

    // ── All Links section ──
    if (includeLinks && crawlResult.allLinks.length) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'All Links Found on Website', font: 'Calibri' })],
        }));

        const linkRows = [
            new TableRow({
                children: [
                    new TableCell({ width: { size: 800, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: '#', bold: true, size: 20, font: 'Calibri' })] })] }),
                    new TableCell({ width: { size: 8200, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: 'URL', bold: true, size: 20, font: 'Calibri' })] })] }),
                ],
            }),
            ...crawlResult.allLinks.slice(0, 100).map((link, idx) =>
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${idx + 1}`, size: 18, font: 'Calibri' })] })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: link.substring(0, 120), size: 16, font: 'Calibri', color: '4444aa' })] })] }),
                    ],
                })
            ),
        ];
        children.push(new Table({ rows: linkRows }));
    }

    // ── Footer ──
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400 },
        children: [new TextRun({
            text: `Generated by Web2Doc | ${crawlResult.pagesCrawled} pages crawled | ${new Date().toLocaleString()}`,
            size: 16, color: '999999', font: 'Calibri',
        })],
    }));

    const doc = new Document({ sections: [{ children }] });
    return await Packer.toBuffer(doc);
}

// ── API Handler ──
module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { url, includeImages = true, includeLinks = true, maxPages = MAX_PAGES } = req.body || {};

    if (!url) {
        return res.status(400).json({ error: 'Please provide a URL' });
    }

    try {
        // Limit maxPages to prevent abuse
        const pageLimit = Math.min(Math.max(1, maxPages), MAX_PAGES);

        // Crawl the entire website
        const crawlResult = await crawlWebsite(url, pageLimit);

        if (!crawlResult.pages.length) {
            return res.status(400).json({
                error: 'Could not extract meaningful content. The site may use JavaScript rendering or block scraping.'
            });
        }

        // Build docx
        const docBuffer = await buildDocx(crawlResult, includeImages, includeLinks);

        // Filename
        let hostname = 'website';
        try { hostname = new URL(crawlResult.pages[0].url).hostname.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30); } catch {}
        const filename = `${hostname}_full_${Date.now().toString(36)}.docx`;

        const docBase64 = docBuffer.toString('base64');

        return res.status(200).json({
            title: crawlResult.siteTitle,
            sections: crawlResult.pages.reduce((sum, p) => sum + p.sections.length, 0),
            images: crawlResult.totalImages,
            links: crawlResult.totalLinks,
            pagesCrawled: crawlResult.pagesCrawled,
            totalPages: crawlResult.pages.length,
            filename,
            docBase64,
        });

    } catch (err) {
        const msg = err.message || 'Something went wrong';
        if (msg.includes('status')) return res.status(400).json({ error: msg });
        if (msg.includes('ENOTFOUND') || msg.includes('getaddrinfo')) return res.status(400).json({ error: 'Could not connect to the website. Please check the URL.' });
        if (msg.includes('timeout')) return res.status(400).json({ error: 'The website took too long to respond.' });
        return res.status(500).json({ error: msg });
    }
};
