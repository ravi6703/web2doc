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

// ── Fetch page ──
async function fetchPage(url) {
    if (!url.startsWith('http://') && !url.startsWith('https://')) url = 'https://' + url;
    const resp = await fetch(url, { headers: HEADERS, timeout: 25000, redirect: 'follow' });
    if (!resp.ok) throw new Error(`Website returned status ${resp.status}`);
    const html = await resp.text();
    return { html, finalUrl: resp.url || url };
}

// ── Extract content ──
function extractContent(html, baseUrl) {
    const $ = cheerio.load(html);

    // Remove unwanted tags
    $('script, style, nav, footer, noscript, iframe, svg').remove();

    const result = {
        title: $('title').text().trim() || '',
        metaDescription: $('meta[name="description"]').attr('content') || '',
        sections: [],
        images: [],
        links: [],
        url: baseUrl,
    };

    // Find main content area
    const main = $('main').length ? $('main') :
                 $('article').length ? $('article') :
                 $('[role="main"]').length ? $('[role="main"]') :
                 $('body');

    let currentSection = { heading: 'Introduction', level: 1, paragraphs: [] };

    main.find('h1, h2, h3, h4, h5, h6, p, ul, ol, blockquote, pre, table').each((_, el) => {
        const tag = $(el).prop('tagName').toLowerCase();

        if (/^h[1-6]$/.test(tag)) {
            if (currentSection.paragraphs.length) result.sections.push(currentSection);
            const level = parseInt(tag[1]);
            currentSection = { heading: $(el).text().trim(), level, paragraphs: [] };
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

    // Fallback: if no sections, grab all text
    if (!result.sections.length) {
        const allText = main.text().replace(/\s+/g, ' ').trim();
        if (allText.length > 20) {
            result.sections.push({
                heading: result.title || 'Page Content',
                level: 1,
                paragraphs: [{ type: 'text', content: allText }]
            });
        }
    }

    // Images
    const seenImgs = new Set();
    main.find('img[src]').each((_, img) => {
        let src = $(img).attr('src') || '';
        if (src.startsWith('data:') || seenImgs.has(src)) return;
        try {
            src = new URL(src, baseUrl).href;
        } catch { return; }
        seenImgs.add(src);
        result.images.push({ url: src, alt: ($(img).attr('alt') || '').trim() });
    });

    // Links
    const seenLinks = new Set();
    main.find('a[href]').each((_, a) => {
        let href = $(a).attr('href') || '';
        if (href.startsWith('#') || href.startsWith('javascript:') || href.startsWith('mailto:')) return;
        try {
            href = new URL(href, baseUrl).href;
        } catch { return; }
        if (seenLinks.has(href)) return;
        seenLinks.add(href);
        const text = $(a).text().trim() || href;
        result.links.push({ url: href, text: text.substring(0, 120) });
    });

    return result;
}

// ── Download image ──
async function downloadImage(url) {
    try {
        const resp = await fetch(url, { headers: HEADERS, timeout: 10000 });
        if (!resp.ok) return null;
        const ct = resp.headers.get('content-type') || '';
        if (!ct.includes('image') && !url.match(/\.(png|jpg|jpeg|gif|webp|bmp)$/i)) return null;
        const buffer = await resp.buffer();
        if (buffer.length < 500) return null; // skip tiny images
        return buffer;
    } catch {
        return null;
    }
}

// ── Build DOCX ──
async function buildDocx(content, includeImages, includeLinks) {
    const children = [];

    // Title
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 600 },
        children: [new TextRun({
            text: content.title || 'Extracted Website Content',
            bold: true,
            size: 56,
            color: '1a1a2e',
            font: 'Calibri',
        })],
    }));

    // Source URL & date
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: `Source: ${content.url}`, size: 20, color: '666666', font: 'Calibri' }),
            new TextRun({ text: `\nExtracted: ${new Date().toLocaleString()}`, size: 20, color: '666666', font: 'Calibri', break: 1 }),
        ],
    }));

    // Meta description
    if (content.metaDescription) {
        children.push(new Paragraph({
            spacing: { before: 200, after: 300 },
            children: [new TextRun({
                text: content.metaDescription,
                italics: true,
                size: 22,
                color: '555555',
                font: 'Calibri',
            })],
        }));
    }

    // Page break
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // Table of Contents (text-based)
    if (content.sections.length > 2) {
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Table of Contents', font: 'Calibri' })],
        }));
        for (const sec of content.sections) {
            const indent = '    '.repeat(Math.max(0, sec.level - 1));
            children.push(new Paragraph({
                spacing: { after: 40 },
                children: [new TextRun({
                    text: `${indent}• ${sec.heading}`,
                    size: 20,
                    color: '4444aa',
                    font: 'Calibri',
                })],
            }));
        }
        children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    // Sections
    for (const section of content.sections) {
        const headingLevel = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4][Math.min(section.level - 1, 3)];

        children.push(new Paragraph({
            heading: headingLevel,
            spacing: { before: 240, after: 120 },
            children: [new TextRun({ text: section.heading, font: 'Calibri' })],
        }));

        for (const para of section.paragraphs) {
            if (para.type === 'text') {
                children.push(new Paragraph({
                    spacing: { after: 120 },
                    children: [new TextRun({ text: para.content, size: 22, font: 'Calibri', color: '333333' })],
                }));
            } else if (para.type === 'list') {
                for (const item of para.items) {
                    children.push(new Paragraph({
                        bullet: { level: 0 },
                        spacing: { after: 60 },
                        children: [new TextRun({ text: item, size: 22, font: 'Calibri', color: '333333' })],
                    }));
                }
            } else if (para.type === 'quote') {
                children.push(new Paragraph({
                    indent: { left: 720 },
                    spacing: { after: 120 },
                    children: [new TextRun({
                        text: `"${para.content}"`,
                        italics: true,
                        size: 22,
                        color: '555555',
                        font: 'Calibri',
                    })],
                }));
            } else if (para.type === 'code') {
                children.push(new Paragraph({
                    indent: { left: 360 },
                    spacing: { after: 120 },
                    children: [new TextRun({
                        text: para.content,
                        size: 18,
                        font: 'Consolas',
                        color: '2d2d2d',
                    })],
                }));
            } else if (para.type === 'table' && para.rows.length) {
                const maxCols = Math.max(...para.rows.map(r => r.length));
                const tableRows = para.rows.map((row, ri) =>
                    new TableRow({
                        children: Array.from({ length: maxCols }, (_, ci) =>
                            new TableCell({
                                width: { size: Math.floor(9000 / maxCols), type: WidthType.DXA },
                                children: [new Paragraph({
                                    children: [new TextRun({
                                        text: row[ci] || '',
                                        bold: ri === 0,
                                        size: 20,
                                        font: 'Calibri',
                                    })],
                                })],
                            })
                        ),
                    })
                );
                children.push(new Table({ rows: tableRows }));
                children.push(new Paragraph({ spacing: { after: 200 }, children: [] }));
            }
        }
    }

    // Images section
    if (includeImages && content.images.length) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Images', font: 'Calibri' })],
        }));

        let imgCount = 0;
        for (const img of content.images.slice(0, 10)) {
            const buf = await downloadImage(img.url);
            if (buf) {
                try {
                    children.push(new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 80 },
                        children: [new ImageRun({
                            data: buf,
                            transformation: { width: 450, height: 300 },
                            type: 'jpg',
                        })],
                    }));
                    if (img.alt) {
                        children.push(new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({
                                text: img.alt,
                                italics: true,
                                size: 18,
                                color: '777777',
                                font: 'Calibri',
                            })],
                        }));
                    }
                    imgCount++;
                } catch {}
            }
            if (imgCount >= 8) break;
        }

        if (content.images.length > 10) {
            children.push(new Paragraph({
                children: [new TextRun({
                    text: `... and ${content.images.length - 10} more images on the page`,
                    size: 20, color: '888888', font: 'Calibri',
                })],
            }));
        }
    }

    // Links section
    if (includeLinks && content.links.length) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Links Found on Page', font: 'Calibri' })],
        }));

        const linkRows = [
            new TableRow({
                children: [
                    new TableCell({ width: { size: 4000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: 'Link Text', bold: true, size: 20, font: 'Calibri' })] })] }),
                    new TableCell({ width: { size: 5000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: 'URL', bold: true, size: 20, font: 'Calibri' })] })] }),
                ],
            }),
            ...content.links.slice(0, 80).map(link =>
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: link.text.substring(0, 80), size: 18, font: 'Calibri' })] })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: link.url.substring(0, 100), size: 16, font: 'Calibri', color: '4444aa' })] })] }),
                    ],
                })
            ),
        ];

        children.push(new Table({ rows: linkRows }));
    }

    // Footer
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400 },
        children: [new TextRun({
            text: `Generated by Web2Doc from ${content.url}`,
            size: 16, color: '999999', font: 'Calibri',
        })],
    }));

    const doc = new Document({
        sections: [{ children }],
    });

    return await Packer.toBuffer(doc);
}

// ── API Handler ──
module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { url, includeImages = true, includeLinks = true } = req.body || {};

    if (!url) {
        return res.status(400).json({ error: 'Please provide a URL' });
    }

    try {
        // 1. Fetch
        const { html, finalUrl } = await fetchPage(url);

        // 2. Extract
        const content = extractContent(html, finalUrl);

        if (!content.sections.length && !content.title) {
            return res.status(400).json({
                error: 'Could not extract meaningful content. The site may use JavaScript rendering or block scraping.'
            });
        }

        // 3. Build docx
        const docBuffer = await buildDocx(content, includeImages, includeLinks);

        // 4. Create safe filename
        let hostname = 'website';
        try { hostname = new URL(finalUrl).hostname.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30); } catch {}
        const filename = `${hostname}_${Date.now().toString(36)}.docx`;

        // 5. Return as base64 (for client-side download)
        const docBase64 = docBuffer.toString('base64');

        return res.status(200).json({
            title: content.title,
            sections: content.sections.length,
            images: content.images.length,
            links: content.links.length,
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
