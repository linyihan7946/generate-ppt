import * as fs from 'fs';
import mammoth from 'mammoth';
import { DocumentData, SlideContent } from '../types';

interface OutlineNode {
    text: string;
    children: OutlineNode[];
    images: string[];
}

export class ParserService {
    async parseMarkdown(filePath: string): Promise<DocumentData> {
        const content = fs.readFileSync(filePath, 'utf-8');
        const lines = content.split(/\r?\n/);

        const slides: SlideContent[] = [];
        let currentSlide: SlideContent | null = null;
        let docTitle = '';

        const pushCurrentSlide = () => {
            if (!currentSlide) return;
            currentSlide.bullets = currentSlide.bullets.map((item) => item.trim()).filter(Boolean);
            slides.push(currentSlide);
            currentSlide = null;
        };

        for (const line of lines) {
            const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
            if (headingMatch) {
                const level = headingMatch[1].length;
                const headingText = headingMatch[2].trim();

                if (!docTitle && level === 1) {
                    docTitle = headingText;
                }

                pushCurrentSlide();
                currentSlide = {
                    title: headingText,
                    bullets: [],
                    images: [],
                    level,
                };
                continue;
            }

            const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.+)$/);
            if (listMatch) {
                if (!currentSlide) {
                    currentSlide = { title: docTitle || 'Markdown 内容', bullets: [], images: [], level: 1 };
                }
                const indentLevel = Math.floor(listMatch[1].length / 2);
                const bulletText = `${'  '.repeat(indentLevel)}${listMatch[3].trim()}`;
                currentSlide.bullets.push(bulletText);
                continue;
            }

            const imageRegex = /!\[[^\]]*\]\(([^)]+)\)/g;
            let imageMatch: RegExpExecArray | null = null;
            while ((imageMatch = imageRegex.exec(line)) !== null) {
                if (!currentSlide) {
                    currentSlide = { title: docTitle || 'Markdown 内容', bullets: [], images: [], level: 1 };
                }
                currentSlide.images.push(imageMatch[1].trim());
            }

            const plainLine = line.replace(/!\[[^\]]*\]\(([^)]+)\)/g, '').trim();
            if (plainLine) {
                if (!currentSlide) {
                    currentSlide = {
                        title: docTitle || plainLine,
                        bullets: [],
                        images: [],
                        level: 1,
                    };
                } else {
                    currentSlide.bullets.push(plainLine);
                }
            }
        }

        pushCurrentSlide();

        if (slides.length === 0) {
            slides.push({
                title: docTitle || 'Markdown Presentation',
                bullets: [content.trim()].filter(Boolean),
                images: [],
                level: 1,
            });
        }

        return {
            title: docTitle || slides[0].title || 'Markdown Presentation',
            slides,
        };
    }

    async parseDocx(filePath: string): Promise<DocumentData> {
        const buffer = fs.readFileSync(filePath);
        const { value: html } = await mammoth.convertToHtml(
            { buffer },
            {
                convertImage: mammoth.images.imgElement(async (image) => {
                    const base64 = await image.readAsBase64String();
                    return { src: `data:${image.contentType};base64,${base64}` };
                }),
            },
        );

        const title = this.extractFirstHeading(html) || this.extractFirstText(html) || 'Docx Presentation';
        const topLevelLists = this.extractTopLevelListBlocks(html);

        let slides: SlideContent[] = [];

        if (topLevelLists.length > 0) {
            const allNodes = topLevelLists.flatMap((block) => this.parseListBlock(block));
            slides = this.buildSlidesFromNodes(allNodes, [], 1);
        }

        if (slides.length === 0 && /<h[1-6]\b/i.test(html)) {
            slides = this.parseByHeadings(html);
        }

        if (slides.length === 0) {
            slides = this.parseByParagraphs(html, title);
        }

        if (slides.length === 0) {
            slides = [{ title, bullets: [], images: [], level: 1 }];
        }

        return { title, slides };
    }

    async parsePdf(filePath: string): Promise<DocumentData> {
        const pdfParse = this.loadPdfParse();
        const dataBuffer = fs.readFileSync(filePath);
        const data = await (pdfParse as any)(dataBuffer);

        const sections = data.text
            .split(/\n{2,}/)
            .map((section: string) => section.trim())
            .filter((section: string) => section.length > 0);

        const slides: SlideContent[] = sections.map((section: string, index: number) => {
            const lines = section
                .split('\n')
                .map((line: string) => line.trim())
                .filter((line: string) => line.length > 0);

            const title = lines[0] || `Slide ${index + 1}`;
            const bullets = lines.slice(1);

            return {
                title: title.substring(0, 80),
                bullets,
                images: [],
                level: 1,
            };
        });

        return {
            title: slides[0]?.title || 'PDF Presentation',
            slides: slides.length > 0 ? slides : [{ title: 'PDF Presentation', bullets: [], images: [], level: 1 }],
        };
    }

    private loadPdfParse(): any {
        // Lazy require avoids loading pdf-parse unless PDF parsing is actually requested.
        // This keeps markdown/docx flow compatible with older Node runtimes.
        try {
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const mod = require('pdf-parse');
            return mod.default || mod;
        } catch (error) {
            throw new Error(
                `PDF parser init failed. Please use a newer Node.js runtime (recommended >=16). Raw error: ${
                    error instanceof Error ? error.message : String(error)
                }`,
            );
        }
    }

    private parseByHeadings(html: string): SlideContent[] {
        const slides: SlideContent[] = [];
        const headingRegex = /<h([1-6])[^>]*>([\s\S]*?)<\/h\1>/gi;
        const headings: Array<{ level: number; title: string; index: number; end: number }> = [];

        let match: RegExpExecArray | null = null;
        while ((match = headingRegex.exec(html)) !== null) {
            headings.push({
                level: Number(match[1]),
                title: this.cleanText(match[2]),
                index: match.index,
                end: headingRegex.lastIndex,
            });
        }

        headings.forEach((heading, idx) => {
            const contentStart = heading.end;
            const contentEnd = idx + 1 < headings.length ? headings[idx + 1].index : html.length;
            const block = html.slice(contentStart, contentEnd);

            const bullets = this.extractBulletsFromHtml(block);
            const images = this.extractImagesFromHtml(block);

            slides.push({
                title: heading.title || `Section ${idx + 1}`,
                bullets,
                images,
                level: heading.level,
            });
        });

        return slides;
    }

    private parseByParagraphs(html: string, fallbackTitle: string): SlideContent[] {
        const paragraphBlocks = html.match(/<p[^>]*>[\s\S]*?<\/p>/gi) || [];
        const paragraphs = paragraphBlocks
            .map((block) => this.cleanText(block))
            .filter((text) => text.length > 0);

        if (paragraphs.length === 0) {
            return [];
        }

        const slides: SlideContent[] = [];
        const chunkSize = 6;

        for (let i = 0; i < paragraphs.length; i += chunkSize) {
            const chunk = paragraphs.slice(i, i + chunkSize);
            const title = i === 0 ? fallbackTitle : `${fallbackTitle}（续 ${Math.floor(i / chunkSize)}）`;
            slides.push({
                title,
                bullets: chunk,
                images: [],
                level: 1,
            });
        }

        return slides;
    }

    private buildSlidesFromNodes(nodes: OutlineNode[], path: string[], level: number): SlideContent[] {
        const slides: SlideContent[] = [];

        for (const node of nodes) {
            if (!node.text) {
                continue;
            }

            const shouldCreateSlide = node.children.length > 0 || level <= 2;
            if (shouldCreateSlide) {
                const bullets = this.buildBulletsForNode(node);
                slides.push({
                    title: node.text,
                    bullets,
                    images: node.images,
                    level,
                    breadcrumb: path.join(' / '),
                });
            }

            if (node.children.length > 0) {
                slides.push(...this.buildSlidesFromNodes(node.children, [...path, node.text], level + 1));
            }
        }

        return slides;
    }

    private buildBulletsForNode(node: OutlineNode): string[] {
        if (node.children.length === 0) {
            return [];
        }

        const allChildrenAreLeaf = node.children.every((child) => child.children.length === 0);
        if (allChildrenAreLeaf) {
            return node.children.map((child) => child.text).filter(Boolean);
        }

        return node.children.map((child) => child.text).filter(Boolean);
    }

    private parseListBlock(listHtml: string): OutlineNode[] {
        const inner = listHtml.replace(/^<(ol|ul)\b[^>]*>/i, '').replace(/<\/(ol|ul)>$/i, '');
        const itemBlocks = this.extractTopLevelLiBlocks(inner);

        return itemBlocks
            .map((item) => this.parseListItem(item))
            .filter((node) => node.text.length > 0 || node.children.length > 0);
    }

    private parseListItem(liBlock: string): OutlineNode {
        const inner = liBlock.replace(/^<li\b[^>]*>/i, '').replace(/<\/li>$/i, '');
        const nestedListBlocks = this.extractTopLevelListBlocks(inner);
        const textRegion = this.removeSegments(inner, nestedListBlocks);

        const text = this.cleanText(textRegion);
        const images = this.extractImagesFromHtml(textRegion);
        const children = nestedListBlocks.flatMap((block) => this.parseListBlock(block));

        return {
            text: text || '未命名主题',
            children,
            images,
        };
    }

    private extractTopLevelListBlocks(html: string): string[] {
        const regex = /<\/?(ol|ul)\b[^>]*>/gi;
        const stack: string[] = [];
        const blocks: string[] = [];
        let blockStart = -1;
        let match: RegExpExecArray | null = null;

        while ((match = regex.exec(html)) !== null) {
            const token = match[0];
            const isClosing = token.startsWith('</');

            if (!isClosing) {
                if (stack.length === 0) {
                    blockStart = match.index;
                }
                stack.push(match[1].toLowerCase());
            } else if (stack.length > 0) {
                stack.pop();
                if (stack.length === 0 && blockStart >= 0) {
                    blocks.push(html.slice(blockStart, regex.lastIndex));
                    blockStart = -1;
                }
            }
        }

        return blocks;
    }

    private extractTopLevelLiBlocks(html: string): string[] {
        const regex = /<\/?li\b[^>]*>/gi;
        const blocks: string[] = [];
        let depth = 0;
        let start = -1;
        let match: RegExpExecArray | null = null;

        while ((match = regex.exec(html)) !== null) {
            const token = match[0];
            const isClosing = token.startsWith('</');

            if (!isClosing) {
                if (depth === 0) {
                    start = match.index;
                }
                depth += 1;
            } else if (depth > 0) {
                depth -= 1;
                if (depth === 0 && start >= 0) {
                    blocks.push(html.slice(start, regex.lastIndex));
                    start = -1;
                }
            }
        }

        return blocks;
    }

    private removeSegments(html: string, segments: string[]): string {
        let result = html;
        for (const segment of segments) {
            result = result.replace(segment, ' ');
        }
        return result;
    }

    private extractBulletsFromHtml(html: string): string[] {
        const bullets: string[] = [];
        const listBlocks = this.extractTopLevelListBlocks(html);

        if (listBlocks.length > 0) {
            for (const block of listBlocks) {
                const nodes = this.parseListBlock(block);
                for (const node of nodes) {
                    bullets.push(node.text);
                }
            }
        }

        if (bullets.length > 0) {
            return bullets.filter(Boolean);
        }

        const paragraphBlocks = html.match(/<(p|div)[^>]*>[\s\S]*?<\/(p|div)>/gi) || [];
        for (const block of paragraphBlocks) {
            const text = this.cleanText(block);
            if (text) {
                bullets.push(text);
            }
        }

        return bullets.filter(Boolean);
    }

    private extractImagesFromHtml(html: string): string[] {
        const images: string[] = [];
        const imgRegex = /<img[^>]+src="([^"]+)"/gi;
        let match: RegExpExecArray | null = null;

        while ((match = imgRegex.exec(html)) !== null) {
            if (match[1]) {
                images.push(match[1]);
            }
        }

        return images;
    }

    private extractFirstHeading(html: string): string {
        const headingMatch = html.match(/<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/i);
        return headingMatch ? this.cleanText(headingMatch[1]) : '';
    }

    private extractFirstText(html: string): string {
        const text = this.cleanText(html);
        if (!text) return '';
        return text.length > 60 ? `${text.slice(0, 60)}...` : text;
    }

    private cleanText(html: string): string {
        if (!html) return '';
        const noTags = html.replace(/<[^>]+>/g, ' ');
        const collapsed = noTags.replace(/\s+/g, ' ').trim();
        return this.decodeHtmlEntities(collapsed);
    }

    private decodeHtmlEntities(text: string): string {
        const entities: Record<string, string> = {
            '&nbsp;': ' ',
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&apos;': "'",
        };

        let decoded = text.replace(/&(nbsp|amp|lt|gt|quot|#39|apos);/g, (entity) => entities[entity] || entity);
        decoded = decoded.replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
        decoded = decoded.replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCharCode(parseInt(code, 16)));
        return decoded;
    }
}
