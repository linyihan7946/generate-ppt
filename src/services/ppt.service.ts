import pptxgen from 'pptxgenjs';
import { DocumentData, SlideContent } from '../types';

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;

interface PptRenderConfig {
    templateStyle: boolean;
    imageOnlyMode: boolean;
    keepText: boolean;
    maxBulletsPerSlide: number;
}

export class PPTService {
    async generate(data: DocumentData, outputPath: string): Promise<string> {
        const pres = new pptxgen();
        pres.layout = 'LAYOUT_WIDE';
        pres.author = 'generate-ppt';
        pres.company = 'generate-ppt';
        pres.subject = 'Automatically generated presentation';
        pres.title = data.title;
        pres.theme = {
            headFontFace: 'Microsoft YaHei',
            bodyFontFace: 'Microsoft YaHei',
        };

        const renderConfig = this.loadRenderConfig();
        const normalizedSlides = this.paginateSlides(data.slides, renderConfig.maxBulletsPerSlide);

        this.addTitleSlide(pres, data.title, normalizedSlides, renderConfig);
        normalizedSlides.forEach((slideData, index) => {
            this.addContentSlide(pres, slideData, index + 1, normalizedSlides.length, renderConfig);
        });

        await pres.writeFile({ fileName: outputPath });
        return outputPath;
    }

    private loadRenderConfig(): PptRenderConfig {
        return {
            templateStyle: process.env.PPT_TEMPLATE_STYLE !== 'false',
            imageOnlyMode: process.env.PPT_IMAGE_ONLY_MODE === 'true',
            keepText: process.env.PPT_KEEP_TEXT !== 'false',
            maxBulletsPerSlide: Math.max(3, Number(process.env.PPT_MAX_BULLETS_PER_SLIDE || 5)),
        };
    }

    private addTitleSlide(
        pres: pptxgen,
        title: string,
        slides: SlideContent[],
        config: PptRenderConfig,
    ): void {
        const slide = pres.addSlide();
        const coverImage = slides.find((s) => s.images.length > 0)?.images[0];

        if (config.templateStyle && coverImage) {
            this.addImage(slide, coverImage, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
        } else {
            this.addFallbackBackground(slide);
        }

        slide.addShape('rect', {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: SLIDE_HEIGHT,
            line: { color: '000000', transparency: 100 },
            fill: { color: '000000', transparency: coverImage ? 48 : 25 },
        });

        slide.addText(title, {
            x: 0.8,
            y: 2.25,
            w: 11.8,
            h: 1.8,
            fontSize: 44,
            color: 'FFFFFF',
            bold: true,
            fit: 'shrink',
        });

        slide.addShape('rect', {
            x: 0.8,
            y: 5.88,
            w: 7.8,
            h: 0.08,
            line: { color: 'FFFFFF', transparency: 100 },
            fill: { color: 'FFFFFF', transparency: 0 },
        });

        slide.addText('Auto-generated presentation', {
            x: 0.82,
            y: 4.72,
            w: 5.5,
            h: 0.35,
            fontSize: 14,
            color: 'E2E8F0',
        });

        slide.addText(`Content slides: ${slides.length}`, {
            x: 0.82,
            y: 5.18,
            w: 5.5,
            h: 0.35,
            fontSize: 12,
            color: 'CBD5E1',
        });
    }

    private addContentSlide(
        pres: pptxgen,
        slideData: SlideContent,
        page: number,
        totalSlides: number,
        config: PptRenderConfig,
    ): void {
        const slide = pres.addSlide();
        const hasImage = slideData.images.length > 0;
        const preferImageOnly = config.imageOnlyMode || slideData.layout === 'image_only';
        const shouldKeepText = config.keepText && !preferImageOnly && this.canOverlayText(slideData);

        if (hasImage) {
            // Keep close to mark/ template style: full-screen visual on every content slide.
            this.addImage(slide, slideData.images[0], 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
        } else {
            this.addFallbackBackground(slide);
        }

        slide.addShape('rect', {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: shouldKeepText ? 0.5 : 0.35,
            line: { color: '000000', transparency: 100 },
            fill: { color: '000000', transparency: hasImage ? 40 : 24 },
        });

        if (shouldKeepText) {
            this.addTextOverlay(slide, slideData);
        }

        slide.addText(`${page}/${totalSlides}`, {
            x: 12.15,
            y: shouldKeepText ? 0.11 : 0.07,
            w: 1.0,
            h: 0.2,
            align: 'right',
            fontSize: shouldKeepText ? 10 : 9,
            color: 'E2E8F0',
        });
    }

    private addTextOverlay(slide: pptxgen.Slide, slideData: SlideContent): void {
        slide.addShape('roundRect', {
            x: 0.65,
            y: 0.9,
            w: 7.55,
            h: 5.95,
            line: { color: '000000', transparency: 100 },
            fill: { color: '000000', transparency: 58 },
        });

        if (slideData.breadcrumb) {
            slide.addText(slideData.breadcrumb, {
                x: 0.92,
                y: 1.14,
                w: 6.95,
                h: 0.25,
                fontSize: 10,
                color: 'CBD5E1',
                fit: 'shrink',
            });
        }

        slide.addText(slideData.title, {
            x: 0.9,
            y: 1.52,
            w: 7.0,
            h: 0.9,
            fontSize: this.titleFontSize(slideData.title),
            color: 'FFFFFF',
            bold: true,
            fit: 'shrink',
        });

        const rows = this.buildOverlayTextRows(slideData);
        slide.addText(rows, {
            x: 0.95,
            y: 2.55,
            w: 6.85,
            h: 4.0,
            valign: 'top',
            fit: 'shrink',
        });
    }

    private buildOverlayTextRows(slideData: SlideContent): Array<{ text: string; options: Record<string, unknown> }> {
        const rows: Array<{ text: string; options: Record<string, unknown> }> = [];
        const dedupedBullets = this.deduplicateBullets(slideData.bullets);
        const summary = this.cleanInlineText(slideData.summary || '');
        const shouldShowSummary = summary && !this.isSummaryRedundant(summary, dedupedBullets, slideData.title);

        if (shouldShowSummary) {
            rows.push({
                text: summary,
                options: {
                    breakLine: true,
                    color: 'BFDBFE',
                    fontSize: 14,
                    bold: true,
                },
            });
        }

        if (dedupedBullets.length === 0) {
            rows.push({
                text: 'No sub-items in source content for this node.',
                options: {
                    color: 'CBD5E1',
                    fontSize: 15,
                },
            });
            return rows;
        }

        dedupedBullets.forEach((raw, index) => {
            const normalized = this.normalizeBullet(raw);
            rows.push({
                text: `${'  '.repeat(normalized.level)}• ${normalized.text}`,
                options: {
                    breakLine: index < dedupedBullets.length - 1,
                    color: 'F8FAFC',
                    fontSize: Math.max(14, 19 - normalized.level),
                },
            });
        });

        return rows;
    }

    private deduplicateBullets(bullets: string[]): string[] {
        const unique = new Set<string>();
        const deduped: string[] = [];
        for (const raw of bullets) {
            const normalized = this.normalizeBullet(raw);
            if (!normalized.text) continue;
            const key = this.normalizeForCompare(normalized.text);
            if (!key || unique.has(key)) continue;
            unique.add(key);
            deduped.push(raw);
        }
        return deduped;
    }

    private isSummaryRedundant(summary: string, bullets: string[], title: string): boolean {
        const summaryNorm = this.normalizeForCompare(summary);
        if (!summaryNorm) return true;

        const summaryWithoutTitleNorm = this.normalizeForCompare(
            summary.replace(new RegExp(`^\\s*${this.escapeRegExp(title)}\\s*[:：,，。\\-]*\\s*`, 'i'), ''),
        );

        const candidates = [summaryNorm, summaryWithoutTitleNorm].filter(Boolean);
        for (const bullet of bullets) {
            const bulletNorm = this.normalizeForCompare(this.normalizeBullet(bullet).text);
            if (!bulletNorm) continue;
            for (const candidate of candidates) {
                if (!candidate) continue;
                if (candidate === bulletNorm) {
                    return true;
                }
                if (candidate.length >= 8 && bulletNorm.length >= 8) {
                    if (candidate.includes(bulletNorm) || bulletNorm.includes(candidate)) {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private cleanInlineText(text: string): string {
        return text.replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
    }

    private normalizeForCompare(text: string): string {
        return this.cleanInlineText(text).toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '');
    }

    private escapeRegExp(input: string): string {
        return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    private normalizeBullet(raw: string): { text: string; level: number } {
        const expanded = raw.replace(/\t/g, '  ');
        const leadingSpacesMatch = expanded.match(/^(\s*)/);
        const leadingSpaces = leadingSpacesMatch ? leadingSpacesMatch[1].length : 0;
        const level = Math.min(3, Math.floor(leadingSpaces / 2));
        const text = expanded.trim().replace(/^[-*]\s+/, '');
        return { text, level };
    }

    private canOverlayText(slideData: SlideContent): boolean {
        const titleTooLong = slideData.title.length > 56;
        const bulletCountTooHigh = slideData.bullets.length > 6;
        const bulletTextLength = slideData.bullets.reduce((sum, b) => sum + b.length, 0);
        const bulletTextTooLong = bulletTextLength > 260;
        return !titleTooLong && !bulletCountTooHigh && !bulletTextTooLong;
    }

    private paginateSlides(slides: SlideContent[], maxBulletsPerSlide: number): SlideContent[] {
        const paginated: SlideContent[] = [];

        for (const slide of slides) {
            if (slide.bullets.length <= maxBulletsPerSlide) {
                paginated.push(slide);
                continue;
            }

            let remaining = [...slide.bullets];
            let chunkIndex = 0;

            while (remaining.length > 0) {
                const chunk = remaining.slice(0, maxBulletsPerSlide);
                remaining = remaining.slice(maxBulletsPerSlide);

                paginated.push({
                    ...slide,
                    title: chunkIndex === 0 ? slide.title : `${slide.title} (Part ${chunkIndex + 1})`,
                    bullets: chunk,
                    images: slide.images,
                });

                chunkIndex += 1;
            }
        }

        return paginated;
    }

    private addFallbackBackground(slide: pptxgen.Slide): void {
        slide.addShape('rect', {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: SLIDE_HEIGHT,
            line: { color: '0F172A' },
            fill: { color: '0F172A' },
        });

        slide.addShape('roundRect', {
            x: 7.9,
            y: -0.8,
            w: 6.4,
            h: 4.8,
            line: { color: '38BDF8', transparency: 100 },
            fill: { color: '38BDF8', transparency: 82 },
        });

        slide.addShape('roundRect', {
            x: -0.8,
            y: 4.65,
            w: 6.0,
            h: 3.6,
            line: { color: '1D4ED8', transparency: 100 },
            fill: { color: '1D4ED8', transparency: 86 },
        });
    }

    private addImage(slide: pptxgen.Slide, image: string, x: number, y: number, w: number, h: number): void {
        if (!image) return;

        if (image.startsWith('data:image')) {
            slide.addImage({ data: image, x, y, w, h });
            return;
        }

        slide.addImage({ path: image, x, y, w, h });
    }

    private titleFontSize(title: string): number {
        if (title.length <= 14) return 34;
        if (title.length <= 24) return 30;
        if (title.length <= 34) return 27;
        return 23;
    }
}
