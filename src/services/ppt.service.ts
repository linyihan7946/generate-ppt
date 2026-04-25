import pptxgen from 'pptxgenjs';
import { DeckBrief, DocumentData, SlideContent, SlideRole } from '../types';

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;
const COLORS = {
    ink: '0F172A',
    darkBg: '0A0F1E',
    slate: '334155',
    muted: '64748B',
    line: 'CBD5E1',
    paper: 'F8FAFC',
    panel: 'FFFFFF',
    panelSoft: 'E2E8F0',
    accent: '0EA5E9',
    accentDark: '0369A1',
    accentSoft: 'BAE6FD',
    accentLight: 'E0F2FE',
    success: '16A34A',
    successSoft: 'DCFCE7',
    amber: 'D97706',
    amberSoft: 'FEF3C7',
    violet: '8B5CF6',
    violetSoft: 'EDE9FE',
    violetBg: '7C3AED',
    white: 'FFFFFF',
    whiteAlpha06: 'F0F0F0',
    whiteAlpha10: 'E6E6E6',
    blueFade: '93C5FD',
};

interface PptRenderConfig {
    templateStyle: boolean;
    imageOnlyMode: boolean;
    keepText: boolean;
    maxBulletsPerSlide: number;
    showSourceRefs: boolean;
}

interface ComparisonColumns {
    leftTitle: string;
    rightTitle: string;
    leftItems: string[];
    rightItems: string[];
}

interface TimelineEvent {
    label: string;
    detail: string;
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

        const config = this.loadRenderConfig();
        const slides = this.paginateSlides(data.slides, config.maxBulletsPerSlide);

        this.addTitleSlide(pres, data, slides, config);
        slides.forEach((slideData, index) => {
            this.addRoleAwareSlide(pres, slideData, index + 1, slides.length, data.brief, config);
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
            showSourceRefs: process.env.PPT_SHOW_SOURCE_REFS !== 'false',
        };
    }

    private addTitleSlide(
        pres: pptxgen,
        data: DocumentData,
        slides: SlideContent[],
        config: PptRenderConfig,
    ): void {
        const slide = pres.addSlide();
        const coverImage = slides.find((item) => item.images.length > 0)?.images[0];
        const brief = data.brief;
        const isCjk = this.isMostlyCjk([data.title, brief?.deckGoal || '', slides.map((item) => item.title).join(' ')].join(' '));

        // Dark background base
        slide.background = { color: COLORS.darkBg };

        // Background pattern: radial glow accent (left)
        slide.addShape('ellipse', {
            x: -1.0, y: 0.5, w: 7.0, h: 6.0,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 85 },
        });
        // Background pattern: radial glow violet (right)
        slide.addShape('ellipse', {
            x: 7.5, y: -1.0, w: 5.5, h: 4.5,
            line: { color: COLORS.violet, transparency: 100 },
            fill: { color: COLORS.violet, transparency: 88 },
        });

        // Cover image on right 55% with gradient mask simulation
        if (coverImage && config.templateStyle) {
            this.addImage(slide, coverImage, 6.0, 0, 7.333, SLIDE_HEIGHT);
            // Gradient overlay: left part dark
            slide.addShape('rect', {
                x: 6.0, y: 0, w: 3.5, h: SLIDE_HEIGHT,
                line: { color: COLORS.darkBg, transparency: 100 },
                fill: { color: COLORS.darkBg, transparency: 20 },
            });
            // Semi-transparent overlay on image
            slide.addShape('rect', {
                x: 6.0, y: 0, w: 7.333, h: SLIDE_HEIGHT,
                line: { color: COLORS.darkBg, transparency: 100 },
                fill: { color: COLORS.darkBg, transparency: 55 },
            });
        }

        // Bottom gradient overlay
        slide.addShape('rect', {
            x: 0, y: 5.0, w: SLIDE_WIDTH, h: 2.5,
            line: { color: COLORS.darkBg, transparency: 100 },
            fill: { color: COLORS.darkBg, transparency: 10 },
        });

        // Decorative circles (border only)
        slide.addShape('ellipse', {
            x: 10.0, y: -1.0, w: 5.0, h: 5.0,
            line: { color: COLORS.accent, width: 0.5, transparency: 85 },
            fill: { type: 'none' } as any,
        });
        slide.addShape('ellipse', {
            x: 10.8, y: 0.3, w: 3.0, h: 3.0,
            line: { color: COLORS.violet, width: 0.5, transparency: 88 },
            fill: { type: 'none' } as any,
        });

        // PRESENTATION badge
        slide.addShape('roundRect', {
            x: 0.75, y: 1.0, w: 2.6, h: 0.42,
            rectRadius: 0.21,
            line: { color: COLORS.accent, width: 1.0, transparency: 60 },
            fill: { type: 'none' } as any,
        });
        slide.addText('PRESENTATION', {
            x: 0.75, y: 1.0, w: 2.6, h: 0.42,
            fontSize: 11, fontFace: 'Microsoft YaHei',
            color: COLORS.accent,
            bold: true,
            charSpacing: 3,
            align: 'center',
            valign: 'middle',
        });

        // Main title (large)
        slide.addText(data.title, {
            x: 0.75, y: 1.7, w: 8.5, h: 2.0,
            fontSize: 38,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
            valign: 'middle',
        });

        // Goal / subtitle
        const goal = this.cleanInlineText(brief?.deckGoal || '');
        if (goal) {
            slide.addText(goal, {
                x: 0.78, y: 3.85, w: 7.2, h: 0.65,
                fontSize: 16,
                color: 'B4C6DB',
                fit: 'shrink',
                valign: 'top',
            });
        }

        // Meta tags row (audience + style capsules)
        const audience = brief?.audience || '';
        const style = brief?.style || '';
        let metaX = 0.78;
        if (audience) {
            const tagW = Math.max(1.5, audience.length * 0.16 + 0.8);
            slide.addShape('roundRect', {
                x: metaX, y: 4.7, w: tagW, h: 0.38,
                rectRadius: 0.08,
                line: { color: COLORS.white, width: 0.5, transparency: 90 },
                fill: { color: COLORS.white, transparency: 94 },
            });
            slide.addText(`\u{1F465}  ${String(audience)}`, {
                x: metaX, y: 4.7, w: tagW, h: 0.38,
                fontSize: 11, color: 'B0BEC5',
                align: 'center', valign: 'middle',
            });
            metaX += tagW + 0.2;
        }
        if (style) {
            const tagW = Math.max(1.5, style.length * 0.16 + 0.8);
            slide.addShape('roundRect', {
                x: metaX, y: 4.7, w: tagW, h: 0.38,
                rectRadius: 0.08,
                line: { color: COLORS.white, width: 0.5, transparency: 90 },
                fill: { color: COLORS.white, transparency: 94 },
            });
            slide.addText(`\u{1F3A8}  ${String(style)}`, {
                x: metaX, y: 4.7, w: tagW, h: 0.38,
                fontSize: 11, color: 'B0BEC5',
                align: 'center', valign: 'middle',
            });
        }

        // Accent line (bottom-left)
        slide.addShape('rect', {
            x: 0.75, y: 6.2, w: 2.0, h: 0.04,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 0 },
        });
    }

    private addRoleAwareSlide(
        pres: pptxgen,
        slideData: SlideContent,
        page: number,
        totalSlides: number,
        brief: DeckBrief | undefined,
        config: PptRenderConfig,
    ): void {
        const role = slideData.slideRole || 'content';
        const slide = pres.addSlide();

        switch (role) {
            case 'agenda':
                this.addAgendaSlide(slide, slideData, brief);
                break;
            case 'section_divider':
                this.addSectionDividerSlide(slide, slideData, config);
                break;
            case 'timeline':
                this.addTimelineSlide(slide, slideData);
                break;
            case 'comparison':
                this.addComparisonSlide(slide, slideData);
                break;
            case 'process':
                this.addProcessSlide(slide, slideData);
                break;
            case 'data_highlight':
                this.addDataHighlightSlide(slide, slideData, config);
                break;
            case 'summary':
                this.addSummarySlide(slide, slideData, brief);
                break;
            case 'next_step':
                this.addNextStepSlide(slide, slideData);
                break;
            case 'key_insight':
                this.addKeyInsightSlide(slide, slideData, config);
                break;
            case 'content':
            default:
                this.addContentSlide(slide, slideData, config);
                break;
        }

        this.addFooter(slide, slideData, page, totalSlides, config, role);
    }

    private addAgendaSlide(slide: pptxgen.Slide, slideData: SlideContent, brief?: DeckBrief): void {
        // Dark background matching HTML agenda-slide
        slide.background = { color: COLORS.ink };

        // Subtle radial glows
        slide.addShape('ellipse', {
            x: -2.0, y: 5.0, w: 8.0, h: 5.0,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 92 },
        });
        slide.addShape('ellipse', {
            x: 9.0, y: -2.0, w: 6.0, h: 5.0,
            line: { color: COLORS.violet, transparency: 100 },
            fill: { color: COLORS.violet, transparency: 94 },
        });

        const isCjk = this.isMostlyCjk([slideData.title, slideData.summary || '', ...(brief?.coreTakeaways || [])].join(' '));

        // CONTENTS badge
        slide.addText(isCjk ? '\u76EE\u5F55' : 'CONTENTS', {
            x: 0.75, y: 0.65, w: 2.5, h: 0.3,
            fontSize: 11, color: COLORS.accent,
            bold: true, charSpacing: 3,
        });

        // Title
        slide.addText(slideData.title, {
            x: 0.75, y: 1.1, w: 11.5, h: 0.85,
            fontSize: 30, color: COLORS.white,
            bold: true, fit: 'shrink',
        });

        // Agenda items in 2-column grid
        const agendaItems = this.deduplicateBullets(
            (brief?.chapterTitles || []).length > 0 ? brief!.chapterTitles : slideData.bullets,
        ).slice(0, 8);

        const cols = 2;
        const cardW = 5.5;
        const cardH = 0.72;
        const gapX = 0.5;
        const gapY = 0.18;
        const startX = 0.75;
        const startY = 2.3;

        agendaItems.forEach((item, index) => {
            const col = index % cols;
            const row = Math.floor(index / cols);
            const x = startX + col * (cardW + gapX);
            const y = startY + row * (cardH + gapY);

            // Card background (semi-transparent)
            slide.addShape('roundRect', {
                x, y, w: cardW, h: cardH,
                rectRadius: 0.1,
                line: { color: COLORS.white, width: 0.5, transparency: 92 },
                fill: { color: COLORS.white, transparency: 96 },
            });

            // Large number (accent color)
            slide.addText(String(index + 1).padStart(2, '0'), {
                x: x + 0.2, y: y + 0.1, w: 0.65, h: 0.52,
                fontSize: 22, color: COLORS.accent,
                bold: true, valign: 'middle',
            });

            // Item label
            slide.addText(item, {
                x: x + 0.9, y: y + 0.1, w: cardW - 1.2, h: 0.52,
                fontSize: 16, color: 'D9E2EC',
                fit: 'shrink', valign: 'middle',
            });
        });
    }

    private addSectionDividerSlide(slide: pptxgen.Slide, slideData: SlideContent, config: PptRenderConfig): void {
        const heroImage = slideData.images[0];
        if (heroImage && config.templateStyle) {
            this.addImage(slide, heroImage, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
            slide.addShape('rect', {
                x: 0,
                y: 0,
                w: SLIDE_WIDTH,
                h: SLIDE_HEIGHT,
                line: { color: COLORS.ink, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 34 },
            });
        } else {
            this.addDarkBackground(slide);
        }

        if (slideData.breadcrumb) {
            slide.addText(slideData.breadcrumb, {
                x: 0.92,
                y: 0.95,
                w: 4.9,
                h: 0.25,
                fontSize: 11,
                color: COLORS.accentSoft,
                fit: 'shrink',
            });
        }

        slide.addText(slideData.title, {
            x: 0.9,
            y: 2.08,
            w: 8.5,
            h: 1.35,
            fontSize: 30,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
            valign: 'middle',
        });

        const summary = this.cleanInlineText(slideData.keyMessage || slideData.summary || '');
        if (summary) {
            slide.addText(summary, {
                x: 0.95,
                y: 3.72,
                w: 6.75,
                h: 0.7,
                fontSize: 16,
                color: 'E2E8F0',
                fit: 'shrink',
            });
        }

        slide.addShape('rect', {
            x: 0.92,
            y: 5.68,
            w: 7.0,
            h: 0.06,
            line: { color: COLORS.white, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 0 },
        });
    }

    private addTimelineSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        // Dark background matching HTML timeline-slide
        this.addDarkBackground(slide);

        // Title
        slide.addText(slideData.title, {
            x: 0.6, y: 0.5, w: 12.0, h: 0.85,
            fontSize: 28, color: COLORS.white,
            bold: true, fit: 'shrink',
        });

        const events = this.buildTimelineEvents(slideData).slice(0, 5);
        const count = events.length;
        const trackY = 3.1;
        const totalW = 11.5;
        const startX = 0.9;
        const itemW = totalW / count;

        // Horizontal track line
        slide.addShape('rect', {
            x: startX, y: trackY, w: totalW, h: 0.04,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 70 },
        });

        events.forEach((event, index) => {
            const cx = startX + index * itemW + itemW / 2;

            // Glow circle (larger, semi-transparent)
            slide.addShape('ellipse', {
                x: cx - 0.18, y: trackY - 0.16, w: 0.36, h: 0.36,
                line: { color: COLORS.accent, transparency: 100 },
                fill: { color: COLORS.accent, transparency: 60 },
            });
            // Dot center
            slide.addShape('ellipse', {
                x: cx - 0.1, y: trackY - 0.08, w: 0.2, h: 0.2,
                line: { color: COLORS.accent, transparency: 100 },
                fill: { color: COLORS.accent, transparency: 0 },
            });

            // Connector line to next (skip last)
            if (index < count - 1) {
                const nextCx = startX + (index + 1) * itemW + itemW / 2;
                slide.addShape('rect', {
                    x: cx + 0.1, y: trackY - 0.01, w: nextCx - cx - 0.2, h: 0.06,
                    line: { color: COLORS.accent, transparency: 100 },
                    fill: { color: COLORS.accent, transparency: 70 },
                });
            }

            // Card below track
            const cardY = trackY + 0.6;
            const cardW = itemW - 0.3;
            const cardX = cx - cardW / 2;

            slide.addShape('roundRect', {
                x: cardX, y: cardY, w: cardW, h: 2.2,
                rectRadius: 0.1,
                line: { color: COLORS.white, width: 0.5, transparency: 92 },
                fill: { color: COLORS.white, transparency: 96 },
            });

            // STEP label
            const isCjk = this.isMostlyCjk(event.detail);
            slide.addText(isCjk ? `\u6B65\u9AA4 ${index + 1}` : `STEP ${index + 1}`, {
                x: cardX + 0.15, y: cardY + 0.15, w: cardW - 0.3, h: 0.25,
                fontSize: 10, color: COLORS.accent,
                bold: true, charSpacing: 2,
            });

            // Event label
            slide.addText(event.label, {
                x: cardX + 0.15, y: cardY + 0.45, w: cardW - 0.3, h: 0.35,
                fontSize: 13, color: COLORS.white,
                bold: true, fit: 'shrink', valign: 'top',
            });

            // Event detail
            slide.addText(event.detail, {
                x: cardX + 0.15, y: cardY + 0.85, w: cardW - 0.3, h: 1.15,
                fontSize: 11, color: 'B0BEC5',
                fit: 'shrink', valign: 'top',
            });
        });
    }

    private addComparisonSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        this.addLightBackground(slide);

        // Header
        slide.addText(slideData.title, {
            x: 0.6, y: 0.5, w: 12.0, h: 0.85,
            fontSize: 28, color: COLORS.ink,
            bold: true, fit: 'shrink',
        });
        if (slideData.keyMessage) {
            slide.addText(this.cleanInlineText(slideData.keyMessage), {
                x: 0.6, y: 1.4, w: 12.0, h: 0.35,
                fontSize: 13, color: COLORS.muted, fit: 'shrink',
            });
        }

        const columns = this.buildComparisonColumns(slideData);
        const colW = 5.6;
        const leftX = 0.6;
        const rightX = 7.0;
        const startY = 2.0;

        // Column A title (blue capsule)
        slide.addShape('roundRect', {
            x: leftX, y: startY, w: 0.8, h: 0.42,
            rectRadius: 0.06,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accentLight, transparency: 0 },
        });
        slide.addText('A', {
            x: leftX, y: startY, w: 0.8, h: 0.42,
            fontSize: 16, color: COLORS.accent,
            bold: true, align: 'center', valign: 'middle',
        });
        slide.addText(columns.leftTitle, {
            x: leftX + 0.95, y: startY, w: colW - 1.0, h: 0.42,
            fontSize: 16, color: COLORS.ink, bold: true, fit: 'shrink', valign: 'middle',
        });

        // Column B title (violet capsule)
        slide.addShape('roundRect', {
            x: rightX, y: startY, w: 0.8, h: 0.42,
            rectRadius: 0.06,
            line: { color: COLORS.violet, transparency: 100 },
            fill: { color: COLORS.violetSoft, transparency: 0 },
        });
        slide.addText('B', {
            x: rightX, y: startY, w: 0.8, h: 0.42,
            fontSize: 16, color: COLORS.violet,
            bold: true, align: 'center', valign: 'middle',
        });
        slide.addText(columns.rightTitle, {
            x: rightX + 0.95, y: startY, w: colW - 1.0, h: 0.42,
            fontSize: 16, color: COLORS.ink, bold: true, fit: 'shrink', valign: 'middle',
        });

        // Center divider
        slide.addShape('rect', {
            x: 6.55, y: startY + 0.6, w: 0.03, h: 4.0,
            line: { color: COLORS.panelSoft, transparency: 100 },
            fill: { color: COLORS.panelSoft, transparency: 0 },
        });

        // Left items with blue dots
        columns.leftItems.slice(0, 5).forEach((item, i) => {
            const y = startY + 0.7 + i * 0.72;
            slide.addShape('ellipse', {
                x: leftX + 0.05, y: y + 0.15, w: 0.14, h: 0.14,
                line: { color: COLORS.accent, transparency: 100 },
                fill: { color: COLORS.accent, transparency: 0 },
            });
            slide.addText(item, {
                x: leftX + 0.35, y, w: colW - 0.4, h: 0.5,
                fontSize: 13, color: COLORS.slate,
                fit: 'shrink', valign: 'middle',
            });
        });

        // Right items with violet dots
        columns.rightItems.slice(0, 5).forEach((item, i) => {
            const y = startY + 0.7 + i * 0.72;
            slide.addShape('ellipse', {
                x: rightX + 0.05, y: y + 0.15, w: 0.14, h: 0.14,
                line: { color: COLORS.violet, transparency: 100 },
                fill: { color: COLORS.violet, transparency: 0 },
            });
            slide.addText(item, {
                x: rightX + 0.35, y, w: colW - 0.4, h: 0.5,
                fontSize: 13, color: COLORS.slate,
                fit: 'shrink', valign: 'middle',
            });
        });
    }

    private addProcessSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        this.addLightBackground(slide);
        this.addStandardHeader(slide, slideData, this.getRoleTagLabel('process', slideData), COLORS.success);

        const steps = this.deduplicateBullets(slideData.bullets).slice(0, 4);
        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));
        const stepWidth = steps.length <= 3 ? 3.35 : 2.45;
        const gap = steps.length <= 3 ? 0.42 : 0.24;
        const totalWidth = steps.length * stepWidth + Math.max(0, steps.length - 1) * gap;
        const startX = (SLIDE_WIDTH - totalWidth) / 2;

        steps.forEach((step, index) => {
            const x = startX + index * (stepWidth + gap);
            slide.addShape('roundRect', {
                x,
                y: 2.65,
                w: stepWidth,
                h: 2.0,
                line: { color: COLORS.line, transparency: 30 },
                fill: { color: COLORS.panel, transparency: 0 },
            });
            slide.addShape('roundRect', {
                x: x + 0.18,
                y: 2.82,
                w: 0.75,
                h: 0.45,
                line: { color: COLORS.success, transparency: 100 },
                fill: { color: COLORS.successSoft, transparency: 0 },
            });
            slide.addText(isCjk ? `步骤 ${index + 1}` : `Step ${index + 1}`, {
                x: x + 0.28,
                y: 2.95,
                w: 0.55,
                h: 0.14,
                fontSize: 10,
                color: COLORS.success,
                bold: true,
                align: 'center',
            });
            slide.addText(step, {
                x: x + 0.2,
                y: 3.47,
                w: stepWidth - 0.4,
                h: 0.88,
                fontSize: 16,
                color: COLORS.ink,
                bold: true,
                fit: 'shrink',
                valign: 'middle',
                align: 'center',
            });
            if (index < steps.length - 1) {
                slide.addText('->', {
                    x: x + stepWidth + 0.05,
                    y: 3.42,
                    w: gap - 0.1,
                    h: 0.28,
                    fontSize: 18,
                    color: COLORS.success,
                    bold: true,
                    align: 'center',
                });
            }
        });
    }

    private addDataHighlightSlide(slide: pptxgen.Slide, slideData: SlideContent, config: PptRenderConfig): void {
        const heroImage = slideData.images[0];
        if (heroImage && config.templateStyle) {
            this.addImage(slide, heroImage, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
            slide.addShape('rect', {
                x: 0,
                y: 0,
                w: SLIDE_WIDTH,
                h: SLIDE_HEIGHT,
                line: { color: COLORS.ink, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 48 },
            });
        } else {
            this.addDarkBackground(slide);
        }

        this.addSectionTag(slide, this.getRoleTagLabel('data_highlight', slideData), COLORS.accent, 0.88, 0.7, 1.85, false);
        slide.addText(slideData.title, {
            x: 0.9,
            y: 1.4,
            w: 4.7,
            h: 0.75,
            fontSize: 22,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
        });

        const headline = this.extractKeyFigure(slideData) || this.cleanInlineText(slideData.keyMessage || slideData.summary || slideData.title);
        slide.addText(headline, {
            x: 0.92,
            y: 2.35,
            w: 4.95,
            h: 1.75,
            fontSize: this.figureFontSize(headline),
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
            valign: 'middle',
        });

        slide.addShape('roundRect', {
            x: 6.35,
            y: 1.4,
            w: 5.95,
            h: 4.95,
            line: { color: COLORS.white, transparency: 100 },
            fill: { color: COLORS.white, transparency: 84 },
        });
        slide.addText(this.buildBodyRows(slideData, false), {
            x: 6.7,
            y: 1.92,
            w: 5.22,
            h: 3.95,
            fontSize: 16,
            color: COLORS.paper,
            fit: 'shrink',
            valign: 'top',
        });
    }

    private addSummarySlide(slide: pptxgen.Slide, slideData: SlideContent, brief?: DeckBrief): void {
        // Dark gradient background matching HTML summary-slide
        slide.background = { color: COLORS.ink };
        slide.addShape('rect', {
            x: 0, y: 0, w: SLIDE_WIDTH, h: SLIDE_HEIGHT,
            line: { color: '1E293B', transparency: 100 },
            fill: { color: '1E293B', transparency: 50 },
        });
        // Radial accent glow
        slide.addShape('ellipse', {
            x: 6.0, y: -1.0, w: 7.0, h: 5.0,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 90 },
        });

        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));

        // SUMMARY badge (green)
        slide.addText(isCjk ? '\u6838\u5FC3\u603B\u7ED3' : 'SUMMARY', {
            x: 0.75, y: 0.65, w: 2.5, h: 0.3,
            fontSize: 11, color: COLORS.success,
            bold: true, charSpacing: 3,
        });

        // Title
        slide.addText(slideData.title, {
            x: 0.75, y: 1.1, w: 11.5, h: 0.85,
            fontSize: 30, color: COLORS.white,
            bold: true, fit: 'shrink',
        });

        // Key message
        if (slideData.keyMessage) {
            slide.addText(this.cleanInlineText(slideData.keyMessage), {
                x: 0.75, y: 2.0, w: 11.5, h: 0.4,
                fontSize: 16, color: 'B0BEC5', fit: 'shrink',
            });
        }

        // Items with check marks in semi-transparent cards
        const cards = this.deduplicateBullets(
            slideData.bullets.length > 0 ? slideData.bullets : brief?.coreTakeaways || [],
        ).slice(0, 6);

        const startY = slideData.keyMessage ? 2.65 : 2.2;
        cards.forEach((item, index) => {
            const y = startY + index * 0.72;

            // Semi-transparent card
            slide.addShape('roundRect', {
                x: 0.75, y, w: 11.5, h: 0.58,
                rectRadius: 0.1,
                line: { color: COLORS.white, width: 0.5, transparency: 92 },
                fill: { color: COLORS.white, transparency: 96 },
            });

            // Check icon
            slide.addText('\u2713', {
                x: 0.95, y, w: 0.4, h: 0.58,
                fontSize: 16, color: COLORS.success,
                bold: true, align: 'center', valign: 'middle',
            });

            // Item text
            slide.addText(item, {
                x: 1.45, y, w: 10.5, h: 0.58,
                fontSize: 15, color: 'D9E2EC',
                fit: 'shrink', valign: 'middle',
            });
        });
    }

    private addNextStepSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        // Dark gradient background matching HTML
        slide.background = { color: COLORS.ink };
        slide.addShape('rect', {
            x: 0, y: 0, w: SLIDE_WIDTH, h: SLIDE_HEIGHT,
            line: { color: '1E293B', transparency: 100 },
            fill: { color: '1E293B', transparency: 50 },
        });
        slide.addShape('ellipse', {
            x: 6.0, y: -1.0, w: 7.0, h: 5.0,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 90 },
        });

        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));

        // NEXT STEPS badge (amber)
        slide.addText(isCjk ? '\u4E0B\u4E00\u6B65' : 'NEXT STEPS', {
            x: 0.75, y: 0.65, w: 2.5, h: 0.3,
            fontSize: 11, color: COLORS.amber,
            bold: true, charSpacing: 3,
        });

        // Title
        slide.addText(slideData.title, {
            x: 0.75, y: 1.1, w: 11.5, h: 0.85,
            fontSize: 30, color: COLORS.white,
            bold: true, fit: 'shrink',
        });

        // Key message
        if (slideData.keyMessage) {
            slide.addText(this.cleanInlineText(slideData.keyMessage), {
                x: 0.75, y: 2.0, w: 11.5, h: 0.4,
                fontSize: 16, color: 'B0BEC5', fit: 'shrink',
            });
        }

        // Action items with arrows in semi-transparent cards
        const actions = this.deduplicateBullets(slideData.bullets).slice(0, 6);
        const startY = slideData.keyMessage ? 2.65 : 2.2;

        actions.forEach((item, index) => {
            const y = startY + index * 0.72;

            // Semi-transparent card
            slide.addShape('roundRect', {
                x: 0.75, y, w: 11.5, h: 0.58,
                rectRadius: 0.1,
                line: { color: COLORS.white, width: 0.5, transparency: 92 },
                fill: { color: COLORS.white, transparency: 96 },
            });

            // Arrow icon
            slide.addText('\u2192', {
                x: 0.95, y, w: 0.4, h: 0.58,
                fontSize: 16, color: COLORS.amber,
                bold: true, align: 'center', valign: 'middle',
            });

            // Item text
            slide.addText(item, {
                x: 1.45, y, w: 10.5, h: 0.58,
                fontSize: 15, color: 'D9E2EC',
                fit: 'shrink', valign: 'middle',
            });
        });
    }

    private addKeyInsightSlide(slide: pptxgen.Slide, slideData: SlideContent, config: PptRenderConfig): void {
        const heroImage = slideData.images[0];
        if ((config.imageOnlyMode || slideData.layout === 'image_only') && heroImage && config.templateStyle) {
            this.addImage(slide, heroImage, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
            slide.addShape('rect', {
                x: 0,
                y: 0,
                w: SLIDE_WIDTH,
                h: SLIDE_HEIGHT,
                line: { color: COLORS.ink, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 42 },
            });
        } else {
            this.addDarkBackground(slide);
            if (heroImage && config.templateStyle) {
                this.addImage(slide, heroImage, 8.3, 0.82, 4.2, 5.65);
                slide.addShape('roundRect', {
                    x: 8.3,
                    y: 0.82,
                    w: 4.2,
                    h: 5.65,
                    line: { color: COLORS.white, transparency: 100 },
                    fill: { color: COLORS.ink, transparency: 64 },
                });
            }
        }

        this.addSectionTag(slide, this.getRoleTagLabel('key_insight', slideData), COLORS.accent, 0.88, 0.74, 1.65, false);
        slide.addText(slideData.title, {
            x: 0.9,
            y: 1.32,
            w: 6.6,
            h: 0.75,
            fontSize: 22,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
        });

        const keyMessage = this.cleanInlineText(slideData.keyMessage || slideData.summary || slideData.title);
        slide.addText(keyMessage, {
            x: 0.92,
            y: 2.1,
            w: 6.7,
            h: 1.5,
            fontSize: 24,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
            valign: 'middle',
        });

        this.deduplicateBullets(slideData.bullets)
            .slice(0, 3)
            .forEach((item, index) => {
                slide.addShape('roundRect', {
                    x: 0.95,
                    y: 4.08 + index * 0.68,
                    w: 6.45,
                    h: 0.5,
                    line: { color: COLORS.white, transparency: 100 },
                    fill: { color: COLORS.white, transparency: 88 },
                });
                slide.addText(item, {
                    x: 1.18,
                    y: 4.23 + index * 0.68,
                    w: 5.95,
                    h: 0.15,
                    fontSize: 13,
                    color: COLORS.paper,
                    fit: 'shrink',
                });
            });
    }

    private addContentSlide(slide: pptxgen.Slide, slideData: SlideContent, config: PptRenderConfig): void {
        const heroImage = slideData.images[0];
        const hasImage = Boolean(heroImage && config.templateStyle);
        const bullets = this.deduplicateBullets(slideData.bullets);

        // Light background
        this.addLightBackground(slide);

        // Top progress bar
        const accentHue = COLORS.accent;
        slide.addShape('rect', {
            x: 0, y: 0, w: SLIDE_WIDTH, h: 0.06,
            line: { color: COLORS.panelSoft, transparency: 100 },
            fill: { color: COLORS.panelSoft, transparency: 0 },
        });
        slide.addShape('rect', {
            x: 0, y: 0, w: SLIDE_WIDTH * 0.3, h: 0.06,
            line: { color: accentHue, transparency: 100 },
            fill: { color: accentHue, transparency: 0 },
        });

        // Page number (top-right)
        const pageInfo = slideData.sourceIndex !== undefined ? slideData.sourceIndex : 0;
        slide.addText(`${String(pageInfo + 1).padStart(2, '0')}`, {
            x: 11.6, y: 0.22, w: 0.6, h: 0.22,
            fontSize: 11, color: COLORS.muted, align: 'right',
        });

        // Title (large, bold)
        const contentW = hasImage ? 6.8 : 11.5;
        slide.addText(slideData.title, {
            x: 0.6, y: 0.55, w: contentW, h: 0.85,
            fontSize: 28, color: COLORS.ink,
            bold: true, fit: 'shrink', valign: 'middle',
        });

        // Key message with left border decoration
        const keyMessage = this.cleanInlineText(slideData.keyMessage || slideData.summary || '');
        if (keyMessage) {
            slide.addShape('rect', {
                x: 0.6, y: 1.55, w: 0.05, h: 0.5,
                line: { color: COLORS.line, transparency: 100 },
                fill: { color: COLORS.line, transparency: 0 },
            });
            slide.addText(keyMessage, {
                x: 0.82, y: 1.55, w: contentW - 0.22, h: 0.5,
                fontSize: 13, color: COLORS.muted,
                fit: 'shrink', valign: 'middle',
            });
        }

        // Card-style bullets
        const bulletStartY = keyMessage ? 2.25 : 1.65;
        const maxBullets = Math.min(bullets.length, 6);
        const bulletCardH = 0.52;
        const bulletGap = 0.12;

        for (let i = 0; i < maxBullets; i++) {
            const normalized = this.normalizeBullet(bullets[i]);
            const y = bulletStartY + i * (bulletCardH + bulletGap);

            // Card background
            slide.addShape('roundRect', {
                x: 0.6, y, w: contentW, h: bulletCardH,
                rectRadius: 0.08,
                line: { color: 'F1F5F9', transparency: 0 },
                fill: { color: COLORS.white, transparency: 0 },
            });

            // Numbered marker
            slide.addShape('roundRect', {
                x: 0.72, y: y + 0.08, w: 0.36, h: 0.36,
                rectRadius: 0.06,
                line: { color: COLORS.accent, transparency: 100 },
                fill: { color: COLORS.accent, transparency: 0 },
            });
            slide.addText(String(i + 1), {
                x: 0.72, y: y + 0.08, w: 0.36, h: 0.36,
                fontSize: 12, color: COLORS.white,
                bold: true, align: 'center', valign: 'middle',
            });

            // Bullet text
            slide.addText(normalized.text, {
                x: 1.22, y: y + 0.06, w: contentW - 0.8, h: bulletCardH - 0.12,
                fontSize: 13, color: COLORS.slate,
                fit: 'shrink', valign: 'middle',
            });
        }

        // Right side: image or accent decoration block
        if (hasImage) {
            this.addImage(slide, heroImage!, 7.8, 0.6, 5.0, 6.0);
            // Subtle overlay
            slide.addShape('rect', {
                x: 7.8, y: 6.2, w: 5.0, h: 0.4,
                line: { color: COLORS.ink, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 80 },
            });
        } else if (bullets.length > 0) {
            // Decorative accent block (no image)
            slide.addShape('roundRect', {
                x: 9.5, y: 2.0, w: 2.8, h: 2.8,
                rectRadius: 0.2,
                line: { color: COLORS.accentSoft, transparency: 100 },
                fill: { color: COLORS.accentSoft, transparency: 40 },
            });
            slide.addText('\u2726', {
                x: 9.5, y: 2.0, w: 2.8, h: 2.8,
                fontSize: 48, color: COLORS.accent,
                align: 'center', valign: 'middle',
                transparency: 60,
            });
        }
    }

    private addFooter(
        slide: pptxgen.Slide,
        slideData: SlideContent,
        page: number,
        totalSlides: number,
        config: PptRenderConfig,
        role: SlideRole,
    ): void {
        const darkMode = role === 'section_divider' || role === 'key_insight' || role === 'data_highlight'
            || role === 'agenda' || role === 'timeline' || role === 'summary' || role === 'next_step';
        const footerColor = darkMode ? COLORS.white : COLORS.muted;
        const pageFill = darkMode ? COLORS.white : COLORS.ink;
        const pageText = darkMode ? COLORS.ink : COLORS.white;

        slide.addShape('roundRect', {
            x: 12.07,
            y: 6.88,
            w: 0.72,
            h: 0.34,
            line: { color: pageFill, transparency: 100 },
            fill: { color: pageFill, transparency: 0 },
        });
        slide.addText(`${page}`, {
            x: 12.28,
            y: 6.97,
            w: 0.24,
            h: 0.12,
            fontSize: 10,
            color: pageText,
            bold: true,
            align: 'center',
        });
        slide.addText(`/ ${totalSlides}`, {
            x: 12.77,
            y: 6.97,
            w: 0.3,
            h: 0.12,
            fontSize: 10,
            color: footerColor,
            align: 'left',
        });

        if (config.showSourceRefs && slideData.sourceRefs && slideData.sourceRefs.length > 0) {
            slide.addText(`${this.isMostlyCjk(this.collectSlideLanguageSeed(slideData)) ? '来源' : 'Sources'}: ${slideData.sourceRefs.join(', ')}`, {
                x: 0.88,
                y: 6.94,
                w: 2.8,
                h: 0.16,
                fontSize: 9,
                color: footerColor,
                fit: 'shrink',
            });
        }
    }

    private addSectionTag(
        slide: pptxgen.Slide,
        label: string,
        fillColor: string,
        x: number,
        y: number,
        w: number,
        darkText: boolean,
    ): void {
        slide.addShape('roundRect', {
            x,
            y,
            w,
            h: 0.38,
            line: { color: fillColor, transparency: 100 },
            fill: { color: fillColor, transparency: 0 },
        });
        slide.addText(label, {
            x: x + 0.16,
            y: y + 0.1,
            w: w - 0.32,
            h: 0.15,
            fontSize: 10,
            color: darkText ? COLORS.ink : COLORS.white,
            bold: true,
            align: 'center',
            fit: 'shrink',
        });
    }

    private addStandardHeader(slide: pptxgen.Slide, slideData: SlideContent, tag: string, tagColor: string): void {
        this.addSectionTag(slide, tag, tagColor, 0.85, 0.62, Math.max(1.25, tag.length * 0.18 + 0.6), false);
        slide.addText(slideData.title, {
            x: 0.86,
            y: 1.22,
            w: 11.5,
            h: 0.82,
            fontSize: 23,
            color: COLORS.ink,
            bold: true,
            fit: 'shrink',
        });

        const summary = this.cleanInlineText(slideData.summary || slideData.keyMessage || '');
        if (summary) {
            slide.addText(summary, {
                x: 0.9,
                y: 2.0,
                w: 11.2,
                h: 0.45,
                fontSize: 13,
                color: COLORS.slate,
                fit: 'shrink',
            });
        }
    }

    private addComparisonColumn(
        slide: pptxgen.Slide,
        x: number,
        title: string,
        items: string[],
        fillColor: string,
    ): void {
        slide.addShape('roundRect', {
            x,
            y: 2.55,
            w: 5.55,
            h: 3.95,
            line: { color: COLORS.line, transparency: 30 },
            fill: { color: fillColor, transparency: 20 },
        });
        slide.addText(title, {
            x: x + 0.24,
            y: 2.82,
            w: 5.0,
            h: 0.34,
            fontSize: 18,
            color: COLORS.ink,
            bold: true,
            fit: 'shrink',
        });

        items.slice(0, 4).forEach((item, index) => {
            slide.addShape('roundRect', {
                x: x + 0.24,
                y: 3.34 + index * 0.72,
                w: 5.05,
                h: 0.52,
                line: { color: COLORS.white, transparency: 100 },
                fill: { color: COLORS.panel, transparency: 25 },
            });
            slide.addText(item, {
                x: x + 0.4,
                y: 3.49 + index * 0.72,
                w: 4.7,
                h: 0.16,
                fontSize: 13,
                color: COLORS.ink,
                fit: 'shrink',
            });
        });
    }

    private addLightBackground(slide: pptxgen.Slide): void {
        slide.background = { color: COLORS.paper };
    }

    private addDarkBackground(slide: pptxgen.Slide): void {
        slide.background = { color: COLORS.ink };
        // Radial gradient simulation: accent glow top-right
        slide.addShape('ellipse', {
            x: 1.5,
            y: -1.5,
            w: 6.0,
            h: 5.0,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 85 },
        });
        // Violet glow right
        slide.addShape('ellipse', {
            x: 8.0,
            y: -1.0,
            w: 5.0,
            h: 4.0,
            line: { color: COLORS.violet, transparency: 100 },
            fill: { color: COLORS.violet, transparency: 88 },
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

    private paginateSlides(slides: SlideContent[], maxBulletsPerSlide: number): SlideContent[] {
        const paginated: SlideContent[] = [];
        for (const slide of slides) {
            const role = slide.slideRole || 'content';
            const canSplit = role === 'content' || role === 'data_highlight';
            if (!canSplit || slide.bullets.length <= maxBulletsPerSlide) {
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
                });
                chunkIndex += 1;
            }
        }
        return paginated;
    }

    private buildBodyRows(
        slideData: SlideContent,
        includeSummary: boolean,
    ): Array<{ text: string; options: Record<string, unknown> }> {
        const rows: Array<{ text: string; options: Record<string, unknown> }> = [];
        const bullets = this.deduplicateBullets(slideData.bullets);
        const summary = this.cleanInlineText(slideData.summary || '');
        if (includeSummary && summary && !this.isSummaryRedundant(summary, bullets, slideData.title)) {
            rows.push({
                text: summary,
                options: { breakLine: true, color: COLORS.accentDark, fontSize: 14, bold: true },
            });
        }

        if (bullets.length === 0) {
            rows.push({
                text: this.cleanInlineText(slideData.keyMessage || slideData.title),
                options: { color: COLORS.ink, fontSize: 16 },
            });
            return rows;
        }

        bullets.forEach((raw, index) => {
            const normalized = this.normalizeBullet(raw);
            rows.push({
                text: `${'  '.repeat(normalized.level)}- ${normalized.text}`,
                options: {
                    breakLine: index < bullets.length - 1,
                    color: COLORS.ink,
                    fontSize: Math.max(13, 17 - normalized.level),
                },
            });
        });
        return rows;
    }

    private buildComparisonColumns(slideData: SlideContent): ComparisonColumns {
        const match = slideData.title.match(/(.+?)\s+(?:vs|compare|contrast|and)\s+(.+)/i);
        const bullets = this.deduplicateBullets(slideData.bullets);
        const midpoint = Math.max(1, Math.ceil(bullets.length / 2));
        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));
        return {
            leftTitle: match?.[1]?.trim() || (isCjk ? '方向 A' : 'Perspective A'),
            rightTitle: match?.[2]?.trim() || (isCjk ? '方向 B' : 'Perspective B'),
            leftItems: bullets.slice(0, midpoint),
            rightItems: bullets.slice(midpoint),
        };
    }

    private buildTimelineEvents(slideData: SlideContent): TimelineEvent[] {
        const events: TimelineEvent[] = [];
        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));
        this.deduplicateBullets(slideData.bullets).forEach((item, index) => {
            const trimmed = this.cleanInlineText(item);
            const match = trimmed.match(/^((?:19|20)\d{2}|Q[1-4]\s+\d{4}|Phase\s+\d+|Step\s+\d+)\s*[:\-\uFF1A]?\s*(.*)$/i);
            if (match) {
                events.push({ label: match[1], detail: match[2] || trimmed });
                return;
            }
            const parts = trimmed.split(/[:\uFF1A]/);
            if (parts.length >= 2) {
                events.push({ label: parts[0], detail: parts.slice(1).join(':').trim() });
                return;
            }
            events.push({ label: isCjk ? `节点 ${index + 1}` : `Point ${index + 1}`, detail: trimmed });
        });
        return events.length > 0
            ? events
            : [{ label: isCjk ? '阶段 1' : 'Stage 1', detail: this.cleanInlineText(slideData.keyMessage || slideData.title) }];
    }

    private buildCoverTakeaways(brief: DeckBrief | undefined, slides: SlideContent[]): string[] {
        return this.pickAudienceFacingItems(
            (brief?.coreTakeaways || slides.map((item) => item.keyMessage || item.title)).filter(Boolean),
            3,
        );
    }

    private buildCoverTopics(brief: DeckBrief | undefined, slides: SlideContent[]): string[] {
        return this.pickAudienceFacingItems(
            (brief?.chapterTitles || slides.map((item) => item.title)).filter(Boolean),
            4,
        );
    }

    private buildAgendaFocusItems(brief: DeckBrief | undefined, slideData: SlideContent): string[] {
        return this.pickAudienceFacingItems(
            [
                ...(brief?.coreTakeaways || []),
                slideData.summary || '',
                slideData.keyMessage || '',
                ...slideData.bullets,
            ],
            4,
        );
    }

    private pickAudienceFacingItems(items: string[], maxItems: number): string[] {
        const unique = new Set<string>();
        const results: string[] = [];

        for (const raw of items) {
            const text = this.cleanInlineText(raw || '');
            if (!text || this.isPresenterArtifactText(text)) {
                continue;
            }

            const key = this.normalizeForCompare(text);
            if (!key || unique.has(key)) {
                continue;
            }

            unique.add(key);
            results.push(text);
            if (results.length >= maxItems) {
                break;
            }
        }

        return results;
    }

    private getRoleTagLabel(role: SlideRole, slideData: SlideContent): string {
        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));
        if (isCjk) {
            switch (role) {
                case 'timeline':
                    return '时间线';
                case 'comparison':
                    return '对比分析';
                case 'process':
                    return '流程拆解';
                case 'data_highlight':
                    return '重点数据';
                case 'summary':
                    return '核心总结';
                case 'next_step':
                    return '下一步';
                case 'key_insight':
                    return '核心观点';
                case 'content':
                default:
                    return '内容页';
            }
        }

        switch (role) {
            case 'timeline':
                return 'Timeline';
            case 'comparison':
                return 'Comparison';
            case 'process':
                return 'Process';
            case 'data_highlight':
                return 'Data Highlight';
            case 'summary':
                return 'Summary';
            case 'next_step':
                return 'Action Plan';
            case 'key_insight':
                return 'Key Insight';
            case 'content':
            default:
                return 'Content';
        }
    }

    private collectSlideLanguageSeed(slideData: SlideContent): string {
        return [
            slideData.title,
            slideData.summary || '',
            slideData.keyMessage || '',
            slideData.breadcrumb || '',
            slideData.bullets.join(' '),
        ].join(' ');
    }

    private isPresenterArtifactText(text: string): boolean {
        const normalized = this.cleanInlineText(text).toLowerCase();
        if (!normalized) {
            return false;
        }

        const patterns = [
            /\bhelp .* understand\b/,
            /\bpresentation framing\b/,
            /\boverview-driven\b/,
            /\btimeline-focused\b/,
            /\bcomparison-driven\b/,
            /\bprocess-oriented\b/,
            /\bargument-led\b/,
            /\baudience\s*:/,
            /\bformat\s*:/,
            /\bfocus\s*:/,
            /\bstyle\s*:/,
            /\blength\s*:/,
            /\bcontent slides\b/,
            /\bai-synthesized deck\b/,
        ];

        return patterns.some((pattern) => pattern.test(normalized));
    }

    private isMostlyCjk(text: string): boolean {
        const cleaned = this.cleanInlineText(text);
        const matches = cleaned.match(/[\u4e00-\u9fff]/g) || [];
        return matches.length >= Math.max(2, Math.floor(cleaned.length / 5));
    }

    private extractKeyFigure(slideData: SlideContent): string {
        const text = `${slideData.title} ${slideData.keyMessage || ''} ${slideData.summary || ''} ${slideData.bullets.join(' ')}`;
        const match = text.match(/\b\d+(?:\.\d+)?%?\b/);
        return match ? match[0] : '';
    }

    private figureFontSize(text: string): number {
        if (text.length <= 8) return 34;
        if (text.length <= 18) return 28;
        if (text.length <= 36) return 24;
        return 20;
    }

    private deduplicateBullets(bullets: string[]): string[] {
        const seen = new Set<string>();
        const output: string[] = [];
        for (const raw of bullets) {
            const normalized = this.normalizeBullet(raw);
            if (!normalized.text) continue;
            const key = this.normalizeForCompare(normalized.text);
            if (!key || seen.has(key)) continue;
            seen.add(key);
            output.push(raw);
        }
        return output;
    }

    private isSummaryRedundant(summary: string, bullets: string[], title: string): boolean {
        const summaryNorm = this.normalizeForCompare(summary);
        if (!summaryNorm) return true;
        const withoutTitle = this.normalizeForCompare(
            summary.replace(new RegExp(`^\\s*${this.escapeRegExp(title)}\\s*[:锛氾紝銆乗\-]*\\s*`, 'i'), ''),
        );
        for (const bullet of bullets) {
            const bulletNorm = this.normalizeForCompare(this.normalizeBullet(bullet).text);
            if (!bulletNorm) continue;
            if (summaryNorm === bulletNorm || withoutTitle === bulletNorm) return true;
            if (summaryNorm.length >= 8 && bulletNorm.length >= 8 && (summaryNorm.includes(bulletNorm) || bulletNorm.includes(summaryNorm))) {
                return true;
            }
        }
        return false;
    }

    private normalizeBullet(raw: string): { text: string; level: number } {
        const expanded = raw.replace(/\t/g, '  ');
        const leadingSpaces = (expanded.match(/^(\s*)/)?.[1].length || 0);
        const level = Math.min(3, Math.floor(leadingSpaces / 2));
        const text = expanded.trim().replace(/^[-*]\s+/, '');
        return { text, level };
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
}

