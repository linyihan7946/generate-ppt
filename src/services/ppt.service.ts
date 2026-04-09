import pptxgen from 'pptxgenjs';
import { DeckBrief, DocumentData, SlideContent, SlideRole } from '../types';

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;
const COLORS = {
    ink: '0F172A',
    slate: '334155',
    muted: '64748B',
    line: 'CBD5E1',
    paper: 'F8FAFC',
    panel: 'FFFFFF',
    panelSoft: 'E2E8F0',
    accent: '0EA5E9',
    accentDark: '0369A1',
    accentSoft: 'BAE6FD',
    success: '16A34A',
    successSoft: 'DCFCE7',
    amber: 'D97706',
    amberSoft: 'FEF3C7',
    violetSoft: 'EDE9FE',
    white: 'FFFFFF',
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

        if (coverImage && config.templateStyle) {
            this.addImage(slide, coverImage, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT);
            slide.addShape('rect', {
                x: 0,
                y: 0,
                w: SLIDE_WIDTH,
                h: SLIDE_HEIGHT,
                line: { color: COLORS.ink, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 38 },
            });
        } else {
            this.addDarkBackground(slide);
        }

        this.addSectionTag(slide, isCjk ? '主题演示' : 'Presentation', COLORS.accent, 0.82, 0.74, 2.3, false);
        slide.addText(data.title, {
            x: 0.8,
            y: 1.65,
            w: 8.9,
            h: 1.85,
            fontSize: 28,
            color: COLORS.white,
            bold: true,
            fit: 'shrink',
            valign: 'middle',
        });

        const goal = this.cleanInlineText(brief?.deckGoal || '');
        if (goal) {
            slide.addText(goal, {
                x: 0.82,
                y: 3.65,
                w: 7.6,
                h: 0.55,
                fontSize: 15,
                color: 'DBEAFE',
                fit: 'shrink',
            });
        }

        const takeaways = this.buildCoverTakeaways(brief, slides).slice(0, 3);
        takeaways.forEach((item, index) => {
            slide.addShape('roundRect', {
                x: 0.85,
                y: 4.45 + index * 0.62,
                w: 4.75,
                h: 0.42,
                line: { color: COLORS.white, transparency: 100 },
                fill: { color: COLORS.white, transparency: 88 },
            });
            slide.addText(item, {
                x: 1.02,
                y: 4.56 + index * 0.62,
                w: 4.35,
                h: 0.16,
                fontSize: 11,
                color: COLORS.paper,
                fit: 'shrink',
            });
        });

        const coverTopics = this.buildCoverTopics(brief, slides).slice(0, 4);
        if (coverTopics.length > 0) {
            slide.addShape('roundRect', {
                x: 8.95,
                y: 1.18,
                w: 3.6,
                h: 4.12,
                line: { color: COLORS.white, transparency: 100 },
                fill: { color: COLORS.white, transparency: 82 },
            });
            slide.addText(isCjk ? '内容脉络' : 'Outline', {
                x: 9.22,
                y: 1.46,
                w: 2.3,
                h: 0.24,
                fontSize: 15,
                color: COLORS.ink,
                bold: true,
                fit: 'shrink',
            });

            coverTopics.forEach((item, index) => {
                slide.addShape('roundRect', {
                    x: 9.18,
                    y: 1.92 + index * 0.72,
                    w: 3.02,
                    h: 0.5,
                    line: { color: COLORS.line, transparency: 100 },
                    fill: { color: COLORS.panelSoft, transparency: 10 },
                });
                slide.addText(item, {
                    x: 9.38,
                    y: 2.08 + index * 0.72,
                    w: 2.62,
                    h: 0.16,
                    fontSize: 12,
                    color: COLORS.ink,
                    fit: 'shrink',
                });
            });
        }
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
        this.addLightBackground(slide);
        const isCjk = this.isMostlyCjk([slideData.title, slideData.summary || '', ...(brief?.coreTakeaways || [])].join(' '));
        this.addSectionTag(slide, isCjk ? '内容导航' : 'Agenda', COLORS.accent, 0.85, 0.62, 1.45, false);
        slide.addText(slideData.title, {
            x: 0.85,
            y: 1.15,
            w: 4.4,
            h: 0.75,
            fontSize: 24,
            color: COLORS.ink,
            bold: true,
            fit: 'shrink',
        });

        const summary = this.cleanInlineText(slideData.summary || slideData.keyMessage || '');
        if (summary) {
            slide.addText(summary, {
                x: 0.88,
                y: 2.02,
                w: 3.95,
                h: 0.9,
                fontSize: 13,
                color: COLORS.slate,
                fit: 'shrink',
            });
        }

        slide.addShape('roundRect', {
            x: 0.82,
            y: 3.02,
            w: 4.1,
            h: 3.45,
            line: { color: COLORS.line, transparency: 20 },
            fill: { color: COLORS.panel, transparency: 0 },
        });
        slide.addText(isCjk ? '本次聚焦' : 'Focus', {
            x: 1.05,
            y: 3.28,
            w: 1.7,
            h: 0.22,
            fontSize: 15,
            color: COLORS.ink,
            bold: true,
            fit: 'shrink',
        });
        const focusItems = this.buildAgendaFocusItems(brief, slideData).slice(0, 4);
        focusItems.forEach((item, index) => {
            slide.addShape('roundRect', {
                x: 1.0,
                y: 3.68 + index * 0.65,
                w: 3.72,
                h: 0.46,
                line: { color: COLORS.line, transparency: 100 },
                fill: { color: index % 2 === 0 ? COLORS.panelSoft : COLORS.accentSoft, transparency: 18 },
            });
            slide.addText(item, {
                x: 1.18,
                y: 3.82 + index * 0.65,
                w: 3.35,
                h: 0.15,
                fontSize: 12,
                color: COLORS.ink,
                fit: 'shrink',
            });
        });

        const agendaItems = this.deduplicateBullets(
            (brief?.chapterTitles || []).length > 0 ? brief!.chapterTitles : slideData.bullets,
        ).slice(0, 6);

        agendaItems.forEach((item, index) => {
            const y = 1.45 + index * 0.86;
            slide.addShape('roundRect', {
                x: 5.4,
                y,
                w: 6.95,
                h: 0.65,
                line: { color: COLORS.line, transparency: 35 },
                fill: { color: index % 2 === 0 ? COLORS.panel : COLORS.panelSoft, transparency: 12 },
            });
            slide.addShape('roundRect', {
                x: 5.62,
                y: y + 0.11,
                w: 0.55,
                h: 0.43,
                line: { color: COLORS.accentDark, transparency: 100 },
                fill: { color: COLORS.accent, transparency: 0 },
            });
            slide.addText(String(index + 1), {
                x: 5.81,
                y: y + 0.18,
                w: 0.18,
                h: 0.13,
                fontSize: 11,
                color: COLORS.white,
                bold: true,
                align: 'center',
            });
            slide.addText(item, {
                x: 6.36,
                y: y + 0.15,
                w: 5.62,
                h: 0.28,
                fontSize: 15,
                color: COLORS.ink,
                fit: 'shrink',
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
        this.addLightBackground(slide);
        this.addStandardHeader(slide, slideData, this.getRoleTagLabel('timeline', slideData), COLORS.amber);

        const events = this.buildTimelineEvents(slideData).slice(0, 4);
        slide.addShape('rect', {
            x: 1.15,
            y: 3.25,
            w: 10.95,
            h: 0.08,
            line: { color: COLORS.amber, transparency: 100 },
            fill: { color: COLORS.amber, transparency: 0 },
        });

        events.forEach((event, index) => {
            const x = 1.28 + index * 3.05;
            const cardY = index % 2 === 0 ? 1.75 : 3.82;
            slide.addShape('roundRect', {
                x,
                y: 3.02,
                w: 0.32,
                h: 0.32,
                line: { color: COLORS.amber, transparency: 100 },
                fill: { color: COLORS.amber, transparency: 0 },
            });
            slide.addShape('roundRect', {
                x: x - 0.18,
                y: cardY,
                w: 2.55,
                h: 1.12,
                line: { color: COLORS.line, transparency: 35 },
                fill: { color: COLORS.panel, transparency: 0 },
            });
            slide.addText(event.label, {
                x: x + 0.02,
                y: cardY + 0.12,
                w: 2.02,
                h: 0.2,
                fontSize: 11,
                color: COLORS.amber,
                bold: true,
                fit: 'shrink',
            });
            slide.addText(event.detail, {
                x: x + 0.02,
                y: cardY + 0.34,
                w: 2.16,
                h: 0.58,
                fontSize: 13,
                color: COLORS.ink,
                fit: 'shrink',
                valign: 'middle',
            });
        });
    }

    private addComparisonSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        this.addLightBackground(slide);
        this.addStandardHeader(slide, slideData, this.getRoleTagLabel('comparison', slideData), COLORS.accentDark);
        const columns = this.buildComparisonColumns(slideData);
        this.addComparisonColumn(slide, 0.92, columns.leftTitle, columns.leftItems, COLORS.accentSoft);
        this.addComparisonColumn(slide, 6.78, columns.rightTitle, columns.rightItems, COLORS.violetSoft);
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
        this.addLightBackground(slide);
        this.addStandardHeader(slide, slideData, this.getRoleTagLabel('summary', slideData), COLORS.accent);

        const cards = this.deduplicateBullets(
            slideData.bullets.length > 0 ? slideData.bullets : brief?.coreTakeaways || [],
        ).slice(0, 4);
        const isCjk = this.isMostlyCjk(this.collectSlideLanguageSeed(slideData));

        cards.forEach((item, index) => {
            const col = index % 2;
            const row = Math.floor(index / 2);
            const x = 0.92 + col * 5.95;
            const y = 2.18 + row * 1.82;
            slide.addShape('roundRect', {
                x,
                y,
                w: 5.48,
                h: 1.38,
                line: { color: COLORS.line, transparency: 28 },
                fill: { color: index % 2 === 0 ? COLORS.panel : COLORS.panelSoft, transparency: 10 },
            });
            slide.addText(isCjk ? `要点 ${index + 1}` : `Takeaway ${index + 1}`, {
                x: x + 0.24,
                y: y + 0.18,
                w: 1.1,
                h: 0.16,
                fontSize: 10,
                color: COLORS.accentDark,
                bold: true,
            });
            slide.addText(item, {
                x: x + 0.22,
                y: y + 0.48,
                w: 5.0,
                h: 0.62,
                fontSize: 16,
                color: COLORS.ink,
                fit: 'shrink',
                valign: 'middle',
            });
        });
    }

    private addNextStepSlide(slide: pptxgen.Slide, slideData: SlideContent): void {
        this.addLightBackground(slide);
        this.addStandardHeader(slide, slideData, this.getRoleTagLabel('next_step', slideData), COLORS.success);

        const actions = this.deduplicateBullets(slideData.bullets).slice(0, 5);
        actions.forEach((item, index) => {
            const y = 1.98 + index * 0.9;
            slide.addShape('roundRect', {
                x: 1.08,
                y,
                w: 11.1,
                h: 0.62,
                line: { color: COLORS.line, transparency: 35 },
                fill: { color: index % 2 === 0 ? COLORS.panel : COLORS.successSoft, transparency: 12 },
            });
            slide.addShape('roundRect', {
                x: 1.28,
                y: y + 0.1,
                w: 0.58,
                h: 0.42,
                line: { color: COLORS.success, transparency: 100 },
                fill: { color: COLORS.success, transparency: 0 },
            });
            slide.addText(String(index + 1), {
                x: 1.49,
                y: y + 0.18,
                w: 0.15,
                h: 0.14,
                fontSize: 11,
                color: COLORS.white,
                bold: true,
                align: 'center',
            });
            slide.addText(item, {
                x: 2.02,
                y: y + 0.14,
                w: 9.82,
                h: 0.26,
                fontSize: 15,
                color: COLORS.ink,
                fit: 'shrink',
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
        const useImagePanel = Boolean(heroImage && config.templateStyle);

        this.addLightBackground(slide);
        if (useImagePanel) {
            this.addImage(slide, heroImage!, 7.2, 0.86, 5.15, 5.82);
            slide.addShape('roundRect', {
                x: 7.2,
                y: 0.86,
                w: 5.15,
                h: 5.82,
                line: { color: COLORS.white, transparency: 100 },
                fill: { color: COLORS.ink, transparency: 72 },
            });
        }

        this.addSectionTag(slide, this.getRoleTagLabel('content', slideData), COLORS.accent, 0.85, 0.62, 1.35, false);
        if (slideData.breadcrumb) {
            slide.addText(slideData.breadcrumb, {
                x: 0.88,
                y: 1.05,
                w: 5.7,
                h: 0.18,
                fontSize: 10,
                color: COLORS.muted,
                fit: 'shrink',
            });
        }

        slide.addText(slideData.title, {
            x: 0.86,
            y: 1.34,
            w: useImagePanel ? 5.9 : 11.5,
            h: 0.8,
            fontSize: 22,
            color: COLORS.ink,
            bold: true,
            fit: 'shrink',
        });

        const keyMessage = this.cleanInlineText(slideData.keyMessage || slideData.summary || '');
        if (keyMessage) {
            slide.addText(keyMessage, {
                x: 0.9,
                y: 2.16,
                w: useImagePanel ? 5.9 : 11.5,
                h: 0.55,
                fontSize: 14,
                color: COLORS.accentDark,
                bold: true,
                fit: 'shrink',
            });
        }

        slide.addShape('roundRect', {
            x: 0.84,
            y: 2.92,
            w: useImagePanel ? 5.9 : 11.55,
            h: 3.36,
            line: { color: COLORS.line, transparency: 28 },
            fill: { color: COLORS.panel, transparency: 0 },
        });
        slide.addText(this.buildBodyRows(slideData, true), {
            x: 1.08,
            y: 3.3,
            w: useImagePanel ? 5.35 : 10.95,
            h: 2.7,
            fontSize: 15,
            color: COLORS.ink,
            fit: 'shrink',
            valign: 'top',
        });
    }

    private addFooter(
        slide: pptxgen.Slide,
        slideData: SlideContent,
        page: number,
        totalSlides: number,
        config: PptRenderConfig,
        role: SlideRole,
    ): void {
        const darkMode = role === 'section_divider' || role === 'key_insight' || role === 'data_highlight';
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
        slide.addShape('roundRect', {
            x: 9.4,
            y: -0.82,
            w: 4.2,
            h: 3.0,
            line: { color: COLORS.accentSoft, transparency: 100 },
            fill: { color: COLORS.accentSoft, transparency: 38 },
        });
        slide.addShape('roundRect', {
            x: -0.72,
            y: 5.18,
            w: 4.0,
            h: 2.8,
            line: { color: COLORS.panelSoft, transparency: 100 },
            fill: { color: COLORS.panelSoft, transparency: 28 },
        });
    }

    private addDarkBackground(slide: pptxgen.Slide): void {
        slide.background = { color: COLORS.ink };
        slide.addShape('roundRect', {
            x: 8.85,
            y: -0.92,
            w: 4.6,
            h: 3.2,
            line: { color: COLORS.accent, transparency: 100 },
            fill: { color: COLORS.accent, transparency: 74 },
        });
        slide.addShape('roundRect', {
            x: -0.9,
            y: 4.95,
            w: 4.45,
            h: 3.05,
            line: { color: COLORS.accentDark, transparency: 100 },
            fill: { color: COLORS.accentDark, transparency: 68 },
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

