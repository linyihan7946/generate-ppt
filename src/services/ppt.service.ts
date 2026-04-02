import pptxgen from 'pptxgenjs';
import { DocumentData, SlideContent } from '../types';

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;

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

        const normalizedSlides = this.paginateSlides(data.slides);

        this.addTitleSlide(pres, data.title, normalizedSlides.length);
        normalizedSlides.forEach((slideData, index) => {
            this.addContentSlide(pres, slideData, index + 1, normalizedSlides.length);
        });

        await pres.writeFile({ fileName: outputPath });
        return outputPath;
    }

    private addTitleSlide(pres: pptxgen, title: string, totalSlides: number): void {
        const slide = pres.addSlide();

        slide.addShape(pres.ShapeType.rect, {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: SLIDE_HEIGHT,
            line: { color: 'F1F5F9' },
            fill: { color: 'F8FAFC' },
        });

        slide.addShape(pres.ShapeType.roundRect, {
            x: 0.75,
            y: 1.05,
            w: 11.8,
            h: 0.18,
            line: { color: '0EA5E9', transparency: 100 },
            fill: { color: '0EA5E9', transparency: 15 },
        });

        slide.addShape(pres.ShapeType.roundRect, {
            x: 10.4,
            y: 0.55,
            w: 2.2,
            h: 2.2,
            line: { color: '0EA5E9', transparency: 100 },
            fill: { color: '0EA5E9', transparency: 88 },
        });

        slide.addShape(pres.ShapeType.roundRect, {
            x: 9.2,
            y: 4.9,
            w: 3.4,
            h: 1.8,
            line: { color: '1D4ED8', transparency: 100 },
            fill: { color: '1D4ED8', transparency: 92 },
        });

        slide.addText(title, {
            x: 0.9,
            y: 1.6,
            w: 9.7,
            h: 2.3,
            fontSize: 44,
            color: '0F172A',
            bold: true,
            fit: 'shrink',
        });

        slide.addText('自动生成 · 保留文档层级逻辑 · 智能配图', {
            x: 0.95,
            y: 4.3,
            w: 8.3,
            h: 0.5,
            fontSize: 15,
            color: '334155',
        });

        slide.addText(`内容页：${totalSlides} 张`, {
            x: 0.95,
            y: 5.0,
            w: 5.0,
            h: 0.4,
            fontSize: 12,
            color: '64748B',
        });
    }

    private addContentSlide(
        pres: pptxgen,
        slideData: SlideContent,
        page: number,
        totalSlides: number,
    ): void {
        const slide = pres.addSlide();
        const hasImage = slideData.images.length > 0;
        const accentColor = this.levelColor(slideData.level);

        slide.addShape(pres.ShapeType.rect, {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: SLIDE_HEIGHT,
            line: { color: 'FFFFFF' },
            fill: { color: 'FFFFFF' },
        });

        slide.addShape(pres.ShapeType.rect, {
            x: 0,
            y: 0,
            w: SLIDE_WIDTH,
            h: 0.58,
            line: { color: '0F172A' },
            fill: { color: '0F172A' },
        });

        slide.addShape(pres.ShapeType.rect, {
            x: 0,
            y: 0.58,
            w: SLIDE_WIDTH,
            h: 0.06,
            line: { color: accentColor },
            fill: { color: accentColor },
        });

        if (slideData.breadcrumb) {
            slide.addText(slideData.breadcrumb, {
                x: 0.6,
                y: 0.12,
                w: 10.5,
                h: 0.28,
                fontSize: 10,
                color: 'CBD5E1',
                fit: 'shrink',
            });
        }

        slide.addText(slideData.title, {
            x: 0.7,
            y: 0.84,
            w: hasImage ? 7.0 : 12.0,
            h: 0.9,
            fontSize: this.titleFontSize(slideData.title),
            color: '0F172A',
            bold: true,
            fit: 'shrink',
        });

        const textBoxX = 0.8;
        const textBoxY = 1.85;
        const textBoxW = hasImage ? 6.95 : 11.8;
        const textBoxH = 4.95;
        const bulletRows = this.buildBulletRows(slideData.bullets, hasImage);

        slide.addText(bulletRows, {
            x: textBoxX,
            y: textBoxY,
            w: textBoxW,
            h: textBoxH,
            valign: 'top',
            fit: 'shrink',
        });

        if (hasImage) {
            const imageX = 8.0;
            const imageY = 1.66;
            const imageW = 4.7;
            const imageH = 4.95;

            slide.addShape(pres.ShapeType.roundRect, {
                x: imageX - 0.08,
                y: imageY - 0.08,
                w: imageW + 0.16,
                h: imageH + 0.16,
                line: { color: 'CBD5E1', pt: 1.25 },
                fill: { color: 'F8FAFC' },
            });

            this.addImage(slide, slideData.images[0], imageX, imageY, imageW, imageH);
        }

        slide.addText(`${page}/${totalSlides}`, {
            x: 11.95,
            y: 7.08,
            w: 1.0,
            h: 0.2,
            align: 'right',
            fontSize: 10,
            color: '64748B',
        });
    }

    private buildBulletRows(bullets: string[], hasImage: boolean): Array<{ text: string; options: Record<string, unknown> }> {
        if (bullets.length === 0) {
            return [
                {
                    text: '（此节点在原文中没有下级条目）',
                    options: {
                        color: '94A3B8',
                        fontSize: 15,
                    },
                },
            ];
        }

        const baseFontSize = hasImage ? 17 : 19;
        return bullets.map((raw, index) => {
            const normalized = this.normalizeBullet(raw);
            return {
                text: normalized.text,
                options: {
                    breakLine: index < bullets.length - 1,
                    bullet: { indent: 12 + normalized.level * 10 },
                    hanging: 1.8,
                    color: '1E293B',
                    fontSize: Math.max(13, baseFontSize - normalized.level),
                },
            };
        });
    }

    private normalizeBullet(raw: string): { text: string; level: number } {
        const expanded = raw.replace(/\t/g, '  ');
        const leadingSpacesMatch = expanded.match(/^(\s*)/);
        const leadingSpaces = leadingSpacesMatch ? leadingSpacesMatch[1].length : 0;
        const level = Math.min(3, Math.floor(leadingSpaces / 2));
        const text = expanded.trim().replace(/^[-*]\s+/, '');
        return { text, level };
    }

    private addImage(slide: pptxgen.Slide, image: string, x: number, y: number, w: number, h: number): void {
        if (!image) return;

        if (image.startsWith('data:image')) {
            slide.addImage({ data: image, x, y, w, h });
            return;
        }

        if (image.startsWith('http://') || image.startsWith('https://')) {
            // External URL may not be downloadable by PowerPoint, so this path is best-effort only.
            slide.addImage({ path: image, x, y, w, h });
            return;
        }

        slide.addImage({ path: image, x, y, w, h });
    }

    private paginateSlides(slides: SlideContent[]): SlideContent[] {
        const paginated: SlideContent[] = [];

        for (const slide of slides) {
            if (slide.bullets.length === 0) {
                paginated.push(slide);
                continue;
            }

            let remaining = [...slide.bullets];
            let chunkIndex = 0;

            while (remaining.length > 0) {
                const includeImage = chunkIndex === 0 && slide.images.length > 0;
                const maxBullets = includeImage ? 6 : 10;
                const chunk = remaining.slice(0, maxBullets);
                remaining = remaining.slice(maxBullets);

                paginated.push({
                    ...slide,
                    title: chunkIndex === 0 ? slide.title : `${slide.title}（续 ${chunkIndex}）`,
                    bullets: chunk,
                    images: includeImage ? slide.images : [],
                });

                chunkIndex += 1;
            }
        }

        return paginated;
    }

    private levelColor(level?: number): string {
        if (!level) return '0EA5E9';
        if (level === 1) return '0EA5E9';
        if (level === 2) return '2563EB';
        if (level === 3) return '14B8A6';
        return '6366F1';
    }

    private titleFontSize(title: string): number {
        if (title.length <= 14) return 34;
        if (title.length <= 24) return 30;
        if (title.length <= 34) return 27;
        return 24;
    }
}
