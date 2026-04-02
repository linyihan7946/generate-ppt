import fs from 'fs';
import path from 'path';
import {
    DocumentData,
    QualityDimensionScore,
    QualityMetrics,
    QualityReport,
    SlideContent,
} from '../types';

export class EvaluatorService {
    private readonly logicWeight = 0.4;
    private readonly layoutWeight = 0.3;
    private readonly imageWeight = 0.3;

    evaluate(docData: DocumentData, outputPath?: string): QualityReport {
        const slides = docData.slides;
        const metrics = this.computeMetrics(slides);

        const logic = this.scoreLogic(docData, slides, metrics);
        const layout = this.scoreLayout(slides, metrics);
        const imageSemantics = this.scoreImageSemantics(slides, metrics);

        const overallScore = this.round(
            logic.weightedScore + layout.weightedScore + imageSemantics.weightedScore,
            1,
        );
        const grade = this.getGrade(overallScore);

        const keyFindings = this.collectKeyFindings(metrics, logic, layout, imageSemantics);
        const nextActions = this.collectNextActions(logic, layout, imageSemantics);

        return {
            version: 'v1',
            generatedAt: new Date().toISOString(),
            title: docData.title,
            outputPath,
            overallScore,
            grade,
            dimensions: {
                logic,
                layout,
                imageSemantics,
            },
            metrics,
            keyFindings,
            nextActions,
        };
    }

    saveReport(report: QualityReport, outputPath?: string): { jsonPath: string; markdownPath: string } {
        const outDir = outputPath ? path.dirname(outputPath) : path.resolve(process.cwd(), 'output');
        fs.mkdirSync(outDir, { recursive: true });

        const stem = outputPath
            ? path.basename(outputPath, path.extname(outputPath))
            : `presentation-${Date.now()}`;
        const jsonPath = path.join(outDir, `${stem}.quality.json`);
        const markdownPath = path.join(outDir, `${stem}.quality.md`);

        fs.writeFileSync(jsonPath, JSON.stringify(report, null, 2), 'utf-8');
        fs.writeFileSync(markdownPath, this.toMarkdown(report), 'utf-8');

        return { jsonPath, markdownPath };
    }

    private computeMetrics(slides: SlideContent[]): QualityMetrics {
        const slideCount = slides.length;
        const slideWithImageCount = slides.filter((s) => s.images.length > 0).length;
        const imageCoverage = slideCount === 0 ? 0 : slideWithImageCount / slideCount;
        const totalBullets = slides.reduce((sum, s) => sum + s.bullets.length, 0);
        const avgBulletsPerSlide = slideCount === 0 ? 0 : totalBullets / slideCount;

        const allBullets = slides.flatMap((s) => s.bullets);
        const avgBulletLength =
            allBullets.length === 0
                ? 0
                : allBullets.reduce((sum, b) => sum + this.cleanText(b).length, 0) / allBullets.length;

        const levelJumpViolations = this.countLevelJumpViolations(slides);
        const duplicateTitleCount = this.countDuplicateTitles(slides);
        const redundantContentSlideCount = slides.filter((s) => this.countIntraSlideRedundancyItems(s) > 0).length;
        const redundantContentItemCount = slides.reduce(
            (sum, slide) => sum + this.countIntraSlideRedundancyItems(slide),
            0,
        );
        const overlaySlideCount = slides.filter((s) => s.layout !== 'image_only').length;
        const imageOnlySlideCount = slides.filter((s) => s.layout === 'image_only').length;
        const overflowRiskSlideCount = this.countOverflowRiskSlides(slides);
        const promptAlignmentAvg = this.averagePromptAlignment(slides);
        const fallbackImageCount = slides.filter((s) => s.imageSource === 'ai_fallback' || s.imageSource === 'placeholder').length;

        return {
            slideCount,
            slideWithImageCount,
            imageCoverage: this.round(imageCoverage, 3),
            avgBulletsPerSlide: this.round(avgBulletsPerSlide, 2),
            avgBulletLength: this.round(avgBulletLength, 2),
            levelJumpViolations,
            duplicateTitleCount,
            redundantContentSlideCount,
            redundantContentItemCount,
            overlaySlideCount,
            imageOnlySlideCount,
            overflowRiskSlideCount,
            promptAlignmentAvg: this.round(promptAlignmentAvg, 3),
            fallbackImageCount,
        };
    }

    private scoreLogic(
        docData: DocumentData,
        slides: SlideContent[],
        metrics: QualityMetrics,
    ): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(`Total slides: ${metrics.slideCount}`);
        evidence.push(`Average bullets per slide: ${metrics.avgBulletsPerSlide}`);
        evidence.push(`Intra-slide repeated-content items: ${metrics.redundantContentItemCount}`);

        const emptyTitleCount = slides.filter((s) => this.cleanText(s.title).length === 0).length;
        if (emptyTitleCount > 0) {
            score -= emptyTitleCount * 10;
            issues.push(`${emptyTitleCount} slides have empty titles.`);
            recommendations.push('Ensure each slide has a concise and explicit title.');
        }

        if (metrics.duplicateTitleCount > 0) {
            score -= metrics.duplicateTitleCount * 4;
            issues.push(`${metrics.duplicateTitleCount} duplicate titles detected.`);
            recommendations.push("Rewrite repeated titles to highlight each slide's unique point.");
        }

        if (metrics.redundantContentItemCount > 0) {
            const penalty = Math.min(
                18,
                metrics.redundantContentSlideCount * 4 + metrics.redundantContentItemCount * 2,
            );
            score -= penalty;
            issues.push(
                `${metrics.redundantContentItemCount} repeated content items detected on ${metrics.redundantContentSlideCount} slides.`,
            );
            recommendations.push(
                'Reduce same-slide repetition: avoid repeating summary in bullets and remove duplicate bullet points.',
            );
        } else {
            evidence.push('No obvious same-slide repeated content detected.');
        }

        if (metrics.levelJumpViolations > 0) {
            score -= metrics.levelJumpViolations * 6;
            issues.push(`${metrics.levelJumpViolations} hierarchy jumps detected (level gap > 1).`);
            recommendations.push('Smooth hierarchy transitions to keep narrative progression coherent.');
        }

        const sparseTopLevelCount = slides.filter(
            (s) => (s.level || 1) <= 2 && s.bullets.length === 0 && !s.summary,
        ).length;
        if (sparseTopLevelCount > 0) {
            score -= sparseTopLevelCount * 5;
            issues.push(`${sparseTopLevelCount} key slides have weak or empty content.`);
            recommendations.push('Add at least 2 concise bullets for high-level slides.');
        }

        const longBulletSlides = slides.filter((s) => s.bullets.some((b) => this.cleanText(b).length > 90)).length;
        if (longBulletSlides > 0) {
            score -= longBulletSlides * 3;
            issues.push(`${longBulletSlides} slides contain overly long bullet text.`);
            recommendations.push('Split long bullet sentences into short factual points.');
        }

        if (metrics.avgBulletsPerSlide >= 2 && metrics.avgBulletsPerSlide <= 5) {
            score += 4;
            evidence.push('Bullet density is in an acceptable range (2-5).');
        } else {
            issues.push('Bullet density is outside recommended range (2-5).');
            recommendations.push('Keep each slide around 2-5 bullets for better comprehension.');
        }

        if (this.cleanText(docData.title).length > 0) {
            evidence.push('Presentation title is present.');
        } else {
            score -= 8;
            issues.push('Presentation title is missing.');
            recommendations.push('Provide a clear presentation title.');
        }

        const finalScore = this.clamp(this.round(score, 1), 0, 100);
        return {
            name: '\u5185\u5bb9\u903b\u8f91\u6027\u548c\u5408\u7406\u6027',
            score: finalScore,
            weight: this.logicWeight,
            weightedScore: this.round(finalScore * this.logicWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private scoreLayout(slides: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(
            `Overlay slides: ${metrics.overlaySlideCount}, image-only slides: ${metrics.imageOnlySlideCount}.`,
        );
        evidence.push(`Image coverage: ${(metrics.imageCoverage * 100).toFixed(1)}%.`);

        const noImageSlides = metrics.slideCount - metrics.slideWithImageCount;
        if (noImageSlides > 0) {
            score -= noImageSlides * 10;
            issues.push(`${noImageSlides} slides are missing images.`);
            recommendations.push('Ensure each content slide has one visual to keep style consistency.');
        }

        if (metrics.overflowRiskSlideCount > 0) {
            score -= metrics.overflowRiskSlideCount * 6;
            issues.push(`${metrics.overflowRiskSlideCount} slides have high text overflow risk.`);
            recommendations.push('Reduce text density or switch those slides to image-only layout.');
        } else {
            evidence.push('No high overflow risk slide detected.');
        }

        const overlayDenseSlides = slides.filter(
            (s) => s.layout !== 'image_only' && this.totalTextLength(s) > 280,
        ).length;
        if (overlayDenseSlides > 0) {
            score -= overlayDenseSlides * 4;
            issues.push(`${overlayDenseSlides} overlay slides are text-heavy.`);
            recommendations.push('For dense content, split into multiple slides or use visual summary.');
        }

        if (metrics.imageCoverage >= 0.95) {
            score += 5;
            evidence.push('Visual coverage is high and close to template style.');
        }

        const finalScore = this.clamp(this.round(score, 1), 0, 100);
        return {
            name: '\u6392\u7248\u7f8e\u89c2\u6027',
            score: finalScore,
            weight: this.layoutWeight,
            weightedScore: this.round(finalScore * this.layoutWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private scoreImageSemantics(slides: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(`Average prompt alignment: ${(metrics.promptAlignmentAvg * 100).toFixed(1)}%.`);
        evidence.push(`Fallback/placeholder images: ${metrics.fallbackImageCount}.`);

        const missingPromptSlides = slides.filter((s) => !this.cleanText(s.imagePrompt || '')).length;
        if (missingPromptSlides > 0) {
            score -= missingPromptSlides * 5;
            issues.push(`${missingPromptSlides} slides are missing explicit image prompts.`);
            recommendations.push('Provide imagePrompt for every slide to improve semantic control.');
        }

        const lowAlignSlides = slides.filter((s) => this.promptAlignment(s) < 0.2).length;
        if (lowAlignSlides > 0) {
            score -= lowAlignSlides * 4;
            issues.push(`${lowAlignSlides} slides have weak text-image semantic alignment.`);
            recommendations.push('Rewrite image prompts using title + core bullets + context keywords.');
        }

        if (metrics.fallbackImageCount > 0) {
            score -= metrics.fallbackImageCount * 6;
            issues.push(`${metrics.fallbackImageCount} slides used fallback/placeholder images.`);
            recommendations.push('Improve prompt safety and API reliability to reduce fallback rate.');
        }

        if (metrics.imageCoverage < 0.9) {
            score -= (1 - metrics.imageCoverage) * 30;
            issues.push('Insufficient image coverage affects semantic expression.');
            recommendations.push('Ensure each slide has at least one generated or original image.');
        }

        if (metrics.promptAlignmentAvg >= 0.45) {
            score += 5;
            evidence.push('Prompt-to-content alignment is generally good.');
        }

        const finalScore = this.clamp(this.round(score, 1), 0, 100);
        return {
            name: '\u914d\u56fe\u8868\u8fbe\u51c6\u786e\u6027',
            score: finalScore,
            weight: this.imageWeight,
            weightedScore: this.round(finalScore * this.imageWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private countLevelJumpViolations(slides: SlideContent[]): number {
        let violations = 0;
        for (let i = 1; i < slides.length; i++) {
            const prev = slides[i - 1].level || 1;
            const curr = slides[i].level || 1;
            if (Math.abs(curr - prev) > 1) {
                violations += 1;
            }
        }
        return violations;
    }

    private countDuplicateTitles(slides: SlideContent[]): number {
        const count = new Map<string, number>();
        slides.forEach((s) => {
            const key = this.cleanText(s.title).toLowerCase();
            if (!key) return;
            count.set(key, (count.get(key) || 0) + 1);
        });
        let duplicates = 0;
        count.forEach((value) => {
            if (value > 1) duplicates += value - 1;
        });
        return duplicates;
    }

    private countIntraSlideRedundancyItems(slide: SlideContent): number {
        const bullets = slide.bullets.map((b) => this.cleanText(b)).filter(Boolean);
        let duplicateItems = 0;

        const bulletCountByKey = new Map<string, number>();
        bullets.forEach((bullet) => {
            const key = this.normalizeForCompare(bullet);
            if (!key) return;
            bulletCountByKey.set(key, (bulletCountByKey.get(key) || 0) + 1);
        });

        bulletCountByKey.forEach((count) => {
            if (count > 1) {
                duplicateItems += count - 1;
            }
        });

        const summary = this.cleanText(slide.summary || '');
        if (summary && this.isSummaryRedundant(summary, bullets, slide.title)) {
            duplicateItems += 1;
        }

        return duplicateItems;
    }

    private isSummaryRedundant(summary: string, bullets: string[], title: string): boolean {
        const summaryNormalized = this.normalizeForCompare(summary);
        if (!summaryNormalized) return true;

        const summaryWithoutTitleNormalized = this.normalizeForCompare(
            summary.replace(new RegExp(`^\\s*${this.escapeRegExp(title)}\\s*[:：,，。\\-]*\\s*`, 'i'), ''),
        );
        const candidates = [summaryNormalized, summaryWithoutTitleNormalized].filter(Boolean);

        for (const bullet of bullets) {
            const bulletNormalized = this.normalizeForCompare(bullet);
            if (!bulletNormalized) continue;
            for (const candidate of candidates) {
                if (!candidate) continue;
                if (candidate === bulletNormalized) {
                    return true;
                }
                if (candidate.length >= 8 && bulletNormalized.length >= 8) {
                    if (candidate.includes(bulletNormalized) || bulletNormalized.includes(candidate)) {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private normalizeForCompare(text: string): string {
        return this.cleanText(text).toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '');
    }

    private escapeRegExp(input: string): string {
        return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    private countOverflowRiskSlides(slides: SlideContent[]): number {
        return slides.filter((s) => {
            if (s.layout === 'image_only') return false;
            const titleLen = this.cleanText(s.title).length;
            const bulletCount = s.bullets.length;
            const textLen = this.totalTextLength(s);
            return titleLen > 44 || bulletCount > 6 || textLen > 280;
        }).length;
    }

    private averagePromptAlignment(slides: SlideContent[]): number {
        if (slides.length === 0) return 0;
        const sum = slides.reduce((acc, slide) => acc + this.promptAlignment(slide), 0);
        return sum / slides.length;
    }

    private promptAlignment(slide: SlideContent): number {
        const contentText = [
            slide.title,
            slide.summary || '',
            slide.bullets.join(' '),
            slide.breadcrumb || '',
        ].join(' ');
        const promptText = [slide.imageIntent || '', slide.imagePrompt || ''].join(' ');

        const contentKeywords = this.extractKeywords(contentText);
        const promptKeywords = this.extractKeywords(promptText);
        if (contentKeywords.size === 0 || promptKeywords.size === 0) {
            return 0;
        }

        let overlap = 0;
        contentKeywords.forEach((token) => {
            if (promptKeywords.has(token)) {
                overlap += 1;
            }
        });

        return overlap / contentKeywords.size;
    }

    private extractKeywords(text: string): Set<string> {
        const cleaned = this.cleanText(text).toLowerCase();
        const tokens = cleaned.match(/[\u4e00-\u9fff]{2,}|[a-z0-9]{3,}/g) || [];
        const stop = new Set([
            'the',
            'and',
            'for',
            'with',
            'that',
            'this',
            'from',
            'into',
            'presentation',
            'slide',
            'about',
            'context',
            'style',
            'modern',
            'image',
            'content',
            'summary',
            '\u5173\u952e\u70b9',
            '\u4e0a\u4e0b\u6587',
            '\u5185\u5bb9',
            '\u603b\u7ed3',
            '\u9875\u9762',
            '\u4e3b\u9898',
            '\u4ecb\u7ecd',
        ]);

        const result = new Set<string>();
        tokens.forEach((token) => {
            if (!stop.has(token)) {
                result.add(token);
            }
        });
        return result;
    }

    private totalTextLength(slide: SlideContent): number {
        return (
            this.cleanText(slide.title).length +
            this.cleanText(slide.summary || '').length +
            slide.bullets.reduce((sum, b) => sum + this.cleanText(b).length, 0)
        );
    }

    private collectKeyFindings(
        metrics: QualityMetrics,
        logic: QualityDimensionScore,
        layout: QualityDimensionScore,
        imageSemantics: QualityDimensionScore,
    ): string[] {
        const findings: string[] = [];
        findings.push(`Overall image coverage: ${(metrics.imageCoverage * 100).toFixed(1)}%.`);
        findings.push(`Prompt alignment average: ${(metrics.promptAlignmentAvg * 100).toFixed(1)}%.`);
        findings.push(
            `Dimension scores -> Logic: ${logic.score}, Layout: ${layout.score}, Image: ${imageSemantics.score}.`,
        );
        findings.push(
            `Same-slide repeated content items: ${metrics.redundantContentItemCount} on ${metrics.redundantContentSlideCount} slides.`,
        );

        if (metrics.fallbackImageCount > 0) {
            findings.push(`Fallback images used on ${metrics.fallbackImageCount} slides.`);
        }
        if (metrics.overflowRiskSlideCount > 0) {
            findings.push(`${metrics.overflowRiskSlideCount} slides are at risk of text overflow.`);
        }

        return findings;
    }

    private collectNextActions(...dimensions: QualityDimensionScore[]): string[] {
        const actions = new Set<string>();
        dimensions.forEach((dimension) => {
            dimension.recommendations.forEach((rec) => actions.add(rec));
        });
        return Array.from(actions).slice(0, 8);
    }

    private getGrade(score: number): string {
        if (score >= 90) return 'A';
        if (score >= 80) return 'B';
        if (score >= 70) return 'C';
        if (score >= 60) return 'D';
        return 'E';
    }

    private toMarkdown(report: QualityReport): string {
        const lines: string[] = [];
        lines.push(`# PPT Quality Report`);
        lines.push('');
        lines.push(`- Title: ${report.title}`);
        lines.push(`- Generated At: ${report.generatedAt}`);
        if (report.outputPath) {
            lines.push(`- Output: ${report.outputPath}`);
        }
        lines.push(`- Overall Score: **${report.overallScore} / 100**`);
        lines.push(`- Grade: **${report.grade}**`);
        lines.push('');
        lines.push('## Dimension Scores');
        lines.push('');
        lines.push('| Dimension | Score | Weight | Weighted |');
        lines.push('|---|---:|---:|---:|');
        lines.push(
            `| ${report.dimensions.logic.name} | ${report.dimensions.logic.score} | ${report.dimensions.logic.weight} | ${report.dimensions.logic.weightedScore} |`,
        );
        lines.push(
            `| ${report.dimensions.layout.name} | ${report.dimensions.layout.score} | ${report.dimensions.layout.weight} | ${report.dimensions.layout.weightedScore} |`,
        );
        lines.push(
            `| ${report.dimensions.imageSemantics.name} | ${report.dimensions.imageSemantics.score} | ${report.dimensions.imageSemantics.weight} | ${report.dimensions.imageSemantics.weightedScore} |`,
        );
        lines.push('');
        lines.push('## Metrics');
        lines.push('');
        lines.push('```json');
        lines.push(JSON.stringify(report.metrics, null, 2));
        lines.push('```');
        lines.push('');
        lines.push('## Key Findings');
        lines.push('');
        report.keyFindings.forEach((finding) => lines.push(`- ${finding}`));
        lines.push('');
        lines.push('## Suggested Next Actions');
        lines.push('');
        report.nextActions.forEach((action) => lines.push(`- ${action}`));
        lines.push('');
        return lines.join('\n');
    }

    private cleanText(input: any): string {
        if (typeof input !== 'string') return '';
        return input.replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
    }

    private clamp(value: number, min: number, max: number): number {
        return Math.max(min, Math.min(max, value));
    }

    private round(value: number, digits = 0): number {
        const factor = 10 ** digits;
        return Math.round(value * factor) / factor;
    }
}
