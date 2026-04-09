import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import {
    DocumentData,
    QualityDimensionScore,
    QualityMetrics,
    QualityReport,
    SlideContent,
} from '../types';

interface RenderedDeckInspection {
    renderedSlideCount: number;
    renderedMetaArtifactSlideCount: number;
    renderedInstructionalTextSlideCount: number;
    renderedMixedLanguageSlideCount: number;
}

export class EvaluatorService {
    private readonly logicWeight = 0.2;
    private readonly layoutWeight = 0.16;
    private readonly imageWeight = 0.14;
    private readonly contentRichnessWeight = 0.2;
    private readonly audienceFitWeight = 0.16;
    private readonly consistencyWeight = 0.14;

    async evaluate(docData: DocumentData, outputPath?: string): Promise<QualityReport> {
        const slides = docData.slides;
        const renderedDeck = await this.inspectRenderedDeck(outputPath, docData.title);
        const metrics = this.computeMetrics(slides, renderedDeck);

        const logic = this.scoreLogic(docData, slides, metrics);
        const layout = this.scoreLayout(slides, metrics);
        const imageSemantics = this.scoreImageSemantics(slides, metrics);
        const contentRichness = this.scoreContentRichness(slides, metrics);
        const audienceFit = this.scoreAudienceFit(slides, metrics);
        const consistency = this.scoreConsistency(slides, metrics);

        const overallScore = this.round(
            logic.weightedScore +
                layout.weightedScore +
                imageSemantics.weightedScore +
                contentRichness.weightedScore +
                audienceFit.weightedScore +
                consistency.weightedScore,
            1,
        );

        return {
            version: 'v2',
            generatedAt: new Date().toISOString(),
            title: docData.title,
            outputPath,
            overallScore,
            grade: this.getGrade(overallScore),
            dimensions: {
                logic,
                layout,
                imageSemantics,
                contentRichness,
                audienceFit,
                consistency,
            },
            metrics,
            keyFindings: this.collectKeyFindings(
                metrics,
                logic,
                layout,
                imageSemantics,
                contentRichness,
                audienceFit,
                consistency,
            ),
            nextActions: this.collectNextActions(
                logic,
                layout,
                imageSemantics,
                contentRichness,
                audienceFit,
                consistency,
            ),
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

    private async inspectRenderedDeck(outputPath?: string, deckTitle = ''): Promise<RenderedDeckInspection> {
        if (!outputPath || !fs.existsSync(outputPath)) {
            return this.emptyRenderedDeckInspection();
        }

        try {
            const zip = await JSZip.loadAsync(fs.readFileSync(outputPath));
            const slideEntries = Object.keys(zip.files)
                .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
                .sort((left, right) => this.extractSlideNumber(left) - this.extractSlideNumber(right));

            if (slideEntries.length === 0) {
                return this.emptyRenderedDeckInspection();
            }

            const slideTexts = await Promise.all(
                slideEntries.map(async (name) => this.extractRenderedSlideText(await zip.files[name].async('string'))),
            );
            const dominantLanguage = this.detectDominantLanguage([deckTitle, ...slideTexts].join(' '));

            return {
                renderedSlideCount: slideTexts.length,
                renderedMetaArtifactSlideCount: slideTexts.filter((text) => this.hasRenderedMetaArtifact(text)).length,
                renderedInstructionalTextSlideCount: slideTexts.filter((text) => this.hasRenderedInstructionalArtifact(text))
                    .length,
                renderedMixedLanguageSlideCount: slideTexts.filter((text) =>
                    this.hasRenderedMixedLanguageNarration(text, dominantLanguage),
                ).length,
            };
        } catch (error: any) {
            console.warn('Rendered deck inspection failed:', error?.message || error);
            return this.emptyRenderedDeckInspection();
        }
    }

    private emptyRenderedDeckInspection(): RenderedDeckInspection {
        return {
            renderedSlideCount: 0,
            renderedMetaArtifactSlideCount: 0,
            renderedInstructionalTextSlideCount: 0,
            renderedMixedLanguageSlideCount: 0,
        };
    }

    private extractSlideNumber(name: string): number {
        const match = name.match(/slide(\d+)\.xml$/i);
        return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
    }

    private extractRenderedSlideText(xml: string): string {
        const textRuns = Array.from(xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g), (match) => this.decodeXmlEntities(match[1]));
        return this.cleanText(textRuns.join(' '));
    }

    private decodeXmlEntities(text: string): string {
        return text
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'");
    }

    private detectDominantLanguage(text: string): 'cjk' | 'latin' {
        const cleaned = this.cleanText(text);
        const cjkCount = (cleaned.match(/[\u4e00-\u9fff]/g) || []).length;
        const latinCount = (cleaned.match(/[A-Za-z]/g) || []).length;
        return cjkCount >= Math.max(12, Math.round(latinCount * 0.6)) ? 'cjk' : 'latin';
    }

    private hasRenderedMetaArtifact(text: string): boolean {
        const normalized = this.cleanText(text).toLowerCase();
        if (!normalized) return false;

        const patterns = [
            /\bai-synthesized deck\b/,
            /\bcontent slides\b/,
            /\baudience\s*:/,
            /\bformat\s*:/,
            /\bfocus\s*:/,
            /\bstyle\s*:/,
            /\blength\s*:/,
        ];

        return patterns.some((pattern) => pattern.test(normalized));
    }

    private hasRenderedInstructionalArtifact(text: string): boolean {
        const normalized = this.cleanText(text).toLowerCase();
        if (!normalized) return false;

        const patterns = [
            /\bhelp .* understand\b/,
            /\bpresentation framing\b/,
            /\boverview-driven\b/,
            /\btimeline-focused\b/,
            /\bcomparison-driven\b/,
            /\bprocess-oriented\b/,
            /\bargument-led\b/,
            /\bpractical takeaway\b/,
            /\bfor the audience\b/,
        ];

        return patterns.some((pattern) => pattern.test(normalized));
    }

    private hasRenderedMixedLanguageNarration(text: string, dominantLanguage: 'cjk' | 'latin'): boolean {
        const cleaned = this.cleanText(text);
        if (!cleaned) return false;

        if (dominantLanguage === 'cjk') {
            if (!/[\u4e00-\u9fff]/.test(cleaned)) {
                return false;
            }

            const englishPhrases = cleaned.match(/[A-Za-z][A-Za-z-]*(?:\s+[A-Za-z][A-Za-z-]*)+/g) || [];
            return englishPhrases.some((phrase) => {
                const normalized = phrase.replace(/\s+/g, ' ').trim();
                const wordCount = normalized.split(/\s+/).length;
                const letterCount = (normalized.match(/[A-Za-z]/g) || []).length;
                return wordCount >= 3 || letterCount >= 14;
            });
        }

        if (!/[A-Za-z]/.test(cleaned)) {
            return false;
        }

        return (cleaned.match(/[\u4e00-\u9fff]/g) || []).length >= 6;
    }

    private computeMetrics(slides: SlideContent[], renderedDeck: RenderedDeckInspection): QualityMetrics {
        const slideCount = slides.length;
        const slideWithImageCount = slides.filter((s) => s.images.length > 0).length;
        const slidesWithSummaryCount = slides.filter((s) => this.cleanText(s.summary || '').length > 0).length;
        const slidesWithPromptCount = slides.filter((s) => this.cleanText(s.imagePrompt || '').length > 0).length;
        const imageCoverage = slideCount === 0 ? 0 : slideWithImageCount / slideCount;
        const summaryCoverage = slideCount === 0 ? 0 : slidesWithSummaryCount / slideCount;
        const promptCoverage = slideCount === 0 ? 0 : slidesWithPromptCount / slideCount;
        const totalBullets = slides.reduce((sum, s) => sum + s.bullets.length, 0);
        const avgBulletsPerSlide = slideCount === 0 ? 0 : totalBullets / slideCount;
        const avgTextLengthPerSlide =
            slideCount === 0 ? 0 : slides.reduce((sum, s) => sum + this.totalTextLength(s), 0) / slideCount;
        const allBullets = slides.flatMap((s) => s.bullets);
        const avgBulletLength =
            allBullets.length === 0
                ? 0
                : allBullets.reduce((sum, b) => sum + this.cleanText(b).length, 0) / allBullets.length;

        const overlaySlideCount = slides.filter((s) => s.layout !== 'image_only').length;
        const imageOnlySlideCount = slides.filter((s) => s.layout === 'image_only').length;
        const dominantLayoutRatio = slideCount === 0 ? 0 : Math.max(overlaySlideCount, imageOnlySlideCount) / slideCount;

        return {
            slideCount,
            slideWithImageCount,
            slidesWithSummaryCount,
            slidesWithPromptCount,
            imageCoverage: this.round(imageCoverage, 3),
            summaryCoverage: this.round(summaryCoverage, 3),
            promptCoverage: this.round(promptCoverage, 3),
            avgBulletsPerSlide: this.round(avgBulletsPerSlide, 2),
            avgTextLengthPerSlide: this.round(avgTextLengthPerSlide, 2),
            avgBulletLength: this.round(avgBulletLength, 2),
            levelJumpViolations: this.countLevelJumpViolations(slides),
            duplicateTitleCount: this.countDuplicateTitles(slides),
            genericTitleCount: this.countGenericTitles(slides),
            weakTransitionCount: this.countWeakTransitions(slides),
            actionCueSlideCount: this.countActionCueSlides(slides),
            redundantContentSlideCount: slides.filter((s) => this.countIntraSlideRedundancyItems(s) > 0).length,
            redundantContentItemCount: slides.reduce((sum, s) => sum + this.countIntraSlideRedundancyItems(s), 0),
            sparseContentSlideCount: slides.filter((s) => this.isSparseContentSlide(s)).length,
            severeSparseContentSlideCount: slides.filter((s) => this.isSevereSparseContentSlide(s)).length,
            overlaySlideCount,
            imageOnlySlideCount,
            dominantLayoutRatio: this.round(dominantLayoutRatio, 3),
            overflowRiskSlideCount: this.countOverflowRiskSlides(slides),
            promptAlignmentAvg: this.round(this.averagePromptAlignment(slides), 3),
            fallbackImageCount: slides.filter((s) => s.imageSource === 'ai_fallback' || s.imageSource === 'placeholder')
                .length,
            renderedSlideCount: renderedDeck.renderedSlideCount,
            renderedMetaArtifactSlideCount: renderedDeck.renderedMetaArtifactSlideCount,
            renderedInstructionalTextSlideCount: renderedDeck.renderedInstructionalTextSlideCount,
            renderedMixedLanguageSlideCount: renderedDeck.renderedMixedLanguageSlideCount,
        };
    }

    private scoreLogic(docData: DocumentData, slides: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(`Average bullets: ${metrics.avgBulletsPerSlide}`);
        evidence.push(`Level jumps: ${metrics.levelJumpViolations}`);
        evidence.push(`Weak transitions: ${metrics.weakTransitionCount}`);

        const emptyTitleCount = slides.filter((s) => this.cleanText(s.title).length === 0).length;
        if (emptyTitleCount > 0) {
            score -= emptyTitleCount * 10;
            issues.push(`${emptyTitleCount} slides have empty titles.`);
            recommendations.push('Ensure every slide has a concise title.');
        }
        if (metrics.duplicateTitleCount > 0) {
            score -= metrics.duplicateTitleCount * 4;
            issues.push(`${metrics.duplicateTitleCount} duplicate slide titles detected.`);
            recommendations.push('Rename duplicate titles to highlight unique points.');
        }
        if (metrics.genericTitleCount > 0) {
            score -= metrics.genericTitleCount * 3;
            issues.push(`${metrics.genericTitleCount} titles are generic.`);
            recommendations.push('Use specific, topic-driven title wording.');
        }
        if (metrics.levelJumpViolations > 0) {
            score -= metrics.levelJumpViolations * 6;
            issues.push(`${metrics.levelJumpViolations} hierarchy jumps detected.`);
            recommendations.push('Keep neighboring levels close to maintain narrative continuity.');
        }
        if (metrics.weakTransitionCount > 0) {
            score -= Math.min(20, metrics.weakTransitionCount * 4);
            issues.push(`${metrics.weakTransitionCount} transitions appear disconnected.`);
            recommendations.push('Add transitional context between neighboring slides.');
        } else if (slides.length > 1) {
            evidence.push('Slide-to-slide continuity looks stable.');
        }
        if (metrics.redundantContentItemCount > 0) {
            score -= Math.min(18, metrics.redundantContentSlideCount * 3 + metrics.redundantContentItemCount * 2);
            issues.push(`Repeated content items: ${metrics.redundantContentItemCount}.`);
            recommendations.push('Remove duplicate bullets and repeated summary text.');
        } else {
            evidence.push('No obvious same-slide repetition found.');
        }
        if (metrics.avgBulletsPerSlide >= 2 && metrics.avgBulletsPerSlide <= 5) {
            score += 3;
            evidence.push('Bullet density stays in the recommended range (2-5).');
        } else {
            score -= 4;
            issues.push('Bullet density is outside recommended range (2-5).');
            recommendations.push('Keep most slides around 2-5 bullets.');
        }
        if (this.cleanText(docData.title).length === 0) {
            score -= 8;
            issues.push('Deck title is missing.');
            recommendations.push('Provide a clear deck title to anchor the story.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Content Logic',
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

        evidence.push(`Image coverage: ${(metrics.imageCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Layout dominance: ${(metrics.dominantLayoutRatio * 100).toFixed(1)}%.`);

        const noImageSlides = metrics.slideCount - metrics.slideWithImageCount;
        if (noImageSlides > 0) {
            score -= noImageSlides * 8;
            issues.push(`${noImageSlides} slides have no image.`);
            recommendations.push('Add one image per content slide whenever possible.');
        }
        if (metrics.overflowRiskSlideCount > 0) {
            score -= metrics.overflowRiskSlideCount * 6;
            issues.push(`${metrics.overflowRiskSlideCount} slides have overflow risk.`);
            recommendations.push('Reduce text density or split long slides.');
        } else {
            score += 2;
            evidence.push('No high-risk text overflow slide detected.');
        }
        if (metrics.renderedMetaArtifactSlideCount > 0) {
            score -= Math.min(24, metrics.renderedMetaArtifactSlideCount * 8);
            issues.push(`${metrics.renderedMetaArtifactSlideCount} rendered slides expose planner/debug metadata.`);
            recommendations.push('Remove planner parameters, generator labels, and deck stats from visible slides.');
        } else if (metrics.renderedSlideCount > 0) {
            evidence.push('No planner/debug metadata found in rendered slide text.');
        }
        if (metrics.dominantLayoutRatio > 0.9 && metrics.slideCount >= 6) {
            score -= this.round((metrics.dominantLayoutRatio - 0.9) * 60, 1);
            issues.push('Layout style is too repetitive.');
            recommendations.push('Mix overlay and image-only layouts for better rhythm.');
        } else if (metrics.slideCount >= 6) {
            evidence.push('Layout variation is acceptable.');
        }

        const denseOverlaySlides = slides.filter((s) => s.layout !== 'image_only' && this.totalTextLength(s) > 280).length;
        if (denseOverlaySlides > 0) {
            score -= denseOverlaySlides * 4;
            issues.push(`${denseOverlaySlides} overlay slides are text-heavy.`);
            recommendations.push('Switch dense slides to image-only or shorten text.');
        }
        if (metrics.imageCoverage >= 0.95) {
            score += 4;
            evidence.push('Visual coverage is strong.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Layout Quality',
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

        evidence.push(`Prompt coverage: ${(metrics.promptCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Prompt alignment: ${(metrics.promptAlignmentAvg * 100).toFixed(1)}%.`);
        evidence.push(`Fallback images: ${metrics.fallbackImageCount}.`);

        const missingPromptSlides = metrics.slideCount - metrics.slidesWithPromptCount;
        if (missingPromptSlides > 0) {
            score -= missingPromptSlides * 5;
            issues.push(`${missingPromptSlides} slides are missing image prompts.`);
            recommendations.push('Provide imagePrompt for each slide.');
        }
        const lowAlignSlides = slides.filter((s) => this.promptAlignment(s) < 0.2).length;
        if (lowAlignSlides > 0) {
            score -= lowAlignSlides * 4;
            issues.push(`${lowAlignSlides} slides show weak image-text alignment.`);
            recommendations.push('Use title/summary/bullet keywords in prompts.');
        }
        if (metrics.fallbackImageCount > 0) {
            score -= metrics.fallbackImageCount * 6;
            issues.push(`${metrics.fallbackImageCount} slides used fallback images.`);
            recommendations.push('Improve prompt robustness and retry strategy.');
        }
        if (metrics.promptCoverage < 0.85) {
            score -= this.round((0.85 - metrics.promptCoverage) * 40, 1);
            issues.push('Prompt coverage is below 85%.');
            recommendations.push('Keep prompt coverage above 85%.');
        }
        if (metrics.imageCoverage < 0.9) {
            score -= this.round((0.9 - metrics.imageCoverage) * 40, 1);
            issues.push('Image coverage is below 90%.');
            recommendations.push('Increase image generation coverage.');
        }
        if (metrics.promptAlignmentAvg >= 0.45) {
            score += 4;
            evidence.push('Prompt semantics generally align with content.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Image Semantics',
            score: finalScore,
            weight: this.imageWeight,
            weightedScore: this.round(finalScore * this.imageWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private scoreContentRichness(_: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(`Average text length: ${metrics.avgTextLengthPerSlide}.`);
        evidence.push(`Sparse slides: ${metrics.sparseContentSlideCount} (severe: ${metrics.severeSparseContentSlideCount}).`);
        evidence.push(`Summary coverage: ${(metrics.summaryCoverage * 100).toFixed(1)}%.`);

        if (metrics.severeSparseContentSlideCount > 0) {
            score -= Math.min(45, metrics.severeSparseContentSlideCount * 11);
            issues.push(`${metrics.severeSparseContentSlideCount} slides are severely sparse.`);
            recommendations.push('Expand severe sparse slides to at least 2 bullets + a takeaway.');
        }
        if (metrics.sparseContentSlideCount > 0) {
            score -= Math.min(30, metrics.sparseContentSlideCount * 6);
            issues.push(`${metrics.sparseContentSlideCount} slides are content-sparse.`);
            recommendations.push('Use model expansion for sparse slides while preserving source facts.');
        }
        if (metrics.avgBulletsPerSlide < 1.8) {
            score -= Math.min(18, this.round((1.8 - metrics.avgBulletsPerSlide) * 12, 1));
            issues.push('Average bullets per slide is too low.');
            recommendations.push('Keep average bullets close to 2-5.');
        }
        if (metrics.avgTextLengthPerSlide < 70) {
            score -= Math.min(16, this.round((70 - metrics.avgTextLengthPerSlide) * 0.25, 1));
            issues.push('Average text length is too short.');
            recommendations.push('Add concise context to low-information slides.');
        }
        if (metrics.summaryCoverage < 0.65 && metrics.slideCount >= 5) {
            score -= Math.min(12, this.round((0.65 - metrics.summaryCoverage) * 30, 1));
            issues.push('Summary coverage is low.');
            recommendations.push('Add summary lines to most non-cover slides.');
        } else {
            evidence.push('Summary coverage is acceptable.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Content Richness',
            score: finalScore,
            weight: this.contentRichnessWeight,
            weightedScore: this.round(finalScore * this.contentRichnessWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private scoreAudienceFit(slides: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        const actionCueCoverage = metrics.slideCount === 0 ? 0 : metrics.actionCueSlideCount / metrics.slideCount;
        evidence.push(`Action-cue coverage: ${(actionCueCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Average bullet length: ${metrics.avgBulletLength}.`);
        evidence.push(`Mixed-language rendered slides: ${metrics.renderedMixedLanguageSlideCount}.`);

        if (actionCueCoverage < 0.15 && metrics.slideCount >= 6) {
            score -= Math.min(18, this.round((0.15 - actionCueCoverage) * 80, 1));
            issues.push('Action/takeaway guidance is too sparse.');
            recommendations.push('Add takeaway language on milestone slides.');
        }
        if (metrics.avgBulletLength > 45) {
            score -= 8;
            issues.push('Bullets are too long for quick reading.');
            recommendations.push('Split long bullets into shorter points.');
        } else if (metrics.avgBulletLength > 0 && metrics.avgBulletLength < 8) {
            score -= 6;
            issues.push('Bullets are too short and under-informative.');
            recommendations.push('Extend very short bullets with context.');
        }
        const longBulletSlides = slides.filter((s) => s.bullets.some((b) => this.cleanText(b).length > 90)).length;
        if (longBulletSlides > 0) {
            score -= Math.min(14, longBulletSlides * 3);
            issues.push(`${longBulletSlides} slides contain very long bullet sentences.`);
            recommendations.push('Use speech-friendly short bullets.');
        }
        if (metrics.summaryCoverage < 0.55 && metrics.slideCount >= 5) {
            score -= 8;
            issues.push('Low summary coverage hurts first-glance readability.');
            recommendations.push('Improve summary coverage for audience orientation.');
        }
        if (metrics.renderedInstructionalTextSlideCount > 0) {
            score -= Math.min(18, metrics.renderedInstructionalTextSlideCount * 6);
            issues.push(`${metrics.renderedInstructionalTextSlideCount} rendered slides contain instructional/helper copy.`);
            recommendations.push('Replace helper phrasing with audience-facing content.');
        }
        if (metrics.renderedMixedLanguageSlideCount > 0) {
            score -= Math.min(18, metrics.renderedMixedLanguageSlideCount * 5);
            issues.push(`${metrics.renderedMixedLanguageSlideCount} rendered slides mix languages in visible narration.`);
            recommendations.push('Keep visible narration in one dominant language unless bilingual output is intentional.');
        } else if (metrics.renderedSlideCount > 0) {
            evidence.push('Rendered slide language stays broadly consistent.');
        }
        const lastSlide = slides[slides.length - 1];
        if (lastSlide && !this.hasActionCue(this.collectSlideText(lastSlide))) {
            score -= 5;
            issues.push('Last slide lacks a clear takeaway/next-step cue.');
            recommendations.push('Add a closing takeaway on the final slide.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Audience Fit',
            score: finalScore,
            weight: this.audienceFitWeight,
            weightedScore: this.round(finalScore * this.audienceFitWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private scoreConsistency(_: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;

        evidence.push(`Duplicate titles: ${metrics.duplicateTitleCount}.`);
        evidence.push(`Generic titles: ${metrics.genericTitleCount}.`);
        evidence.push(`Weak transitions: ${metrics.weakTransitionCount}.`);

        if (metrics.duplicateTitleCount > 0) {
            score -= metrics.duplicateTitleCount * 4;
            issues.push('Duplicate titles hurt naming consistency.');
            recommendations.push('Rename duplicate titles with distinct wording.');
        }
        if (metrics.genericTitleCount > 0) {
            score -= metrics.genericTitleCount * 3;
            issues.push('Generic titles weaken style consistency.');
            recommendations.push('Use consistent domain terminology in titles.');
        }
        if (metrics.weakTransitionCount > 0) {
            score -= Math.min(16, metrics.weakTransitionCount * 3);
            issues.push('Some neighboring slides are weakly connected.');
            recommendations.push('Add stronger transitions between neighboring slides.');
        }
        if (metrics.redundantContentItemCount > 0) {
            score -= Math.min(12, metrics.redundantContentItemCount * 1.5);
            issues.push('Repeated content hurts writing consistency.');
            recommendations.push('Trim repeated bullets and repeated summary text.');
        }
        if (metrics.dominantLayoutRatio > 0.92 && metrics.slideCount >= 8) {
            score -= 7;
            issues.push('Layout pattern is overly repetitive for a long deck.');
            recommendations.push('Introduce controlled layout variation.');
        }
        if (Math.abs(metrics.summaryCoverage - metrics.promptCoverage) > 0.4 && metrics.slideCount >= 6) {
            score -= 6;
            issues.push('Planning completeness differs too much between text and image prompts.');
            recommendations.push('Align summary coverage and prompt coverage.');
        }
        if (metrics.fallbackImageCount > 0) {
            score -= Math.min(10, metrics.fallbackImageCount * 2);
            issues.push('Fallback images reduce visual consistency.');
            recommendations.push('Reduce fallback image ratio.');
        }
        if (metrics.renderedMetaArtifactSlideCount > 0) {
            score -= Math.min(10, metrics.renderedMetaArtifactSlideCount * 3);
            issues.push('Visible planner/debug metadata hurts presentation consistency.');
            recommendations.push('Keep only audience-facing labels on slides.');
        }
        if (metrics.renderedMixedLanguageSlideCount > 0) {
            score -= Math.min(10, metrics.renderedMixedLanguageSlideCount * 3);
            issues.push('Mixed-language narration hurts style consistency.');
            recommendations.push('Standardize visible slide language.');
        }
        if (
            metrics.duplicateTitleCount === 0 &&
            metrics.genericTitleCount === 0 &&
            metrics.weakTransitionCount === 0 &&
            metrics.redundantContentItemCount === 0 &&
            metrics.renderedMetaArtifactSlideCount === 0 &&
            metrics.renderedMixedLanguageSlideCount === 0
        ) {
            score += 4;
            evidence.push('Content style consistency is high.');
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Consistency',
            score: finalScore,
            weight: this.consistencyWeight,
            weightedScore: this.round(finalScore * this.consistencyWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private countLevelJumpViolations(slides: SlideContent[]): number {
        let violations = 0;
        for (let i = 1; i < slides.length; i += 1) {
            const prev = slides[i - 1].level || 1;
            const curr = slides[i].level || 1;
            if (Math.abs(curr - prev) > 1) {
                violations += 1;
            }
        }
        return violations;
    }

    private countDuplicateTitles(slides: SlideContent[]): number {
        const countByTitle = new Map<string, number>();
        slides.forEach((slide) => {
            const key = this.cleanText(slide.title).toLowerCase();
            if (!key) return;
            countByTitle.set(key, (countByTitle.get(key) || 0) + 1);
        });
        let duplicates = 0;
        countByTitle.forEach((count) => {
            if (count > 1) {
                duplicates += count - 1;
            }
        });
        return duplicates;
    }

    private countGenericTitles(slides: SlideContent[]): number {
        return slides.filter((slide) => this.isGenericTitle(slide.title)).length;
    }

    private countWeakTransitions(slides: SlideContent[]): number {
        let weakCount = 0;

        for (let i = 1; i < slides.length; i += 1) {
            const prev = slides[i - 1];
            const curr = slides[i];
            if (this.hasTransitionCue(this.collectSlideText(curr))) {
                continue;
            }

            const prevTokens = this.extractKeywords(this.collectSlideText(prev));
            const currTokens = this.extractKeywords(this.collectSlideText(curr));
            const overlap = this.keywordOverlap(prevTokens, currTokens);
            const levelGap = Math.abs((prev.level || 1) - (curr.level || 1));
            const hasBreadcrumb = this.cleanText(curr.breadcrumb || '').length > 0;
            if ((overlap < 0.06 && levelGap > 1 && !hasBreadcrumb) || (overlap < 0.04 && !hasBreadcrumb)) {
                weakCount += 1;
            }
        }

        return weakCount;
    }

    private countActionCueSlides(slides: SlideContent[]): number {
        return slides.filter((slide) => this.hasActionCue(this.collectSlideText(slide))).length;
    }

    private countIntraSlideRedundancyItems(slide: SlideContent): number {
        const bullets = slide.bullets.map((b) => this.cleanText(b)).filter(Boolean);
        let duplicateItems = 0;
        const countByNormalized = new Map<string, number>();

        bullets.forEach((bullet) => {
            const key = this.normalizeForCompare(bullet);
            if (!key) return;
            countByNormalized.set(key, (countByNormalized.get(key) || 0) + 1);
        });

        countByNormalized.forEach((count) => {
            if (count > 1) {
                duplicateItems += count - 1;
            }
        });

        const summary = this.cleanText(slide.summary || '');
        if (summary) {
            const summaryNorm = this.normalizeForCompare(summary);
            if (bullets.some((bullet) => this.normalizeForCompare(bullet) === summaryNorm)) {
                duplicateItems += 1;
            }
        }

        return duplicateItems;
    }

    private countOverflowRiskSlides(slides: SlideContent[]): number {
        return slides.filter((slide) => {
            if (slide.layout === 'image_only') return false;
            const titleLen = this.cleanText(slide.title).length;
            const summaryLen = this.cleanText(slide.summary || '').length;
            const bulletCount = slide.bullets.length;
            const totalLen = this.totalTextLength(slide);
            return titleLen > 44 || summaryLen > 130 || bulletCount > 6 || totalLen > 280;
        }).length;
    }

    private averagePromptAlignment(slides: SlideContent[]): number {
        if (slides.length === 0) return 0;
        const sum = slides.reduce((acc, slide) => acc + this.promptAlignment(slide), 0);
        return sum / slides.length;
    }

    private promptAlignment(slide: SlideContent): number {
        const contentKeywords = this.extractKeywords(this.collectSlideText(slide));
        const promptKeywords = this.extractKeywords([slide.imageIntent || '', slide.imagePrompt || ''].join(' '));
        return this.keywordOverlap(contentKeywords, promptKeywords);
    }

    private totalTextLength(slide: SlideContent): number {
        return (
            this.cleanText(slide.title).length +
            this.cleanText(slide.summary || '').length +
            slide.bullets.reduce((sum, b) => sum + this.cleanText(b).length, 0)
        );
    }

    private isSparseContentSlide(slide: SlideContent): boolean {
        const bulletCount = slide.bullets.filter((b) => this.cleanText(b).length > 0).length;
        const textLen = this.totalTextLength(slide);
        const hasSummary = this.cleanText(slide.summary || '').length > 0;
        return (bulletCount <= 1 && textLen < 90) || (!hasSummary && bulletCount === 0);
    }

    private isSevereSparseContentSlide(slide: SlideContent): boolean {
        const bulletCount = slide.bullets.filter((b) => this.cleanText(b).length > 0).length;
        const textLen = this.totalTextLength(slide);
        const titleLen = this.cleanText(slide.title).length;
        return bulletCount === 0 || (bulletCount <= 1 && textLen < 55 && titleLen < 28);
    }

    private isGenericTitle(title: string): boolean {
        const normalized = this.cleanText(title).toLowerCase();
        if (!normalized) return true;
        const patterns = [
            /^slide\s*\d+$/,
            /^page\s*\d+$/,
            /^section\s*\d+$/,
            /^chapter\s*\d+$/,
            /^part\s*\d+$/,
            /^topic\s*\d+$/,
            /^overview$/,
            /^summary$/,
            /^untitled/,
            /^内容\d*$/,
            /^页面\d*$/,
            /^章节\d*$/,
            /^未命名/,
            /^主题$/,
        ];
        return patterns.some((pattern) => pattern.test(normalized));
    }

    private hasActionCue(text: string): boolean {
        const normalized = this.cleanText(text).toLowerCase();
        if (!normalized) return false;
        const patterns = [
            /\bnext step\b/,
            /\brecommend(?:ation)?\b/,
            /\baction\b/,
            /\btakeaway\b/,
            /\bconclusion\b/,
            /\bimpact\b/,
            /下一步/,
            /建议/,
            /结论/,
            /启示/,
            /行动/,
            /落地/,
            /关键点/,
            /影响/,
            /总结/,
        ];
        return patterns.some((pattern) => pattern.test(normalized));
    }

    private hasTransitionCue(text: string): boolean {
        const normalized = this.cleanText(text).toLowerCase();
        if (!normalized) return false;
        const patterns = [
            /\bnext\b/,
            /\bthen\b/,
            /\bafter\b/,
            /\btherefore\b/,
            /\bmeanwhile\b/,
            /首先/,
            /其次/,
            /然后/,
            /接着/,
            /最后/,
            /因此/,
            /另一方面/,
            /与此同时/,
            /转向/,
            /对比/,
        ];
        return patterns.some((pattern) => pattern.test(normalized));
    }

    private collectSlideText(slide: SlideContent): string {
        return [slide.title, slide.summary || '', slide.breadcrumb || '', slide.bullets.join(' ')].join(' ');
    }

    private extractKeywords(text: string): Set<string> {
        const cleaned = this.cleanText(text).toLowerCase();
        const tokens = cleaned.match(/[\u4e00-\u9fff]{2,}|[a-z0-9]{3,}/g) || [];
        const stopWords = new Set([
            'the',
            'and',
            'for',
            'with',
            'this',
            'that',
            'from',
            'into',
            'about',
            'presentation',
            'slide',
            'content',
            'summary',
            'style',
            'image',
            'topic',
            'section',
            'context',
            '页面',
            '内容',
            '总结',
            '主题',
            '关键点',
        ]);

        const set = new Set<string>();
        tokens.forEach((token) => {
            if (!stopWords.has(token)) {
                set.add(token);
            }
        });
        return set;
    }

    private keywordOverlap(left: Set<string>, right: Set<string>): number {
        if (left.size === 0 || right.size === 0) return 0;
        let overlap = 0;
        left.forEach((token) => {
            if (right.has(token)) {
                overlap += 1;
            }
        });
        return overlap / left.size;
    }

    private normalizeForCompare(text: string): string {
        return this.cleanText(text).toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '');
    }

    private collectKeyFindings(
        metrics: QualityMetrics,
        logic: QualityDimensionScore,
        layout: QualityDimensionScore,
        image: QualityDimensionScore,
        richness: QualityDimensionScore,
        audience: QualityDimensionScore,
        consistency: QualityDimensionScore,
    ): string[] {
        const findings: string[] = [];
        findings.push(`Image coverage ${(metrics.imageCoverage * 100).toFixed(1)}%, prompt coverage ${(metrics.promptCoverage * 100).toFixed(1)}%.`);
        findings.push(`Summary coverage ${(metrics.summaryCoverage * 100).toFixed(1)}%, sparse slides ${metrics.sparseContentSlideCount} (severe ${metrics.severeSparseContentSlideCount}).`);
        findings.push(`Titles: duplicate ${metrics.duplicateTitleCount}, generic ${metrics.genericTitleCount}; weak transitions ${metrics.weakTransitionCount}.`);
        findings.push(`Scores -> Logic ${logic.score}, Layout ${layout.score}, Image ${image.score}, Richness ${richness.score}, Audience ${audience.score}, Consistency ${consistency.score}.`);
        if (metrics.fallbackImageCount > 0) {
            findings.push(`Fallback images used on ${metrics.fallbackImageCount} slides.`);
        }
        if (metrics.overflowRiskSlideCount > 0) {
            findings.push(`${metrics.overflowRiskSlideCount} slides have potential text overflow risk.`);
        }
        if (metrics.renderedMetaArtifactSlideCount > 0) {
            findings.push(`${metrics.renderedMetaArtifactSlideCount} rendered slides leaked planner/debug metadata.`);
        }
        if (metrics.renderedInstructionalTextSlideCount > 0) {
            findings.push(`${metrics.renderedInstructionalTextSlideCount} rendered slides contain helper-style instructional copy.`);
        }
        if (metrics.renderedMixedLanguageSlideCount > 0) {
            findings.push(`${metrics.renderedMixedLanguageSlideCount} rendered slides contain mixed-language narration.`);
        }
        return findings;
    }

    private collectNextActions(...dimensions: QualityDimensionScore[]): string[] {
        const actions = new Set<string>();
        dimensions.forEach((d) => d.recommendations.forEach((rec) => actions.add(rec)));
        return Array.from(actions).slice(0, 10);
    }

    private getGrade(score: number): string {
        if (score >= 90) return 'A';
        if (score >= 80) return 'B';
        if (score >= 70) return 'C';
        if (score >= 60) return 'D';
        return 'E';
    }

    private toMarkdown(report: QualityReport): string {
        const lines = [
            '# PPT Quality Report',
            '',
            `- Title: ${report.title}`,
            `- Generated At: ${report.generatedAt}`,
            `- Overall Score: **${report.overallScore} / 100**`,
            `- Grade: **${report.grade}**`,
        ];
        if (report.outputPath) {
            lines.push(`- Output: ${report.outputPath}`);
        }
        lines.push('', '## Dimension Scores', '', '| Dimension | Score | Weight | Weighted |', '|---|---:|---:|---:|');
        lines.push(`| ${report.dimensions.logic.name} | ${report.dimensions.logic.score} | ${report.dimensions.logic.weight} | ${report.dimensions.logic.weightedScore} |`);
        lines.push(`| ${report.dimensions.layout.name} | ${report.dimensions.layout.score} | ${report.dimensions.layout.weight} | ${report.dimensions.layout.weightedScore} |`);
        lines.push(`| ${report.dimensions.imageSemantics.name} | ${report.dimensions.imageSemantics.score} | ${report.dimensions.imageSemantics.weight} | ${report.dimensions.imageSemantics.weightedScore} |`);
        lines.push(`| ${report.dimensions.contentRichness.name} | ${report.dimensions.contentRichness.score} | ${report.dimensions.contentRichness.weight} | ${report.dimensions.contentRichness.weightedScore} |`);
        lines.push(`| ${report.dimensions.audienceFit.name} | ${report.dimensions.audienceFit.score} | ${report.dimensions.audienceFit.weight} | ${report.dimensions.audienceFit.weightedScore} |`);
        lines.push(`| ${report.dimensions.consistency.name} | ${report.dimensions.consistency.score} | ${report.dimensions.consistency.weight} | ${report.dimensions.consistency.weightedScore} |`);
        lines.push('', '## Metrics', '', '```json', JSON.stringify(report.metrics, null, 2), '```', '', '## Key Findings', '');
        report.keyFindings.forEach((finding) => lines.push(`- ${finding}`));
        lines.push('', '## Suggested Next Actions', '');
        report.nextActions.forEach((action) => lines.push(`- ${action}`));
        lines.push('');
        return lines.join('\n');
    }

    private cleanText(input: any): string {
        if (typeof input !== 'string') return '';
        return input.replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
    }

    private round(value: number, digits = 0): number {
        const factor = 10 ** digits;
        return Math.round(value * factor) / factor;
    }

    private clamp(value: number, min: number, max: number): number {
        return Math.max(min, Math.min(max, value));
    }
}
