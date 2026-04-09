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
    renderedSlideWithTextCount: number;
    renderedSlideWithImageCount: number;
    renderedImageOnlySlideCount: number;
    renderedUniqueImageCount: number;
    renderedMetaArtifactSlideCount: number;
    renderedInstructionalTextSlideCount: number;
    renderedMixedLanguageSlideCount: number;
}

export class EvaluatorService {
    private readonly logicWeight = 0.17;
    private readonly layoutWeight = 0.14;
    private readonly imageWeight = 0.12;
    private readonly contentRichnessWeight = 0.15;
    private readonly audienceFitWeight = 0.14;
    private readonly consistencyWeight = 0.1;
    private readonly sourceUnderstandingWeight = 0.18;

    async evaluate(docData: DocumentData, outputPath?: string): Promise<QualityReport> {
        const slides = docData.slides;
        const renderedDeck = await this.inspectRenderedDeck(outputPath, docData.title);
        const metrics = this.computeMetrics(docData, slides, renderedDeck);

        const logic = this.scoreLogic(docData, slides, metrics);
        const layout = this.scoreLayout(slides, metrics);
        const imageSemantics = this.scoreImageSemantics(slides, metrics);
        const contentRichness = this.scoreContentRichness(slides, metrics);
        const audienceFit = this.scoreAudienceFit(slides, metrics);
        const consistency = this.scoreConsistency(slides, metrics);
        const sourceUnderstanding = this.scoreSourceUnderstanding(docData, slides, metrics);

        const overallScore = this.round(
            logic.weightedScore +
                layout.weightedScore +
                imageSemantics.weightedScore +
                contentRichness.weightedScore +
                audienceFit.weightedScore +
                consistency.weightedScore +
                sourceUnderstanding.weightedScore,
            1,
        );

        return {
            version: 'v3',
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
                sourceUnderstanding,
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
                sourceUnderstanding,
            ),
            nextActions: this.collectNextActions(
                logic,
                layout,
                imageSemantics,
                contentRichness,
                audienceFit,
                consistency,
                sourceUnderstanding,
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
                slideEntries.map(async (name) => {
                    const slideXml = await zip.files[name].async('string');
                    const slideNumber = this.extractSlideNumber(name);
                    const relName = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
                    const relXml = zip.files[relName] ? await zip.files[relName].async('string') : '';
                    const text = this.extractRenderedSlideText(slideXml);
                    const imageTargets = this.extractRenderedImageTargets(relXml, name, slideXml);
                    return {
                        text,
                        imageTargets,
                    };
                }),
            );
            const dominantLanguage = this.detectDominantLanguage([deckTitle, ...slideTexts.map((item) => item.text)].join(' '));
            const uniqueImageTargets = new Set(slideTexts.flatMap((item) => item.imageTargets));

            return {
                renderedSlideCount: slideTexts.length,
                renderedSlideWithTextCount: slideTexts.filter((item) => this.cleanText(item.text).length > 0).length,
                renderedSlideWithImageCount: slideTexts.filter((item) => item.imageTargets.length > 0).length,
                renderedImageOnlySlideCount: slideTexts.filter(
                    (item) => item.imageTargets.length > 0 && this.cleanText(item.text).length < 12,
                ).length,
                renderedUniqueImageCount: uniqueImageTargets.size,
                renderedMetaArtifactSlideCount: slideTexts.filter((item) => this.hasRenderedMetaArtifact(item.text)).length,
                renderedInstructionalTextSlideCount: slideTexts.filter((item) =>
                    this.hasRenderedInstructionalArtifact(item.text),
                ).length,
                renderedMixedLanguageSlideCount: slideTexts.filter((item) =>
                    this.hasRenderedMixedLanguageNarration(item.text, dominantLanguage),
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
            renderedSlideWithTextCount: 0,
            renderedSlideWithImageCount: 0,
            renderedImageOnlySlideCount: 0,
            renderedUniqueImageCount: 0,
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

    private extractRenderedImageTargets(relXml: string, fallbackId: string, slideXml: string): string[] {
        const targets = new Set<string>();
        const relMatches = Array.from(
            relXml.matchAll(/<Relationship\b[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"/gi),
        );

        relMatches.forEach((match) => {
            const type = match[1] || '';
            const target = match[2] || '';
            if (/\/image$/i.test(type) || /(?:^|\/)media\//i.test(target) || /\.\.\/media\//i.test(target)) {
                targets.add(target.replace(/^(\.\.\/)+/g, 'ppt/'));
            }
        });

        if (targets.size === 0 && /<p:pic\b/i.test(slideXml)) {
            targets.add(`inline:${fallbackId}`);
        }

        return Array.from(targets);
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

    private computeMetrics(docData: DocumentData, slides: SlideContent[], renderedDeck: RenderedDeckInspection): QualityMetrics {
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
        const renderedImageCoverage =
            renderedDeck.renderedSlideCount === 0 ? 0 : renderedDeck.renderedSlideWithImageCount / renderedDeck.renderedSlideCount;
        const renderedTextCoverage =
            renderedDeck.renderedSlideCount === 0 ? 0 : renderedDeck.renderedSlideWithTextCount / renderedDeck.renderedSlideCount;
        const visualFirstDeck =
            renderedDeck.renderedSlideCount >= 4 &&
            renderedImageCoverage >= 0.85 &&
            renderedTextCoverage <= 0.2 &&
            renderedDeck.renderedImageOnlySlideCount >= Math.max(3, Math.round(renderedDeck.renderedSlideCount * 0.6));
        const sourceAwareMetrics = this.computeSourceAwareMetrics(docData, slides);

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
            renderedSlideWithImageCount: renderedDeck.renderedSlideWithImageCount,
            renderedImageCoverage: this.round(renderedImageCoverage, 3),
            renderedTextCoverage: this.round(renderedTextCoverage, 3),
            renderedImageOnlySlideCount: renderedDeck.renderedImageOnlySlideCount,
            renderedUniqueImageCount: renderedDeck.renderedUniqueImageCount,
            renderedMetaArtifactSlideCount: renderedDeck.renderedMetaArtifactSlideCount,
            renderedInstructionalTextSlideCount: renderedDeck.renderedInstructionalTextSlideCount,
            renderedMixedLanguageSlideCount: renderedDeck.renderedMixedLanguageSlideCount,
            visualFirstDeck,
            ...sourceAwareMetrics,
        };
    }

    private computeSourceAwareMetrics(
        docData: DocumentData,
        slides: SlideContent[],
    ): Pick<
        QualityMetrics,
        | 'sourceContextAvailable'
        | 'sourceTopicCoverage'
        | 'sourceChapterCoverage'
        | 'sourceSignalCoverage'
        | 'sourceSignalCount'
        | 'sourceRefCoverage'
        | 'thesisAlignment'
        | 'transformedTitleRatio'
        | 'copiedTitleRatio'
        | 'unsupportedTitleRatio'
    > {
        const sourceTopics = this.collectSourceTopics(docData);
        const sourceChapters = this.collectSourceChapterTitles(docData);
        const thesis = this.cleanText(docData.understanding?.thesis || docData.brief?.deckGoal || docData.title);
        const sourceContextAvailable = sourceTopics.length > 0 || sourceChapters.length > 0 || thesis.length > 0;
        const slideTexts = slides.map((slide) => this.collectSlideText(slide));
        const sourceTopicCoverage = sourceTopics.length === 0 ? 1 : this.weightedPhraseCoverage(sourceTopics, slideTexts);
        const sourceChapterCoverage = sourceChapters.length === 0 ? 1 : this.phraseCoverage(sourceChapters, slideTexts);
        const { coverage: sourceSignalCoverage, count: sourceSignalCount } = this.computeSourceSignalCoverage(docData, slides);
        const sourceRefCoverage =
            slides.length === 0 ? 0 : slides.filter((slide) => (slide.sourceRefs || []).length > 0).length / slides.length;
        const thesisAlignment = sourceContextAvailable ? this.computeThesisAlignment(docData, slides) : 0;
        const titleRewriteMetrics = this.computeTitleRewriteMetrics(slides, sourceChapters, sourceTopics);

        return {
            sourceContextAvailable,
            sourceTopicCoverage: this.round(sourceTopicCoverage, 3),
            sourceChapterCoverage: this.round(sourceChapterCoverage, 3),
            sourceSignalCoverage: this.round(sourceSignalCoverage, 3),
            sourceSignalCount,
            sourceRefCoverage: this.round(sourceRefCoverage, 3),
            thesisAlignment: this.round(thesisAlignment, 3),
            transformedTitleRatio: this.round(titleRewriteMetrics.transformedTitleRatio, 3),
            copiedTitleRatio: this.round(titleRewriteMetrics.copiedTitleRatio, 3),
            unsupportedTitleRatio: this.round(titleRewriteMetrics.unsupportedTitleRatio, 3),
        };
    }

    private scoreLogic(docData: DocumentData, slides: SlideContent[], metrics: QualityMetrics): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];
        let score = 100;
        const visualFirstDeck = metrics.visualFirstDeck;
        const textHeuristicLimited = this.shouldDownweightTextHeuristics(metrics);

        evidence.push(`Average bullets: ${metrics.avgBulletsPerSlide}`);
        evidence.push(`Level jumps: ${metrics.levelJumpViolations}`);
        evidence.push(`Weak transitions: ${metrics.weakTransitionCount}`);
        if (visualFirstDeck) {
            evidence.push('Rendered deck is image-first, so text-only logic heuristics are downweighted.');
        }

        const emptyTitleCount = slides.filter((s) => this.cleanText(s.title).length === 0).length;
        if (emptyTitleCount > 0 && !textHeuristicLimited) {
            score -= emptyTitleCount * 10;
            issues.push(`${emptyTitleCount} slides have empty titles.`);
            recommendations.push('Ensure every slide has a concise title.');
        } else if (emptyTitleCount > 0) {
            evidence.push('Empty-title penalty was relaxed because the rendered deck exposes very little text.');
        }
        if (metrics.duplicateTitleCount > 0) {
            score -= metrics.duplicateTitleCount * 4;
            issues.push(`${metrics.duplicateTitleCount} duplicate slide titles detected.`);
            recommendations.push('Rename duplicate titles to highlight unique points.');
        }
        if (metrics.genericTitleCount > 0 && !textHeuristicLimited) {
            score -= metrics.genericTitleCount * 3;
            issues.push(`${metrics.genericTitleCount} titles are generic.`);
            recommendations.push('Use specific, topic-driven title wording.');
        } else if (metrics.genericTitleCount > 0) {
            evidence.push('Generic-title penalty was relaxed because titles are not machine-readable in the rendered deck.');
        }
        if (metrics.levelJumpViolations > 0) {
            score -= metrics.levelJumpViolations * 6;
            issues.push(`${metrics.levelJumpViolations} hierarchy jumps detected.`);
            recommendations.push('Keep neighboring levels close to maintain narrative continuity.');
        }
        if (metrics.weakTransitionCount > 0 && !textHeuristicLimited) {
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
        } else if (textHeuristicLimited) {
            evidence.push('Bullet-density penalty was skipped because this rendered deck is primarily visual.');
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
        const effectiveImageCoverage = this.effectiveImageCoverage(metrics);
        const effectiveNoImageSlides = Math.max(0, Math.round((1 - effectiveImageCoverage) * metrics.slideCount));

        evidence.push(`Planned image coverage: ${(metrics.imageCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Rendered image coverage: ${(metrics.renderedImageCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Layout dominance: ${(metrics.dominantLayoutRatio * 100).toFixed(1)}%.`);
        if (metrics.visualFirstDeck) {
            evidence.push('Rendered output behaves like an image-first deck.');
        }

        if (effectiveNoImageSlides > 0) {
            score -= effectiveNoImageSlides * 8;
            issues.push(`${effectiveNoImageSlides} slides have no image.`);
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
            const repetitionPenalty = this.round((metrics.dominantLayoutRatio - 0.9) * 60, 1) * (metrics.visualFirstDeck ? 0.4 : 1);
            score -= repetitionPenalty;
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
        if (effectiveImageCoverage >= 0.95) {
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
        const effectiveImageCoverage = this.effectiveImageCoverage(metrics);
        const visualRichnessRatio = this.renderedVisualDiversity(metrics);
        const alignmentPenaltyMultiplier = metrics.visualFirstDeck ? 0.35 : 1;

        evidence.push(`Prompt coverage: ${(metrics.promptCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Prompt alignment: ${(metrics.promptAlignmentAvg * 100).toFixed(1)}%.`);
        evidence.push(`Fallback images: ${metrics.fallbackImageCount}.`);
        evidence.push(`Rendered image coverage: ${(metrics.renderedImageCoverage * 100).toFixed(1)}%.`);
        if (metrics.renderedSlideWithImageCount > 0) {
            evidence.push(`Rendered image diversity: ${(visualRichnessRatio * 100).toFixed(1)}%.`);
        }
        if (metrics.visualFirstDeck) {
            evidence.push('Rendered deck is image-first, so prompt-based image checks are downweighted.');
        }

        const missingPromptSlides = metrics.slideCount - metrics.slidesWithPromptCount;
        if (missingPromptSlides > 0) {
            const promptPenalty = missingPromptSlides * 5 * (metrics.visualFirstDeck ? 0.35 : 1);
            score -= promptPenalty;
            issues.push(`${missingPromptSlides} slides are missing image prompts.`);
            recommendations.push('Provide imagePrompt for each slide.');
        }
        const lowAlignSlides = slides.filter((s) => this.promptAlignment(s) < 0.2).length;
        if (lowAlignSlides > 0) {
            score -= lowAlignSlides * 4 * alignmentPenaltyMultiplier;
            issues.push(`${lowAlignSlides} slides show weak image-text alignment.`);
            recommendations.push('Use title/summary/bullet keywords in prompts.');
        }
        if (metrics.fallbackImageCount > 0) {
            score -= metrics.fallbackImageCount * 6;
            issues.push(`${metrics.fallbackImageCount} slides used fallback images.`);
            recommendations.push('Improve prompt robustness and retry strategy.');
        }
        if (metrics.promptCoverage < 0.85) {
            const promptCoveragePenalty = this.round((0.85 - metrics.promptCoverage) * 40, 1) * (metrics.visualFirstDeck ? 0.35 : 1);
            score -= promptCoveragePenalty;
            issues.push('Prompt coverage is below 85%.');
            recommendations.push('Keep prompt coverage above 85%.');
        }
        if (effectiveImageCoverage < 0.9) {
            score -= this.round((0.9 - effectiveImageCoverage) * 40, 1);
            issues.push('Image coverage is below 90%.');
            recommendations.push('Increase image generation coverage.');
        }
        if (metrics.promptAlignmentAvg >= 0.45) {
            score += 4;
            evidence.push('Prompt semantics generally align with content.');
        }
        if (metrics.visualFirstDeck && metrics.renderedImageCoverage >= 0.95) {
            score += 4;
            evidence.push('Rendered image coverage is strong enough to support image-led storytelling.');
        }
        if (visualRichnessRatio >= 0.85 && metrics.renderedSlideWithImageCount >= 4) {
            score += 3;
            evidence.push('Rendered imagery stays visually diverse across slides.');
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
        const visualFirstDeck = metrics.visualFirstDeck;
        const sparsePenaltyMultiplier = visualFirstDeck ? 0.15 : 1;
        const visualRichnessRatio = this.renderedVisualDiversity(metrics);

        evidence.push(`Average text length: ${metrics.avgTextLengthPerSlide}.`);
        evidence.push(`Sparse slides: ${metrics.sparseContentSlideCount} (severe: ${metrics.severeSparseContentSlideCount}).`);
        evidence.push(`Summary coverage: ${(metrics.summaryCoverage * 100).toFixed(1)}%.`);
        if (visualFirstDeck) {
            evidence.push(`Rendered image coverage: ${(metrics.renderedImageCoverage * 100).toFixed(1)}%.`);
            if (metrics.renderedSlideWithImageCount > 0) {
                evidence.push(`Rendered image diversity: ${(visualRichnessRatio * 100).toFixed(1)}%.`);
            }
            evidence.push('Content richness is being evaluated partly through visual storytelling rather than text only.');
        }

        if (metrics.severeSparseContentSlideCount > 0) {
            score -= Math.min(45, metrics.severeSparseContentSlideCount * 11) * sparsePenaltyMultiplier;
            issues.push(`${metrics.severeSparseContentSlideCount} slides are severely sparse.`);
            recommendations.push('Expand severe sparse slides to at least 2 bullets + a takeaway.');
        }
        if (metrics.sparseContentSlideCount > 0) {
            score -= Math.min(30, metrics.sparseContentSlideCount * 6) * sparsePenaltyMultiplier;
            issues.push(`${metrics.sparseContentSlideCount} slides are content-sparse.`);
            recommendations.push('Use model expansion for sparse slides while preserving source facts.');
        }
        if (metrics.avgBulletsPerSlide < 1.8 && !visualFirstDeck) {
            score -= Math.min(18, this.round((1.8 - metrics.avgBulletsPerSlide) * 12, 1));
            issues.push('Average bullets per slide is too low.');
            recommendations.push('Keep average bullets close to 2-5.');
        } else if (visualFirstDeck) {
            evidence.push('Low bullet density is acceptable for an image-first rendered deck.');
        }
        if (metrics.avgTextLengthPerSlide < 70 && !visualFirstDeck) {
            score -= Math.min(16, this.round((70 - metrics.avgTextLengthPerSlide) * 0.25, 1));
            issues.push('Average text length is too short.');
            recommendations.push('Add concise context to low-information slides.');
        } else if (visualFirstDeck) {
            evidence.push('Low visible text is expected here; visual density is used as a compensating signal.');
        }
        if (metrics.summaryCoverage < 0.65 && metrics.slideCount >= 5 && !visualFirstDeck) {
            score -= Math.min(12, this.round((0.65 - metrics.summaryCoverage) * 30, 1));
            issues.push('Summary coverage is low.');
            recommendations.push('Add summary lines to most non-cover slides.');
        } else {
            evidence.push('Summary coverage is acceptable.');
        }
        if (visualFirstDeck && metrics.renderedImageCoverage >= 0.95) {
            score += 6;
            evidence.push('Rendered image coverage strongly supports visual richness.');
        }
        if (visualFirstDeck && visualRichnessRatio >= 0.85 && metrics.renderedSlideWithImageCount >= 4) {
            score += 6;
            evidence.push('Most slides use distinct imagery, which increases content richness.');
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
        const audiencePenaltyMultiplier = metrics.visualFirstDeck ? 0.35 : 1;

        const actionCueCoverage = metrics.slideCount === 0 ? 0 : metrics.actionCueSlideCount / metrics.slideCount;
        evidence.push(`Action-cue coverage: ${(actionCueCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Average bullet length: ${metrics.avgBulletLength}.`);
        evidence.push(`Mixed-language rendered slides: ${metrics.renderedMixedLanguageSlideCount}.`);
        if (metrics.visualFirstDeck) {
            evidence.push('Audience-fit text penalties are downweighted for this image-first deck.');
        }

        if (actionCueCoverage < 0.15 && metrics.slideCount >= 6) {
            score -= Math.min(18, this.round((0.15 - actionCueCoverage) * 80, 1)) * audiencePenaltyMultiplier;
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
            score -= 8 * audiencePenaltyMultiplier;
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
        if (lastSlide && !this.hasActionCue(this.collectSlideText(lastSlide)) && !metrics.visualFirstDeck) {
            score -= 5;
            issues.push('Last slide lacks a clear takeaway/next-step cue.');
            recommendations.push('Add a closing takeaway on the final slide.');
        } else if (lastSlide && metrics.visualFirstDeck) {
            evidence.push('Final-slide text cue check was skipped because the rendered deck is primarily visual.');
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
        const textHeuristicLimited = this.shouldDownweightTextHeuristics(metrics);

        evidence.push(`Duplicate titles: ${metrics.duplicateTitleCount}.`);
        evidence.push(`Generic titles: ${metrics.genericTitleCount}.`);
        evidence.push(`Weak transitions: ${metrics.weakTransitionCount}.`);
        if (metrics.visualFirstDeck) {
            evidence.push('Rendered deck is image-first, so text-only consistency checks are partially relaxed.');
        }

        if (metrics.duplicateTitleCount > 0) {
            score -= metrics.duplicateTitleCount * 4;
            issues.push('Duplicate titles hurt naming consistency.');
            recommendations.push('Rename duplicate titles with distinct wording.');
        }
        if (metrics.genericTitleCount > 0 && !textHeuristicLimited) {
            score -= metrics.genericTitleCount * 3;
            issues.push('Generic titles weaken style consistency.');
            recommendations.push('Use consistent domain terminology in titles.');
        } else if (metrics.genericTitleCount > 0) {
            evidence.push('Generic-title consistency penalty was relaxed because rendered text is not reliably extractable.');
        }
        if (metrics.weakTransitionCount > 0 && !textHeuristicLimited) {
            score -= Math.min(16, metrics.weakTransitionCount * 3);
            issues.push('Some neighboring slides are weakly connected.');
            recommendations.push('Add stronger transitions between neighboring slides.');
        } else if (metrics.weakTransitionCount > 0) {
            evidence.push('Transition-consistency penalty was relaxed for the image-first rendered deck.');
        }
        if (metrics.redundantContentItemCount > 0) {
            score -= Math.min(12, metrics.redundantContentItemCount * 1.5);
            issues.push('Repeated content hurts writing consistency.');
            recommendations.push('Trim repeated bullets and repeated summary text.');
        }
        if (metrics.dominantLayoutRatio > 0.92 && metrics.slideCount >= 8) {
            score -= metrics.visualFirstDeck ? 3 : 7;
            issues.push('Layout pattern is overly repetitive for a long deck.');
            recommendations.push('Introduce controlled layout variation.');
        }
        if (Math.abs(metrics.summaryCoverage - metrics.promptCoverage) > 0.4 && metrics.slideCount >= 6 && !metrics.visualFirstDeck) {
            score -= 6;
            issues.push('Planning completeness differs too much between text and image prompts.');
            recommendations.push('Align summary coverage and prompt coverage.');
        } else if (Math.abs(metrics.summaryCoverage - metrics.promptCoverage) > 0.4 && metrics.visualFirstDeck) {
            evidence.push('Summary/prompt completeness gap was tolerated because the rendered deck is image-first.');
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
            (metrics.genericTitleCount === 0 || textHeuristicLimited) &&
            (metrics.weakTransitionCount === 0 || textHeuristicLimited) &&
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

    private scoreSourceUnderstanding(
        docData: DocumentData,
        slides: SlideContent[],
        metrics: QualityMetrics,
    ): QualityDimensionScore {
        const evidence: string[] = [];
        const issues: string[] = [];
        const recommendations: string[] = [];

        evidence.push(`Source topic coverage: ${(metrics.sourceTopicCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Source chapter coverage: ${(metrics.sourceChapterCoverage * 100).toFixed(1)}%.`);
        evidence.push(`Thesis alignment: ${(metrics.thesisAlignment * 100).toFixed(1)}%.`);
        evidence.push(`Title rewrite ratio: ${(metrics.transformedTitleRatio * 100).toFixed(1)}% transformed, ${(metrics.copiedTitleRatio * 100).toFixed(1)}% copied.`);
        evidence.push(`Source ref coverage: ${(metrics.sourceRefCoverage * 100).toFixed(1)}%.`);

        if (!metrics.sourceContextAvailable) {
            evidence.push('Source understanding context is unavailable, so this dimension uses a neutral fallback.');
            return {
                name: 'Source Understanding',
                score: 80,
                weight: this.sourceUnderstandingWeight,
                weightedScore: this.round(80 * this.sourceUnderstandingWeight, 1),
                evidence,
                issues,
                recommendations,
            };
        }

        let score = 100;
        const understanding = docData.understanding;
        const hasExplicitSignals = metrics.sourceSignalCount > 0;

        if (metrics.sourceTopicCoverage < 0.75) {
            score -= this.round((0.75 - metrics.sourceTopicCoverage) * 60, 1);
            issues.push('Deck coverage of high-importance source topics is incomplete.');
            recommendations.push('Make sure the major source topics appear explicitly in slide titles or key takeaways.');
        } else if (metrics.sourceTopicCoverage >= 0.9) {
            score += 4;
            evidence.push('High-priority source topics are well covered.');
        }

        if (metrics.sourceChapterCoverage < 0.65) {
            score -= this.round((0.65 - metrics.sourceChapterCoverage) * 36, 1);
            issues.push('Some source chapters are weakly represented in the deck.');
            recommendations.push('Preserve more chapter-level coverage when reorganizing the story.');
        } else {
            evidence.push('Chapter-level coverage looks stable.');
        }

        if (metrics.thesisAlignment < 0.18) {
            score -= this.round((0.18 - metrics.thesisAlignment) * 45, 1);
            issues.push('Deck narrative does not strongly reinforce the source thesis.');
            recommendations.push('Strengthen the deck goal, opening message, and closing summary around the source thesis.');
        } else if (metrics.thesisAlignment >= 0.28) {
            score += 4;
            evidence.push('Deck thesis stays aligned with the source thesis.');
        }

        if (hasExplicitSignals && metrics.sourceSignalCoverage < 0.75) {
            score -= this.round((0.75 - metrics.sourceSignalCoverage) * 28, 1);
            issues.push('Some structural source signals were not clearly preserved.');
            recommendations.push('If the source emphasizes timeline, comparison, process, or key data, reflect that structure in slide roles.');
        } else if (hasExplicitSignals) {
            evidence.push('Structural source signals are preserved in the deck.');
        }

        if (metrics.sourceRefCoverage < 0.55 && slides.length >= 4) {
            score -= this.round((0.55 - metrics.sourceRefCoverage) * 20, 1);
            issues.push('Too few slides retain explicit source grounding references.');
            recommendations.push('Keep sourceRefs on most slides so rewritten content stays traceable.');
        } else if (metrics.sourceRefCoverage >= 0.7) {
            evidence.push('Most slides retain source grounding references.');
        }

        if (metrics.unsupportedTitleRatio > 0.22) {
            score -= this.round((metrics.unsupportedTitleRatio - 0.22) * 32, 1);
            issues.push('Too many rewritten slide titles look weakly grounded in the source.');
            recommendations.push('Rewrite titles boldly, but keep them tied to source topics or slide evidence.');
        }

        if (metrics.copiedTitleRatio > 0.78) {
            score -= this.round((metrics.copiedTitleRatio - 0.78) * 26, 1);
            issues.push('Too many slide titles are copied directly from source headings.');
            recommendations.push('Rewrite some source headings into audience-facing presentation titles.');
        }

        if (metrics.transformedTitleRatio >= 0.35 && metrics.transformedTitleRatio <= 0.9 && metrics.unsupportedTitleRatio <= 0.2) {
            score += 5;
            evidence.push('Title rewriting shows a healthy balance between abstraction and grounding.');
        } else if (metrics.transformedTitleRatio < 0.18 && metrics.copiedTitleRatio > 0.55) {
            score -= 4;
            issues.push('Narrative transformation is limited; the deck still reads close to the source outline.');
            recommendations.push('Increase synthesis in titles and summaries instead of mirroring source headings too closely.');
        }

        const hasClosingSynthesis =
            slides.some((slide) => slide.slideRole === 'summary' || slide.slideRole === 'next_step') ||
            (docData.brief?.coreTakeaways || []).length > 0;
        if (hasClosingSynthesis) {
            evidence.push('Deck includes synthesized closing guidance or takeaways.');
        } else if (slides.length >= 6) {
            score -= 4;
            issues.push('Deck lacks an obvious synthesis layer near the end.');
            recommendations.push('Add a summary or next-step slide to show synthesis beyond source extraction.');
        }

        if (understanding?.topics?.length) {
            evidence.push(`Source understanding topics available: ${understanding.topics.length}.`);
        }

        const boundedScore = this.clamp(this.round(score, 1), 0, 100);
        const finalScore = issues.length > 0 ? Math.min(boundedScore, 98) : boundedScore;
        return {
            name: 'Source Understanding',
            score: finalScore,
            weight: this.sourceUnderstandingWeight,
            weightedScore: this.round(finalScore * this.sourceUnderstandingWeight, 1),
            evidence,
            issues,
            recommendations,
        };
    }

    private collectSourceTopics(docData: DocumentData): Array<{ text: string; weight: number }> {
        const topics = (docData.understanding?.topics || [])
            .map((topic) => ({
                text: this.cleanText(topic.title),
                weight: Math.max(1, topic.importance || 1),
            }))
            .filter((topic) => topic.text.length > 0);

        if (topics.length > 0) {
            return topics.slice(0, 10);
        }

        return (docData.brief?.coreTakeaways || [])
            .map((takeaway, index) => ({
                text: this.cleanText(takeaway),
                weight: Math.max(1, 4 - index),
            }))
            .filter((topic) => topic.text.length > 0)
            .slice(0, 8);
    }

    private collectSourceChapterTitles(docData: DocumentData): string[] {
        const titles = [
            ...(docData.understanding?.chapterTitles || []),
            ...(docData.brief?.chapterTitles || []),
        ]
            .map((title) => this.cleanText(title))
            .filter(Boolean);

        return Array.from(new Set(titles.map((title) => this.normalizeForCompare(title))))
            .map((normalized) => titles.find((title) => this.normalizeForCompare(title) === normalized) || '')
            .filter(Boolean)
            .slice(0, 10);
    }

    private weightedPhraseCoverage(items: Array<{ text: string; weight: number }>, slideTexts: string[]): number {
        if (items.length === 0) return 1;
        const totalWeight = items.reduce((sum, item) => sum + Math.max(1, item.weight), 0);
        if (totalWeight <= 0) return 1;

        const coveredWeight = items.reduce((sum, item) => {
            const match = this.maxPhraseMatch(item.text, slideTexts);
            return sum + (match >= 0.18 ? Math.max(1, item.weight) : 0);
        }, 0);

        return coveredWeight / totalWeight;
    }

    private phraseCoverage(phrases: string[], slideTexts: string[]): number {
        if (phrases.length === 0) return 1;
        const covered = phrases.filter((phrase) => this.maxPhraseMatch(phrase, slideTexts) >= 0.18).length;
        return covered / phrases.length;
    }

    private computeSourceSignalCoverage(docData: DocumentData, slides: SlideContent[]): { coverage: number; count: number } {
        const understanding = docData.understanding;
        const checks: boolean[] = [];

        if ((understanding?.timelineSignals || []).length > 0) {
            checks.push(this.deckHasRoleOrPattern(slides, ['timeline'], /\b(18|19|20)\d{2}\b|年|阶段|历程|演进|timeline|history/i));
        }
        if ((understanding?.comparisonSignals || []).length > 0) {
            checks.push(this.deckHasRoleOrPattern(slides, ['comparison'], /对比|比较|区别|差异|优势|劣势|compare|versus|vs\b/i));
        }
        if ((understanding?.processSignals || []).length > 0) {
            checks.push(this.deckHasRoleOrPattern(slides, ['process'], /流程|步骤|方法|实施|推进|落地|process|workflow|step\b/i));
        }
        if ((understanding?.keyNumbers || []).length > 0) {
            checks.push(this.deckHasRoleOrPattern(slides, ['data_highlight'], /\b\d+(?:\.\d+)?%?\b|数据|指标|增长|占比/i));
        }

        if (checks.length === 0) {
            return { coverage: 1, count: 0 };
        }

        const covered = checks.filter(Boolean).length;
        return { coverage: covered / checks.length, count: checks.length };
    }

    private computeThesisAlignment(docData: DocumentData, slides: SlideContent[]): number {
        const thesis = this.cleanText(docData.understanding?.thesis || docData.brief?.deckGoal || docData.title);
        if (!thesis) {
            return 0;
        }

        const summarySlides = slides.filter((slide) => slide.slideRole === 'summary' || slide.slideRole === 'next_step');
        const narrativeWindow = [
            this.cleanText(docData.brief?.deckGoal || ''),
            (docData.brief?.coreTakeaways || []).join(' '),
            ...slides.slice(0, 3).map((slide) => this.collectSlideText(slide)),
            ...summarySlides.slice(0, 2).map((slide) => this.collectSlideText(slide)),
            slides.length > 0 ? this.collectSlideText(slides[slides.length - 1]) : '',
        ].join(' ');

        return this.comparePhraseToText(thesis, narrativeWindow);
    }

    private computeTitleRewriteMetrics(
        slides: SlideContent[],
        sourceChapters: string[],
        sourceTopics: Array<{ text: string; weight: number }>,
    ): { transformedTitleRatio: number; copiedTitleRatio: number; unsupportedTitleRatio: number } {
        const sourcePhrases = Array.from(
            new Set(
                [...sourceChapters, ...sourceTopics.map((topic) => topic.text)]
                    .map((text) => this.cleanText(text))
                    .filter(Boolean),
            ),
        );

        const eligibleSlides = slides.filter((slide) => !this.isUtilityRole(slide.slideRole) && this.cleanText(slide.title).length > 0);
        if (eligibleSlides.length === 0 || sourcePhrases.length === 0) {
            return {
                transformedTitleRatio: 0,
                copiedTitleRatio: 0,
                unsupportedTitleRatio: 0,
            };
        }

        let transformed = 0;
        let copied = 0;
        let unsupported = 0;

        eligibleSlides.forEach((slide) => {
            const title = this.cleanText(slide.title);
            const exactCopied = sourcePhrases.some((phrase) => this.normalizeForCompare(phrase) === this.normalizeForCompare(title));
            const sourceMatch = this.maxPhraseMatch(title, sourcePhrases);
            const titleSupport = this.comparePhraseToText(title, [slide.summary || '', slide.keyMessage || '', slide.bullets.join(' ')].join(' '));
            const grounded = exactCopied || sourceMatch >= 0.18 || ((slide.sourceRefs || []).length > 0 && titleSupport >= 0.1);

            if (exactCopied) {
                copied += 1;
            } else if (grounded) {
                transformed += 1;
            } else {
                unsupported += 1;
            }
        });

        return {
            transformedTitleRatio: transformed / eligibleSlides.length,
            copiedTitleRatio: copied / eligibleSlides.length,
            unsupportedTitleRatio: unsupported / eligibleSlides.length,
        };
    }

    private deckHasRoleOrPattern(slides: SlideContent[], roles: string[], pattern: RegExp): boolean {
        return slides.some((slide) => {
            const roleMatch = slide.slideRole ? roles.includes(slide.slideRole) : false;
            const textMatch = pattern.test(this.collectSlideText(slide));
            return roleMatch || textMatch;
        });
    }

    private comparePhraseToText(phrase: string, text: string): number {
        const left = this.cleanText(phrase);
        const right = this.cleanText(text);
        if (!left || !right) return 0;

        const leftNorm = this.normalizeForCompare(left);
        const rightNorm = this.normalizeForCompare(right);
        if (!leftNorm || !rightNorm) return 0;

        if (rightNorm.includes(leftNorm) || leftNorm.includes(rightNorm)) {
            return 1;
        }

        return this.keywordOverlap(this.extractKeywords(left), this.extractKeywords(right));
    }

    private maxPhraseMatch(phrase: string, candidates: string[]): number {
        let max = 0;
        candidates.forEach((candidate) => {
            max = Math.max(max, this.comparePhraseToText(phrase, candidate));
        });
        return max;
    }

    private isUtilityRole(role?: string): boolean {
        return role === 'agenda' || role === 'summary' || role === 'next_step' || role === 'section_divider';
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
        sourceUnderstanding: QualityDimensionScore,
    ): string[] {
        const findings: string[] = [];
        findings.push(
            `Image coverage planned/rendered ${(metrics.imageCoverage * 100).toFixed(1)}% / ${(metrics.renderedImageCoverage * 100).toFixed(1)}%, prompt coverage ${(metrics.promptCoverage * 100).toFixed(1)}%.`,
        );
        findings.push(`Summary coverage ${(metrics.summaryCoverage * 100).toFixed(1)}%, sparse slides ${metrics.sparseContentSlideCount} (severe ${metrics.severeSparseContentSlideCount}).`);
        findings.push(`Titles: duplicate ${metrics.duplicateTitleCount}, generic ${metrics.genericTitleCount}; weak transitions ${metrics.weakTransitionCount}.`);
        findings.push(
            `Source coverage topics/chapters ${(metrics.sourceTopicCoverage * 100).toFixed(1)}% / ${(metrics.sourceChapterCoverage * 100).toFixed(1)}%, thesis alignment ${(metrics.thesisAlignment * 100).toFixed(1)}%.`,
        );
        findings.push(
            `Scores -> Logic ${logic.score}, Layout ${layout.score}, Image ${image.score}, Richness ${richness.score}, Audience ${audience.score}, Consistency ${consistency.score}, Source ${sourceUnderstanding.score}.`,
        );
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
        if (metrics.visualFirstDeck) {
            findings.push(
                `Rendered deck is image-first: ${metrics.renderedSlideWithImageCount}/${metrics.renderedSlideCount} slides contain images and ${metrics.renderedUniqueImageCount} unique rendered image assets were detected.`,
            );
        }
        return findings;
    }

    private effectiveImageCoverage(metrics: QualityMetrics): number {
        return Math.max(metrics.imageCoverage, metrics.renderedImageCoverage);
    }

    private renderedVisualDiversity(metrics: QualityMetrics): number {
        if (metrics.renderedSlideWithImageCount === 0) {
            return 0;
        }
        return metrics.renderedUniqueImageCount / metrics.renderedSlideWithImageCount;
    }

    private shouldDownweightTextHeuristics(metrics: QualityMetrics): boolean {
        return metrics.visualFirstDeck && metrics.renderedTextCoverage <= 0.1 && metrics.avgTextLengthPerSlide < 20;
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
        lines.push(`| ${report.dimensions.sourceUnderstanding.name} | ${report.dimensions.sourceUnderstanding.score} | ${report.dimensions.sourceUnderstanding.weight} | ${report.dimensions.sourceUnderstanding.weightedScore} |`);
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
