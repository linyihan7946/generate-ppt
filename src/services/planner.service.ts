import axios from 'axios';
import fs from 'fs';
import {
    DeckAudience,
    DeckBrief,
    DeckFocus,
    DeckFormat,
    DeckLength,
    DeckStyle,
    DocumentData,
    PlannerMode,
    PlannerOptions,
    SlideContent,
    SlideLayoutType,
    SlideRole,
} from '../types';
import { UnderstandingService } from './understanding.service';

interface PlannedSlide {
    title: string;
    summary: string;
    bullets: string[];
    layout: SlideLayoutType;
    imageIntent: string;
    imagePrompt: string;
    slideRole: SlideRole;
    keyMessage: string;
    speakerNotes: string[];
    sourceRefs: number[];
}

interface PlannedDocument {
    title?: string;
    brief?: Partial<DeckBrief>;
    slides: PlannedSlide[];
}

interface SparseExpansionSlide {
    index: number;
    keyMessage: string;
    summary: string;
    bullets: string[];
}

interface PlanningPreferences {
    audience: DeckAudience;
    focus: DeckFocus;
    style: DeckStyle;
    deckFormat: DeckFormat;
    length: DeckLength;
}

export class PlannerService {
    private baseUrl: string;
    private authToken: string;
    private fallbackAuthToken: string;
    private model: string;
    private enabled: boolean;
    private allowGuestLogin: boolean;
    private defaultMode: PlannerMode;
    private sparseExpansionEnabled: boolean;
    private workerUrl: string;
    private workerApiKey: string;
    private readonly understandingService: UnderstandingService;

    constructor() {
        const externalEnv = this.loadExternalPlannerEnv();
        this.baseUrl = process.env.PLANNER_API_BASE_URL || process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn:3001';
        this.authToken = process.env.PLANNER_AUTH_TOKEN || process.env.LLM_AUTH_TOKEN || '';
        this.fallbackAuthToken = process.env.IMAGE_API_KEY || '';
        this.model = process.env.PLANNER_MODEL || 'gemini-3.1-pro-preview';
        this.enabled = process.env.ENABLE_PLANNER !== 'false';
        this.allowGuestLogin = process.env.PLANNER_USE_GUEST_LOGIN === 'true';
        this.defaultMode = process.env.PLANNER_CONTENT_MODE === 'creative' ? 'creative' : 'strict';
        this.sparseExpansionEnabled = process.env.PLANNER_EXPAND_SPARSE_CONTENT !== 'false';
        this.workerUrl = process.env.CLOUDFLARE_WORKER_URL || externalEnv.CLOUDFLARE_WORKER_URL || '';
        this.workerApiKey =
            process.env.LLM_API_KEY || process.env.GOOGLE_API_KEY || externalEnv.LLM_API_KEY || externalEnv.GOOGLE_API_KEY || '';
        this.understandingService = new UnderstandingService();
    }

    async planDocument(docData: DocumentData, options: PlannerOptions = {}): Promise<DocumentData> {
        const mode = this.resolvePlannerMode(options.mode);
        const preferences = this.resolvePlanningPreferences(options);
        const heuristicPlan = this.buildHeuristicPlan(docData, mode, preferences);
        let plannedDoc = heuristicPlan;

        if (this.enabled) {
            const llmPlan = await this.generatePlanWithGemini(docData, heuristicPlan.brief as DeckBrief, mode, preferences);
            if (llmPlan && llmPlan.slides.length > 0) {
                plannedDoc = this.mergePlannedDocument(heuristicPlan, llmPlan, mode, preferences);
            }
        }

        plannedDoc = await this.expandSparseSlidesIfNeeded(plannedDoc, mode, preferences);
        plannedDoc.slides = this.strengthenNarrativeContinuity(plannedDoc.slides, plannedDoc.title || docData.title);
        plannedDoc.slides = this.ensureUniqueTitles(plannedDoc.slides);
        return plannedDoc;
    }

    private async generatePlanWithGemini(
        docData: DocumentData,
        brief: DeckBrief,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): Promise<PlannedDocument | null> {
        if (this.canUseWorkerProxy()) {
            const viaWorker = await this.generatePlanWithWorkerProxy(docData, brief, mode, preferences);
            if (viaWorker) {
                return viaWorker;
            }
        }

        const token = await this.resolveAuthToken();
        if (!token) {
            console.warn('Planner skipped: missing PLANNER_AUTH_TOKEN / LLM_AUTH_TOKEN / IMAGE_API_KEY.');
            return null;
        }

        const payload = {
            prompt: this.buildPlannerPrompt(docData, brief, mode, preferences),
            temperature: mode === 'creative' ? 0.35 : 0.2,
            model: this.model,
            stream: false,
        };

        try {
            const response = await axios.post(`${this.baseUrl}/api/llm/direct`, payload, {
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: `Bearer ${token}`,
                },
                timeout: 120000,
                proxy: false,
                validateStatus: () => true,
            });

            if (response.status !== 200 || response.data?.success === false) {
                console.warn(`Planner API failed: status=${response.status}, message=${response.data?.message || 'unknown'}`);
                return null;
            }

            const rawText = this.extractModelText(response.data);
            if (!rawText) {
                console.warn('Planner API returned empty content.');
                return null;
            }

            const parsed = this.parsePlannedDocument(rawText);
            if (!parsed || parsed.slides.length === 0) {
                console.warn('Planner returned invalid JSON structure.');
                return null;
            }

            return parsed;
        } catch (error: any) {
            console.warn('Planner request failed:', error?.message || error);
            return null;
        }
    }

    private async generatePlanWithWorkerProxy(
        docData: DocumentData,
        brief: DeckBrief,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): Promise<PlannedDocument | null> {
        try {
            const text = await this.callGeminiViaWorker(
                this.buildPlannerPrompt(docData, brief, mode, preferences),
                mode === 'creative' ? 0.35 : 0.2,
            );
            if (!text) {
                return null;
            }

            const parsed = this.parsePlannedDocument(text);
            if (!parsed || parsed.slides.length === 0) {
                console.warn('Planner worker proxy returned invalid JSON structure.');
                return null;
            }

            return parsed;
        } catch (error: any) {
            console.warn('Planner worker proxy request failed:', error?.message || error);
            return null;
        }
    }

    private buildPlannerPrompt(
        docData: DocumentData,
        brief: DeckBrief,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): string {
        const compactSlides = docData.slides.map((slide, idx) => ({
            sourceIndex: idx + 1,
            title: slide.title,
            summary: slide.summary || '',
            bullets: slide.bullets,
            level: slide.level || 1,
            breadcrumb: slide.breadcrumb || '',
        }));

        const rules = [
            'You are a professional presentation strategist.',
            'Task: create a source-grounded PPT plan from uploaded material.',
            'You may merge, split, reorder, or add agenda / summary / next-step slides if it improves presentation quality.',
            'Do not introduce unsupported facts, dates, statistics, or named entities.',
            'Every factual slide should stay grounded in the provided source slides.',
            `Deck format: ${preferences.deckFormat}. Audience: ${preferences.audience}. Focus: ${preferences.focus}. Style: ${preferences.style}. Length: ${preferences.length}.`,
            mode === 'creative'
                ? 'Mode creative: polish phrasing, sharpen takeaways, and make slides presentation-ready without changing facts.'
                : 'Mode strict: stay very close to the source meaning, with only minimal restructuring.',
            'Use slideRole from: content, agenda, section_divider, key_insight, timeline, comparison, process, data_highlight, summary, next_step.',
            'For presenter decks, keep visible text concise and put extra explanation into speakerNotes.',
            'For detailed decks, visible bullets can be slightly fuller but should still be readable.',
            'imagePrompt must be concise, visual, and avoid text/logo/watermark.',
            'Return valid JSON only.',
            'Output schema:',
            '{"title":"string","brief":{"deckGoal":"string","audience":"general","focus":"overview","style":"professional","deckFormat":"presenter","desiredLength":"default","chapterTitles":["..."],"coreTakeaways":["..."]},"slides":[{"title":"string","slideRole":"content","keyMessage":"string","summary":"string","bullets":["..."],"speakerNotes":["..."],"sourceRefs":[1,2],"layout":"image_overlay","imageIntent":"string","imagePrompt":"string"}]}',
        ].join('\n');

        return `${rules}\n\nHeuristic brief:\n${JSON.stringify(brief, null, 2)}\n\nSource material:\n${JSON.stringify(
            { title: docData.title, slides: compactSlides },
            null,
            2,
        )}`;
    }

    private extractModelText(payload: any): string {
        if (!payload) return '';
        if (typeof payload === 'string') return payload;
        if (typeof payload.data === 'string') return payload.data;
        if (payload.data?.reply) return String(payload.data.reply);
        if (payload.data?.text) return String(payload.data.text);
        if (payload.data?.content) return String(payload.data.content);
        if (payload.reply) return String(payload.reply);
        if (payload.message && typeof payload.message === 'string' && payload.message.trim().startsWith('{')) {
            return payload.message;
        }

        const choiceText = payload.choices?.[0]?.message?.content || payload.choices?.[0]?.text;
        if (typeof choiceText === 'string') return choiceText;

        const dataChoiceText = payload.data?.choices?.[0]?.message?.content || payload.data?.choices?.[0]?.text;
        if (typeof dataChoiceText === 'string') return dataChoiceText;

        return '';
    }

    private parsePlannedDocument(raw: string): PlannedDocument | null {
        const jsonText = this.extractJsonBlock(raw);
        if (!jsonText) return null;

        try {
            const parsed = JSON.parse(jsonText);
            if (!parsed || !Array.isArray(parsed.slides)) return null;

            const slides = parsed.slides
                .map((slide: any) => this.normalizePlannedSlide(slide))
                .filter((slide: PlannedSlide | null): slide is PlannedSlide => Boolean(slide));

            if (slides.length === 0) return null;

            return {
                title: typeof parsed.title === 'string' ? this.cleanText(parsed.title, 80) : undefined,
                brief: this.normalizeBrief(parsed.brief),
                slides,
            };
        } catch {
            return null;
        }
    }

    private normalizeBrief(input: any): Partial<DeckBrief> | undefined {
        if (!input || typeof input !== 'object') {
            return undefined;
        }

        return {
            deckGoal: this.cleanText(input.deckGoal, 180),
            audience: this.normalizeAudience(input.audience),
            focus: this.normalizeFocus(input.focus),
            style: this.normalizeStyle(input.style),
            deckFormat: this.normalizeDeckFormat(input.deckFormat),
            desiredLength: this.normalizeLength(input.desiredLength),
            chapterTitles: this.normalizeStringList(input.chapterTitles, 8, 60),
            coreTakeaways: this.normalizeStringList(input.coreTakeaways, 5, 120),
        };
    }

    private normalizePlannedSlide(input: any): PlannedSlide | null {
        const title = this.cleanText(input?.title, 80);
        const bullets = this.normalizeBullets(Array.isArray(input?.bullets) ? input.bullets : [], 6);
        const keyMessage = this.cleanText(input?.keyMessage, 140) || this.cleanText(input?.summary, 140) || title;
        const summary = this.cleanText(input?.summary, 140);
        const speakerNotes = this.normalizeStringList(input?.speakerNotes, 6, 180);
        const sourceRefs = this.normalizeSourceRefs(input?.sourceRefs);
        const slideRole = this.normalizeSlideRole(input?.slideRole);
        const layout = this.normalizeLayout(input?.layout, slideRole);
        const imageIntent = this.cleanText(input?.imageIntent, 180) || title || keyMessage;
        const imagePrompt = this.cleanText(input?.imagePrompt, 220) || this.cleanText(`${title}. ${keyMessage}`, 220);

        if (!title && bullets.length === 0 && !summary) {
            return null;
        }

        return {
            title: title || keyMessage || 'Slide',
            summary,
            bullets,
            layout,
            imageIntent,
            imagePrompt,
            slideRole,
            keyMessage: keyMessage || title || 'Key message',
            speakerNotes,
            sourceRefs,
        };
    }

    private extractJsonBlock(raw: string): string {
        const fenced = raw.match(/```json\s*([\s\S]*?)\s*```/i) || raw.match(/```\s*([\s\S]*?)\s*```/i);
        if (fenced?.[1]) {
            return fenced[1].trim();
        }

        const firstBrace = raw.indexOf('{');
        const lastBrace = raw.lastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace) {
            return raw.slice(firstBrace, lastBrace + 1);
        }

        return '';
    }

    private buildHeuristicPlan(
        docData: DocumentData,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): DocumentData {
        const normalizedSourceSlides = docData.slides.map((slide, idx) => {
            const title = this.cleanText(slide.title, 80) || `Slide ${idx + 1}`;
            const bullets = this.normalizeBullets(slide.bullets, preferences.deckFormat === 'presenter' ? 4 : 6);
            const keyMessage = this.cleanText(slide.summary || bullets[0] || title, 140);
            const summary = this.selectSummary(slide.summary || keyMessage || title, bullets, title);
            const slideRole = this.inferSlideRole(
                { ...slide, title, bullets, summary, sourceIndex: idx + 1 },
                idx,
                docData.slides,
                preferences,
            );

            return this.enrichSlideForRole(
                {
                    ...slide,
                    title,
                    bullets,
                    summary,
                    keyMessage,
                    sourceIndex: idx + 1,
                    sourceRefs: [idx + 1],
                    slideRole,
                },
                docData.title,
                mode,
                preferences,
            );
        });

        const understanding = this.understandingService.analyze({
            title: docData.title,
            slides: normalizedSourceSlides,
        });
        const brief = this.buildDeckBrief(docData.title, understanding, preferences);

        const slides: SlideContent[] = [];
        if (this.shouldAddAgenda(normalizedSourceSlides, preferences)) {
            slides.push(this.buildAgendaSlide(brief, normalizedSourceSlides, docData.title, mode, preferences));
        }

        normalizedSourceSlides.forEach((slide) => slides.push(slide));
        const closedSlides = this.ensureClosingSlides(slides, brief, docData.title, mode, preferences);

        return {
            title: this.cleanText(docData.title, 80) || 'Presentation',
            brief,
            understanding,
            slides: this.ensureUniqueTitles(closedSlides),
        };
    }

    private buildDeckBrief(
        title: string,
        understanding: DocumentData['understanding'],
        preferences: PlanningPreferences,
    ): DeckBrief {
        const chapterTitles = (understanding?.chapterTitles || []).filter(Boolean).slice(0, 6);
        const coreTakeaways = (understanding?.topics || [])
            .map((topic) => topic.title)
            .filter(Boolean)
            .slice(0, 3);

        return {
            deckGoal: this.buildDeckGoal(title, preferences),
            audience: preferences.audience,
            focus: preferences.focus,
            style: preferences.style,
            deckFormat: preferences.deckFormat,
            desiredLength: preferences.length,
            chapterTitles,
            coreTakeaways: coreTakeaways.length > 0 ? coreTakeaways : [this.cleanText(title, 80) || 'Core takeaway'],
        };
    }

    private buildDeckGoal(title: string, preferences: PlanningPreferences): string {
        const deckTitle = this.cleanText(title, 80) || 'the source material';
        const audienceText = this.mapAudienceLabel(preferences.audience);
        const focusText = this.mapFocusLabel(preferences.focus);
        return `Help ${audienceText} understand ${deckTitle} with ${focusText} presentation framing.`;
    }

    private inferSlideRole(
        slide: SlideContent,
        index: number,
        allSlides: SlideContent[],
        preferences: PlanningPreferences,
    ): SlideRole {
        if (preferences.focus === 'timeline' || this.looksLikeTimeline(slide)) {
            return 'timeline';
        }
        if (preferences.focus === 'comparison' || this.looksLikeComparison(slide)) {
            return 'comparison';
        }
        if (preferences.focus === 'process' || this.looksLikeProcess(slide)) {
            return 'process';
        }
        if (this.looksLikeDataHighlight(slide)) {
            return 'data_highlight';
        }
        if (this.isSectionDividerCandidate(slide, index, allSlides)) {
            return 'section_divider';
        }
        if (preferences.deckFormat === 'presenter' && slide.bullets.length <= 3) {
            return 'key_insight';
        }
        return 'content';
    }

    private enrichSlideForRole(
        slide: SlideContent,
        deckTitle: string,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): SlideContent {
        const slideRole = slide.slideRole || 'content';
        const maxBullets = preferences.deckFormat === 'presenter' ? 4 : 6;
        const bullets = this.normalizeBullets(slide.bullets, maxBullets);
        const title = this.cleanText(slide.title, 80) || 'Slide';
        const keyMessage = this.cleanText(slide.keyMessage || slide.summary || bullets[0] || title, 140) || title;
        const summary = this.selectSummary(slide.summary || keyMessage, bullets, title);
        const speakerNotes = this.buildSpeakerNotes(title, keyMessage, bullets, slideRole, preferences);
        const layout = slide.layout || this.heuristicLayout(title, bullets, slideRole, preferences);
        const keywordSeed = [title, keyMessage, ...bullets.slice(0, 2).map((bullet) => this.cleanText(bullet, 80))]
            .filter(Boolean)
            .join(' | ');
        const imageIntent = this.cleanText(slide.imageIntent || keywordSeed, 180) || title;
        const imagePrompt =
            this.cleanText(
                slide.imagePrompt || this.buildRoleAwareImagePrompt(deckTitle, title, keyMessage, bullets, slideRole, mode, preferences),
                220,
            ) || this.cleanText(`${keywordSeed}. visual`, 220);
        const breadcrumb = this.cleanText(slide.breadcrumb || this.buildDefaultBreadcrumb(deckTitle, title, slideRole), 80);

        return {
            ...slide,
            title,
            bullets,
            summary,
            keyMessage,
            speakerNotes,
            layout,
            imageIntent,
            imagePrompt,
            breadcrumb,
            slideRole,
            sourceRefs: this.normalizeSourceRefs(slide.sourceRefs || [slide.sourceIndex || 0]),
        };
    }

    private shouldAddAgenda(slides: SlideContent[], preferences: PlanningPreferences): boolean {
        if (slides.length < 4) return false;
        return preferences.deckFormat === 'presenter' || preferences.length !== 'short';
    }

    private buildAgendaSlide(
        brief: DeckBrief,
        slides: SlideContent[],
        deckTitle: string,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): SlideContent {
        const isCjk = this.isMostlyCjk(deckTitle);
        const chapterTitles =
            brief.chapterTitles.length > 0
                ? brief.chapterTitles
                : slides.map((slide) => slide.title).filter(Boolean).slice(0, 5);

        return this.enrichSlideForRole(
            {
                title: isCjk ? '内容导航' : 'Agenda',
                bullets: chapterTitles,
                images: [],
                summary: brief.deckGoal,
                keyMessage: brief.deckGoal,
                slideRole: 'agenda',
                sourceRefs: slides.map((slide) => slide.sourceIndex || 0).filter((ref) => ref > 0).slice(0, 6),
            },
            deckTitle,
            mode,
            preferences,
        );
    }

    private ensureClosingSlides(
        slides: SlideContent[],
        brief: DeckBrief,
        deckTitle: string,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): SlideContent[] {
        const result = [...slides];
        const hasSummary = result.some((slide) => slide.slideRole === 'summary');
        if (!hasSummary && result.length >= 3) {
            result.push(this.buildSummarySlide(brief, deckTitle, mode, preferences));
        }

        const hasNextStep = result.some((slide) => slide.slideRole === 'next_step');
        if (!hasNextStep && preferences.deckFormat === 'presenter' && result.length >= 4) {
            result.push(this.buildNextStepSlide(brief, deckTitle, mode, preferences));
        }

        return result;
    }

    private buildSummarySlide(
        brief: DeckBrief,
        deckTitle: string,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): SlideContent {
        const isCjk = this.isMostlyCjk(deckTitle);
        return this.enrichSlideForRole(
            {
                title: isCjk ? '核心总结' : 'Key Takeaways',
                bullets: brief.coreTakeaways.slice(0, preferences.deckFormat === 'presenter' ? 3 : 4),
                images: [],
                summary: brief.deckGoal,
                keyMessage: brief.deckGoal,
                slideRole: 'summary',
                sourceRefs: [],
            },
            deckTitle,
            mode,
            preferences,
        );
    }

    private buildNextStepSlide(
        brief: DeckBrief,
        deckTitle: string,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): SlideContent {
        const isCjk = this.isMostlyCjk(deckTitle);
        return this.enrichSlideForRole(
            {
                title: isCjk ? '下一步' : 'Next Steps',
                bullets: this.buildNextStepBullets(brief, isCjk),
                images: [],
                summary: isCjk ? '把关键结论转化为行动。' : 'Turn the main insights into action.',
                keyMessage: isCjk ? '把关键结论转化为行动。' : 'Turn the main insights into action.',
                slideRole: 'next_step',
                sourceRefs: [],
            },
            deckTitle,
            mode,
            preferences,
        );
    }

    private buildNextStepBullets(brief: DeckBrief, isCjk: boolean): string[] {
        if (isCjk) {
            return [
                '回顾本次内容的核心结论',
                `围绕${this.mapFocusLabelZh(brief.focus)}明确后续讨论重点`,
                '根据受众需要补充案例、数据或实施方案',
            ];
        }

        return [
            'Revisit the core takeaway from this deck',
            `Prioritize the next discussion around ${this.mapFocusLabel(brief.focus)}`,
            'Add supporting examples, data, or execution details for the audience',
        ];
    }

    private mergePlannedDocument(
        heuristic: DocumentData,
        llmPlan: PlannedDocument,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): DocumentData {
        const mergedBrief: DeckBrief = {
            ...(heuristic.brief as DeckBrief),
            ...(llmPlan.brief || {}),
            chapterTitles: (llmPlan.brief?.chapterTitles || (heuristic.brief as DeckBrief).chapterTitles || []).filter(Boolean),
            coreTakeaways: (llmPlan.brief?.coreTakeaways || (heuristic.brief as DeckBrief).coreTakeaways || []).filter(Boolean),
        };

        const slides = llmPlan.slides.map((planned, index) => {
            const fallback = this.findFallbackSlide(heuristic.slides, planned.sourceRefs, index);
            return this.enrichSlideForRole(
                {
                    ...fallback,
                    title: planned.title || fallback?.title || `Slide ${index + 1}`,
                    bullets: planned.bullets.length > 0 ? planned.bullets : fallback?.bullets || [],
                    summary: planned.summary || fallback?.summary || '',
                    keyMessage: planned.keyMessage || fallback?.keyMessage || planned.title,
                    slideRole: planned.slideRole || fallback?.slideRole || 'content',
                    layout: planned.layout || fallback?.layout,
                    imageIntent: planned.imageIntent || fallback?.imageIntent,
                    imagePrompt: planned.imagePrompt || fallback?.imagePrompt,
                    sourceRefs: planned.sourceRefs.length > 0 ? planned.sourceRefs : fallback?.sourceRefs || [],
                    speakerNotes: planned.speakerNotes.length > 0 ? planned.speakerNotes : fallback?.speakerNotes || [],
                    images: fallback?.images || [],
                    level: fallback?.level,
                    breadcrumb: fallback?.breadcrumb,
                    sourceIndex: fallback?.sourceIndex,
                },
                heuristic.title,
                mode,
                preferences,
            );
        });

        const closedSlides = this.ensureClosingSlides(slides, mergedBrief, heuristic.title, mode, preferences);

        return {
            title: llmPlan.title || heuristic.title,
            brief: mergedBrief,
            understanding: heuristic.understanding,
            slides: this.ensureUniqueTitles(closedSlides),
        };
    }

    private findFallbackSlide(slides: SlideContent[], sourceRefs: number[], index: number): SlideContent | undefined {
        if (sourceRefs.length > 0) {
            const mapped = slides.find((slide) =>
                (slide.sourceRefs || [slide.sourceIndex || 0]).some((ref) => sourceRefs.includes(ref)),
            );
            if (mapped) {
                return mapped;
            }
        }

        return slides[index];
    }

    private async expandSparseSlidesIfNeeded(
        docData: DocumentData,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): Promise<DocumentData> {
        if (!this.sparseExpansionEnabled) {
            return docData;
        }

        const sparseIndexes = this.collectSparseSlideIndexes(docData.slides);
        if (sparseIndexes.length === 0) {
            return docData;
        }

        const heuristicExpanded = this.applyHeuristicSparseExpansion(docData, sparseIndexes, mode, preferences);
        const token = await this.resolveAuthToken();
        if (!token) {
            return heuristicExpanded;
        }

        const expandedByLlm = await this.generateSparseExpansionWithGemini(
            heuristicExpanded,
            sparseIndexes,
            mode,
            token,
            preferences,
        );
        if (expandedByLlm.size === 0) {
            return heuristicExpanded;
        }

        return this.applySparseExpansion(heuristicExpanded, expandedByLlm, mode, preferences);
    }

    private collectSparseSlideIndexes(slides: SlideContent[]): number[] {
        const skipRoles = new Set<SlideRole>(['agenda', 'section_divider', 'summary', 'next_step']);
        const indexes: number[] = [];

        slides.forEach((slide, idx) => {
            if (skipRoles.has(slide.slideRole || 'content')) {
                return;
            }

            const bulletCount = slide.bullets.filter((bullet) => this.cleanText(bullet, 160).length > 0).length;
            const summaryLength = this.cleanText(slide.summary || '', 160).length;
            const keyMessageLength = this.cleanText(slide.keyMessage || '', 160).length;
            const textLength = this.cleanText(`${slide.title} ${slide.summary || ''} ${slide.bullets.join(' ')}`, 500).length;
            const isSparse = bulletCount <= 1 || textLength < 85 || (summaryLength + keyMessageLength < 45 && bulletCount < 2);

            if (isSparse) {
                indexes.push(idx + 1);
            }
        });

        return indexes;
    }

    private applyHeuristicSparseExpansion(
        docData: DocumentData,
        sparseIndexes: number[],
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): DocumentData {
        const sparseSet = new Set<number>(sparseIndexes);
        const slides = docData.slides.map((slide, idx) => {
            if (!sparseSet.has(idx + 1)) {
                return slide;
            }

            const prevTitle = idx > 0 ? this.cleanText(docData.slides[idx - 1].title, 60) : '';
            const nextTitle = idx + 1 < docData.slides.length ? this.cleanText(docData.slides[idx + 1].title, 60) : '';
            const title = this.cleanText(slide.title, 80) || `Slide ${idx + 1}`;
            const existingBullets = this.normalizeBullets(slide.bullets, preferences.deckFormat === 'presenter' ? 4 : 6);
            const fallbackBullets = [
                this.cleanText(`${title}: background and context`, 80),
                prevTitle ? this.cleanText(`Connects from ${prevTitle}`, 80) : '',
                nextTitle ? this.cleanText(`Leads into ${nextTitle}`, 80) : '',
                mode === 'creative'
                    ? this.cleanText(`${title}: practical takeaway for the audience`, 80)
                    : this.cleanText(`${title}: key takeaway`, 80),
            ].filter(Boolean);

            const maxBullets = preferences.deckFormat === 'presenter' ? 4 : 6;
            const bullets = this.normalizeBullets([...existingBullets, ...fallbackBullets], maxBullets);
            const keyMessage =
                this.cleanText(slide.keyMessage || slide.summary || bullets[0] || title, 140) ||
                this.cleanText(`${title}: key message`, 140);
            const summary = this.selectSummary(slide.summary || keyMessage, bullets, title);

            return {
                ...slide,
                bullets,
                keyMessage,
                summary,
                speakerNotes: this.buildSpeakerNotes(title, keyMessage, bullets, slide.slideRole || 'content', preferences),
            };
        });

        return {
            ...docData,
            slides,
        };
    }

    private async generateSparseExpansionWithGemini(
        docData: DocumentData,
        sparseIndexes: number[],
        mode: PlannerMode,
        token: string,
        preferences: PlanningPreferences,
    ): Promise<Map<number, SparseExpansionSlide>> {
        if (this.canUseWorkerProxy()) {
            const viaWorker = await this.generateSparseExpansionWithWorkerProxy(docData, sparseIndexes, mode, preferences);
            if (viaWorker.size > 0) {
                return viaWorker;
            }
        }

        const targetSlides = sparseIndexes.map((index) => {
            const slide = docData.slides[index - 1];
            return {
                index,
                title: slide.title,
                slideRole: slide.slideRole || 'content',
                keyMessage: slide.keyMessage || '',
                summary: slide.summary || '',
                bullets: slide.bullets,
                prevTitle: index > 1 ? docData.slides[index - 2]?.title || '' : '',
                nextTitle: index < docData.slides.length ? docData.slides[index]?.title || '' : '',
                deckGoal: docData.brief?.deckGoal || '',
            };
        });

        const rules = [
            'You are revising weak PPT slides.',
            'Task: expand sparse slides using local deck context.',
            'Do not add unsupported facts, dates, statistics, or named entities.',
            `Deck format: ${preferences.deckFormat}.`,
            mode === 'creative'
                ? 'Mode creative: improve flow and phrasing while staying source-grounded.'
                : 'Mode strict: stay very close to source meaning.',
            'Return 2-4 concise bullets for presenter decks, 3-6 bullets for detailed decks.',
            'Return valid JSON only.',
            'Schema: {"slides":[{"index":1,"keyMessage":"string","summary":"string","bullets":["..."]}]}',
        ].join('\n');

        const payload = {
            prompt: `${rules}\n\nTarget slides:\n${JSON.stringify(targetSlides, null, 2)}`,
            temperature: mode === 'creative' ? 0.35 : 0.15,
            model: this.model,
            stream: false,
        };

        try {
            const response = await axios.post(`${this.baseUrl}/api/llm/direct`, payload, {
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: `Bearer ${token}`,
                },
                timeout: 120000,
                proxy: false,
                validateStatus: () => true,
            });

            if (response.status !== 200 || response.data?.success === false) {
                console.warn(
                    `Sparse expansion API failed: status=${response.status}, message=${response.data?.message || 'unknown'}`,
                );
                return new Map<number, SparseExpansionSlide>();
            }

            const rawText = this.extractModelText(response.data);
            if (!rawText) {
                return new Map<number, SparseExpansionSlide>();
            }

            return this.parseSparseExpansion(rawText, preferences);
        } catch (error: any) {
            console.warn('Sparse expansion request failed:', error?.message || error);
            return new Map<number, SparseExpansionSlide>();
        }
    }

    private async generateSparseExpansionWithWorkerProxy(
        docData: DocumentData,
        sparseIndexes: number[],
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): Promise<Map<number, SparseExpansionSlide>> {
        const targetSlides = sparseIndexes.map((index) => {
            const slide = docData.slides[index - 1];
            return {
                index,
                title: slide.title,
                slideRole: slide.slideRole || 'content',
                keyMessage: slide.keyMessage || '',
                summary: slide.summary || '',
                bullets: slide.bullets,
                prevTitle: index > 1 ? docData.slides[index - 2]?.title || '' : '',
                nextTitle: index < docData.slides.length ? docData.slides[index]?.title || '' : '',
                deckGoal: docData.brief?.deckGoal || '',
            };
        });

        const rules = [
            'You are revising weak PPT slides.',
            'Task: expand sparse slides using local deck context.',
            'Do not add unsupported facts, dates, statistics, or named entities.',
            `Deck format: ${preferences.deckFormat}.`,
            mode === 'creative'
                ? 'Mode creative: improve flow and phrasing while staying source-grounded.'
                : 'Mode strict: stay very close to source meaning.',
            'Return valid JSON only.',
            'Schema: {"slides":[{"index":1,"keyMessage":"string","summary":"string","bullets":["..."]}]}',
        ].join('\n');

        try {
            const text = await this.callGeminiViaWorker(
                `${rules}\n\nTarget slides:\n${JSON.stringify(targetSlides, null, 2)}`,
                mode === 'creative' ? 0.35 : 0.15,
            );
            if (!text) {
                return new Map<number, SparseExpansionSlide>();
            }

            return this.parseSparseExpansion(text, preferences);
        } catch (error: any) {
            console.warn('Sparse expansion worker proxy request failed:', error?.message || error);
            return new Map<number, SparseExpansionSlide>();
        }
    }

    private parseSparseExpansion(raw: string, preferences: PlanningPreferences): Map<number, SparseExpansionSlide> {
        const mapped = new Map<number, SparseExpansionSlide>();
        const jsonText = this.extractJsonBlock(raw);
        if (!jsonText) {
            return mapped;
        }

        try {
            const parsed = JSON.parse(jsonText);
            const slides = Array.isArray(parsed?.slides) ? parsed.slides : [];
            slides.forEach((item: any) => {
                const index = Number(item?.index);
                if (!Number.isFinite(index) || index < 1) {
                    return;
                }

                const bullets = this.normalizeBullets(
                    Array.isArray(item?.bullets) ? item.bullets : [],
                    preferences.deckFormat === 'presenter' ? 4 : 6,
                );
                const keyMessage = this.cleanText(item?.keyMessage, 140);
                const summary = this.cleanText(item?.summary, 140);
                if (bullets.length === 0 && !summary && !keyMessage) {
                    return;
                }

                mapped.set(index, {
                    index,
                    keyMessage,
                    summary,
                    bullets,
                });
            });
        } catch {
            return mapped;
        }

        return mapped;
    }

    private applySparseExpansion(
        docData: DocumentData,
        expandedByLlm: Map<number, SparseExpansionSlide>,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): DocumentData {
        const slides = docData.slides.map((slide, idx) => {
            const expansion = expandedByLlm.get(idx + 1);
            if (!expansion) {
                return slide;
            }

            const maxBullets = preferences.deckFormat === 'presenter' ? 4 : 6;
            const mergedBullets = this.normalizeBullets([...slide.bullets, ...expansion.bullets], maxBullets);
            const fallbackBullets =
                mergedBullets.length >= 2
                    ? mergedBullets
                    : this.normalizeBullets(
                          [
                              ...mergedBullets,
                              `${slide.title}: key information`,
                              mode === 'creative' ? `${slide.title}: audience takeaway` : `${slide.title}: key takeaway`,
                          ],
                          maxBullets,
                      );
            const keyMessage =
                expansion.keyMessage ||
                this.cleanText(slide.keyMessage || slide.summary || fallbackBullets[0] || slide.title, 140);
            const summary = this.selectSummary(expansion.summary || slide.summary || keyMessage, fallbackBullets, slide.title);

            return {
                ...slide,
                bullets: fallbackBullets,
                keyMessage,
                summary,
                speakerNotes: this.buildSpeakerNotes(
                    slide.title,
                    keyMessage,
                    fallbackBullets,
                    slide.slideRole || 'content',
                    preferences,
                ),
            };
        });

        return {
            ...docData,
            slides,
        };
    }

    private buildSpeakerNotes(
        title: string,
        keyMessage: string,
        bullets: string[],
        slideRole: SlideRole,
        preferences: PlanningPreferences,
    ): string[] {
        const notes: string[] = [];
        const cleanTitle = this.cleanText(title, 80);
        const cleanMessage = this.cleanText(keyMessage, 140);
        if (cleanMessage) {
            notes.push(`Primary message: ${cleanMessage}`);
        }

        if (preferences.deckFormat === 'presenter') {
            if (bullets[0]) notes.push(`Open with ${cleanTitle} and explain why it matters.`);
            if (bullets[1]) notes.push(`Expand on: ${this.cleanText(bullets[1], 120)}`);
            if (slideRole === 'summary' || slideRole === 'next_step') {
                notes.push('Close with a clear recommendation or action cue.');
            }
        } else {
            bullets.slice(0, 3).forEach((bullet) => {
                notes.push(`Supporting detail: ${this.cleanText(bullet, 140)}`);
            });
        }

        return notes.slice(0, 4);
    }

    private normalizeBullets(bullets: string[], maxItems: number): string[] {
        const unique = new Set<string>();
        const normalized: string[] = [];

        for (const raw of bullets) {
            const text = this.cleanText(raw, 120);
            if (!text) continue;
            const key = this.normalizeForCompare(text);
            if (!key || unique.has(key)) continue;
            unique.add(key);
            normalized.push(text);
        }

        return normalized.slice(0, maxItems);
    }

    private selectSummary(summary: string, bullets: string[], title: string): string {
        const cleaned = this.cleanText(summary, 140);
        if (!cleaned) return '';
        const normalizedSummary = this.normalizeForCompare(cleaned);
        const bulletNormalized = bullets.map((bullet) => this.normalizeForCompare(bullet));
        if (bulletNormalized.includes(normalizedSummary)) {
            return '';
        }
        if (normalizedSummary === this.normalizeForCompare(title)) {
            return '';
        }
        return cleaned;
    }

    private normalizeSourceRefs(input: any): number[] {
        if (!Array.isArray(input)) {
            return [];
        }

        return Array.from(
            new Set(
                input
                    .map((item) => Number(item))
                    .filter((item) => Number.isFinite(item) && item > 0)
                    .map((item) => Math.floor(item)),
            ),
        ).slice(0, 8);
    }

    private normalizeStringList(input: any, maxItems: number, maxLength: number): string[] {
        if (!Array.isArray(input)) {
            return [];
        }

        return input
            .map((item) => this.cleanText(item, maxLength))
            .filter(Boolean)
            .slice(0, maxItems);
    }

    private normalizeSlideRole(input: any): SlideRole {
        switch (String(input || '').trim()) {
            case 'agenda':
            case 'section_divider':
            case 'key_insight':
            case 'timeline':
            case 'comparison':
            case 'process':
            case 'data_highlight':
            case 'summary':
            case 'next_step':
                return input;
            default:
                return 'content';
        }
    }

    private normalizeLayout(layout: any, slideRole: SlideRole): SlideLayoutType {
        if (layout === 'image_only') {
            return 'image_only';
        }
        if (slideRole === 'section_divider' || slideRole === 'key_insight') {
            return 'image_only';
        }
        return 'image_overlay';
    }

    private heuristicLayout(
        title: string,
        bullets: string[],
        slideRole: SlideRole,
        preferences: PlanningPreferences,
    ): SlideLayoutType {
        if (slideRole === 'section_divider' || slideRole === 'key_insight') {
            return 'image_only';
        }
        if (
            slideRole === 'timeline' ||
            slideRole === 'comparison' ||
            slideRole === 'process' ||
            slideRole === 'summary' ||
            slideRole === 'next_step'
        ) {
            return 'image_overlay';
        }

        const textSize = title.length + bullets.reduce((sum, bullet) => sum + bullet.length, 0);
        if (preferences.deckFormat === 'presenter' && (bullets.length > 4 || textSize > 180)) {
            return 'image_only';
        }
        return 'image_overlay';
    }

    private buildRoleAwareImagePrompt(
        deckTitle: string,
        title: string,
        keyMessage: string,
        bullets: string[],
        slideRole: SlideRole,
        mode: PlannerMode,
        preferences: PlanningPreferences,
    ): string {
        const style =
            preferences.style === 'bold'
                ? 'bold editorial'
                : preferences.style === 'educational'
                  ? 'clean educational'
                  : 'professional cinematic';
        const keywordSeed = [title, keyMessage, ...bullets.slice(0, 2).map((bullet) => this.cleanText(bullet, 80))]
            .filter(Boolean)
            .join('; ');

        if (slideRole === 'agenda' || slideRole === 'summary' || slideRole === 'next_step') {
            return `${keywordSeed}; ${style} 16:9 presentation background for ${deckTitle}, abstract shapes, clear focal point, no text, no watermark.`;
        }

        if (slideRole === 'section_divider') {
            return `${keywordSeed}; ${style} 16:9 section divider visual for ${title}, thematic atmosphere, minimal composition, no text, no watermark.`;
        }

        if (slideRole === 'timeline') {
            return `${keywordSeed}; ${style} 16:9 visual about ${title}, showing evolution and progression, no text, no watermark.`;
        }

        if (slideRole === 'comparison') {
            return `${keywordSeed}; ${style} 16:9 visual illustrating contrast and comparison for ${title}, balanced dual composition, no text, no watermark.`;
        }

        if (slideRole === 'process') {
            return `${keywordSeed}; ${style} 16:9 process illustration for ${title}, step-by-step movement, structured composition, no text, no watermark.`;
        }

        if (slideRole === 'data_highlight') {
            return `${keywordSeed}; ${style} 16:9 visual emphasizing a key metric for ${title}, polished infographic mood, no text, no watermark.`;
        }

        return `${keywordSeed}; ${mode === 'creative' ? 'cinematic' : 'professional'} ${style} 16:9 slide visual about ${title}. No text, no watermark.`;
    }

    private looksLikeTimeline(slide: SlideContent): boolean {
        const text = `${slide.title} ${slide.summary || ''} ${slide.bullets.join(' ')}`;
        return /\b(18|19|20)\d{2}\b|年|阶段|历程|演进|发展史|timeline|history/i.test(text);
    }

    private looksLikeComparison(slide: SlideContent): boolean {
        const text = `${slide.title} ${slide.summary || ''} ${slide.bullets.join(' ')}`;
        return /对比|比较|区别|差异|优势|劣势|vs\b|versus|compare/i.test(text);
    }

    private looksLikeProcess(slide: SlideContent): boolean {
        const text = `${slide.title} ${slide.summary || ''} ${slide.bullets.join(' ')}`;
        const numberedBullets = slide.bullets.filter((bullet) => /^\s*\d+[.)、]/.test(bullet)).length;
        return numberedBullets >= 2 || /流程|步骤|方法|实施|推进|落地|step\b|process|workflow/i.test(text);
    }

    private looksLikeDataHighlight(slide: SlideContent): boolean {
        const text = `${slide.title} ${slide.summary || ''} ${slide.bullets.join(' ')}`;
        const matches = text.match(/\b\d+(?:\.\d+)?%?\b/g) || [];
        return matches.length >= 2 && slide.bullets.length <= 4;
    }

    private isSectionDividerCandidate(slide: SlideContent, index: number, allSlides: SlideContent[]): boolean {
        if (index === 0) return false;
        const level = slide.level || 1;
        if (level > 2) return false;
        const prevLevel = allSlides[index - 1]?.level || 1;
        return level <= prevLevel && slide.bullets.length <= 3 && this.cleanText(slide.title, 80).length <= 28;
    }

    private ensureUniqueTitles(slides: SlideContent[]): SlideContent[] {
        const seen = new Set<string>();
        return slides.map((slide, index) => {
            const base = this.cleanText(slide.title, 80) || `Slide ${index + 1}`;
            let candidate = base;
            let probe = candidate.toLowerCase();
            let dup = 1;

            while (seen.has(probe)) {
                const suffix = slide.slideRole ? slide.slideRole.replace(/_/g, ' ') : 'slide';
                candidate = `${base} - ${suffix}${dup > 1 ? ` ${dup}` : ''}`;
                probe = candidate.toLowerCase();
                dup += 1;
            }

            seen.add(probe);
            return {
                ...slide,
                title: candidate,
            };
        });
    }

    private strengthenNarrativeContinuity(slides: SlideContent[], deckTitle: string): SlideContent[] {
        let prevLevel = 1;

        return slides.map((slide, index) => {
            const role = slide.slideRole || 'content';
            let nextLevel =
                typeof slide.level === 'number' && slide.level > 0
                    ? Math.round(slide.level)
                    : this.defaultLevelForRole(role, prevLevel, index);

            if (index === 0) {
                nextLevel = 1;
            }

            if (Math.abs(nextLevel - prevLevel) > 1) {
                nextLevel = nextLevel > prevLevel ? prevLevel + 1 : prevLevel - 1;
            }

            nextLevel = Math.max(1, Math.min(3, nextLevel));
            const breadcrumb = this.cleanText(slide.breadcrumb || this.buildDefaultBreadcrumb(deckTitle, slide.title, role), 80);

            prevLevel = nextLevel;
            return {
                ...slide,
                level: nextLevel,
                breadcrumb,
            };
        });
    }

    private defaultLevelForRole(role: SlideRole, prevLevel: number, index: number): number {
        if (index === 0 || role === 'agenda') {
            return 1;
        }
        if (role === 'section_divider') {
            return Math.min(prevLevel + 1, 2);
        }
        if (role === 'summary' || role === 'next_step') {
            return prevLevel;
        }
        if (role === 'timeline' || role === 'comparison' || role === 'process' || role === 'data_highlight') {
            return Math.max(2, prevLevel);
        }
        return Math.max(1, Math.min(2, prevLevel));
    }

    private buildDefaultBreadcrumb(deckTitle: string, title: string, role: SlideRole): string {
        const cleanDeck = this.cleanText(deckTitle, 36) || 'Presentation';
        const cleanTitle = this.cleanText(title, 36) || 'Slide';

        switch (role) {
            case 'agenda':
                return `${cleanDeck} / Agenda`;
            case 'section_divider':
                return `${cleanDeck} / Section`;
            case 'timeline':
                return `${cleanDeck} / Timeline`;
            case 'comparison':
                return `${cleanDeck} / Comparison`;
            case 'process':
                return `${cleanDeck} / Process`;
            case 'data_highlight':
                return `${cleanDeck} / Highlight`;
            case 'summary':
                return `${cleanDeck} / Wrap-up`;
            case 'next_step':
                return `${cleanDeck} / Action`;
            case 'key_insight':
                return `${cleanDeck} / Insight`;
            default:
                return `${cleanDeck} / ${cleanTitle}`;
        }
    }

    private isMostlyCjk(text: string): boolean {
        const cleaned = this.cleanText(text, 200);
        const matches = cleaned.match(/[\u4e00-\u9fff]/g) || [];
        return matches.length >= Math.max(2, Math.floor(cleaned.length / 5));
    }

    private mapAudienceLabel(audience: DeckAudience): string {
        switch (audience) {
            case 'beginner':
                return 'beginners';
            case 'executive':
                return 'decision-makers';
            case 'student':
                return 'learners';
            case 'technical':
                return 'technical readers';
            default:
                return 'the audience';
        }
    }

    private mapFocusLabel(focus: DeckFocus): string {
        switch (focus) {
            case 'timeline':
                return 'a timeline-focused';
            case 'argument':
                return 'an argument-led';
            case 'process':
                return 'a process-oriented';
            case 'comparison':
                return 'a comparison-driven';
            default:
                return 'an overview-driven';
        }
    }

    private mapFocusLabelZh(focus: DeckFocus): string {
        switch (focus) {
            case 'timeline':
                return '时间线';
            case 'argument':
                return '论点';
            case 'process':
                return '流程';
            case 'comparison':
                return '对比';
            default:
                return '全局概览';
        }
    }

    private async resolveAuthToken(): Promise<string> {
        if (this.authToken) {
            return this.authToken;
        }

        if (!this.allowGuestLogin) {
            return this.fallbackAuthToken;
        }

        try {
            const response = await axios.post(
                `${this.baseUrl}/api/auth/guest-login`,
                {},
                {
                    headers: { 'Content-Type': 'application/json' },
                    timeout: 30000,
                    proxy: false,
                    validateStatus: () => true,
                },
            );

            if (response.status === 200 && response.data?.success && response.data?.data?.token) {
                this.authToken = response.data.data.token;
                return this.authToken;
            }
        } catch (error: any) {
            console.warn('Guest login for planner failed:', error?.message || error);
        }

        return this.fallbackAuthToken;
    }

    private resolvePlanningPreferences(options: PlannerOptions): PlanningPreferences {
        return {
            deckFormat: this.normalizeDeckFormat(options.deckFormat || process.env.PLANNER_DECK_FORMAT),
            audience: this.normalizeAudience(options.audience || process.env.PLANNER_AUDIENCE),
            focus: this.normalizeFocus(options.focus || process.env.PLANNER_FOCUS),
            style: this.normalizeStyle(options.style || process.env.PLANNER_STYLE),
            length: this.normalizeLength(options.length || process.env.PLANNER_LENGTH),
        };
    }

    private normalizeDeckFormat(input?: any): DeckFormat {
        return input === 'detailed' ? 'detailed' : 'presenter';
    }

    private normalizeAudience(input?: any): DeckAudience {
        switch (String(input || '').trim()) {
            case 'beginner':
            case 'executive':
            case 'student':
            case 'technical':
                return input;
            default:
                return 'general';
        }
    }

    private normalizeFocus(input?: any): DeckFocus {
        switch (String(input || '').trim()) {
            case 'timeline':
            case 'argument':
            case 'process':
            case 'comparison':
                return input;
            default:
                return 'overview';
        }
    }

    private normalizeStyle(input?: any): DeckStyle {
        switch (String(input || '').trim()) {
            case 'minimal':
            case 'bold':
            case 'educational':
                return input;
            default:
                return 'professional';
        }
    }

    private normalizeLength(input?: any): DeckLength {
        switch (String(input || '').trim()) {
            case 'short':
            case 'long':
                return input;
            default:
                return 'default';
        }
    }

    private canUseWorkerProxy(): boolean {
        return Boolean(this.workerUrl && this.workerApiKey);
    }

    private async callGeminiViaWorker(prompt: string, temperature: number): Promise<string> {
        if (!this.canUseWorkerProxy()) {
            return '';
        }

        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${this.model}:generateContent?key=${this.workerApiKey}`;
        const requestData = {
            contents: [
                {
                    role: 'user',
                    parts: [{ text: prompt }],
                },
            ],
            generationConfig: {
                temperature,
            },
        };

        const response = await axios.post(
            this.workerUrl,
            {
                url: apiUrl,
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                data: requestData,
            },
            {
                headers: {
                    'Content-Type': 'application/json',
                },
                timeout: 120000,
                proxy: false,
                validateStatus: () => true,
            },
        );

        if (response.status !== 200) {
            const body = this.stringifyErrorPayload(response.data);
            throw new Error(`Worker proxy failed: status=${response.status}, body=${body}`);
        }

        return this.cleanText(this.extractGeminiTextFromWorkerResponse(response.data), 20000);
    }

    private extractGeminiTextFromWorkerResponse(payload: any): string {
        if (!payload) return '';

        if (typeof payload === 'string') {
            try {
                const parsed = JSON.parse(payload);
                return this.extractGeminiTextFromWorkerResponse(parsed);
            } catch {
                return payload;
            }
        }

        const fromParts = (root: any): string => {
            const candidates = root?.candidates;
            if (!Array.isArray(candidates) || candidates.length === 0) return '';
            const parts = candidates[0]?.content?.parts;
            if (!Array.isArray(parts)) return '';
            return parts
                .map((part: any) => (typeof part?.text === 'string' ? part.text : ''))
                .filter(Boolean)
                .join('\n')
                .trim();
        };

        const direct = fromParts(payload);
        if (direct) return direct;

        if (typeof payload?.data === 'string') {
            return payload.data;
        }

        const nested = fromParts(payload?.data);
        if (nested) return nested;

        const outputText = payload?.output_text || payload?.data?.output_text;
        if (typeof outputText === 'string') return outputText;

        return '';
    }

    private loadExternalPlannerEnv(): Record<string, string> {
        const envPath =
            process.env.PLANNER_AIWORKFLOW_ENV_PATH || process.env.AIWORKFLOW_BACKEND_ENV_PATH || 'E:\\GitHubWorkSpace\\aiworkflow\\back-end\\.env';
        if (!envPath || !fs.existsSync(envPath)) {
            return {};
        }

        try {
            const raw = fs.readFileSync(envPath, 'utf-8');
            const map: Record<string, string> = {};
            raw.split(/\r?\n/).forEach((line) => {
                const trimmed = line.trim();
                if (!trimmed || trimmed.startsWith('#')) return;
                const eqIndex = trimmed.indexOf('=');
                if (eqIndex <= 0) return;
                const key = trimmed.slice(0, eqIndex).trim();
                const value = trimmed.slice(eqIndex + 1).trim().replace(/^"|"$/g, '');
                if (key) {
                    map[key] = value;
                }
            });
            return map;
        } catch {
            return {};
        }
    }

    private stringifyErrorPayload(payload: any): string {
        try {
            if (typeof payload === 'string') return payload;
            return JSON.stringify(payload);
        } catch {
            return String(payload);
        }
    }

    private resolvePlannerMode(mode?: PlannerMode): PlannerMode {
        if (mode === 'creative' || mode === 'strict') {
            return mode;
        }

        return this.defaultMode;
    }

    private normalizeForCompare(text: string): string {
        return this.cleanText(text, 300)
            .toLowerCase()
            .replace(/[\s\p{P}\p{S}]+/gu, '');
    }

    private cleanText(input: any, maxLength: number): string {
        if (typeof input !== 'string') return '';

        let text = input
            .replace(/\r?\n+/g, ' ')
            .replace(/\s+/g, ' ')
            .replace(/[鈥溾€漖]/g, '"')
            .replace(/[鈥樷€橾]/g, "'")
            .replace(/[\u0000-\u001f]/g, '')
            .trim();

        if (text.length > maxLength) {
            text = `${text.slice(0, maxLength - 3).trim()}...`;
        }

        return text;
    }
}
