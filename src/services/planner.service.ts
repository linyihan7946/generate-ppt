import axios from 'axios';
import { DocumentData, PlannerMode, PlannerOptions, SlideContent, SlideLayoutType } from '../types';

interface PlannedSlide {
    index: number;
    title: string;
    summary: string;
    bullets: string[];
    layout: SlideLayoutType;
    imageIntent: string;
    imagePrompt: string;
}

interface PlannedDocument {
    title?: string;
    slides: PlannedSlide[];
}

export class PlannerService {
    private baseUrl: string;
    private authToken: string;
    private fallbackAuthToken: string;
    private model: string;
    private enabled: boolean;
    private allowGuestLogin: boolean;
    private defaultMode: PlannerMode;

    constructor() {
        this.baseUrl = process.env.PLANNER_API_BASE_URL || process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn:3001';
        this.authToken = process.env.PLANNER_AUTH_TOKEN || process.env.LLM_AUTH_TOKEN || '';
        this.fallbackAuthToken = process.env.IMAGE_API_KEY || '';
        this.model = process.env.PLANNER_MODEL || 'gemini-3.1-pro-preview';
        this.enabled = process.env.ENABLE_PLANNER !== 'false';
        this.allowGuestLogin = process.env.PLANNER_USE_GUEST_LOGIN === 'true';
        this.defaultMode = process.env.PLANNER_CONTENT_MODE === 'creative' ? 'creative' : 'strict';
    }

    async planDocument(docData: DocumentData, options: PlannerOptions = {}): Promise<DocumentData> {
        const mode = this.resolvePlannerMode(options.mode);
        const heuristic = this.buildHeuristicPlan(docData, mode);
        if (!this.enabled) {
            return heuristic;
        }

        const llmPlan = await this.generatePlanWithGemini(docData, mode);
        if (!llmPlan) {
            return heuristic;
        }

        return this.mergePlan(heuristic, llmPlan);
    }

    private async generatePlanWithGemini(docData: DocumentData, mode: PlannerMode): Promise<PlannedDocument | null> {
        const token = await this.resolveAuthToken();
        if (!token) {
            console.warn('Planner skipped: missing PLANNER_AUTH_TOKEN / LLM_AUTH_TOKEN / IMAGE_API_KEY.');
            return null;
        }

        const systemPrompt = [
            'You are a professional presentation strategist.',
            'You must preserve source hierarchy and slide order.',
            this.buildModeDirective(mode),
            'Return only valid JSON with no markdown fences and no extra text.',
        ].join(' ');

        const userPrompt = this.buildUserPrompt(docData, mode);
        const payload = {
            prompt: userPrompt,
            messages: [
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt },
            ],
            temperature: mode === 'creative' ? 0.35 : 0.15,
            useSearch: false,
            model: this.model,
            stream: false,
            sessionId: `ppt-planner-${Date.now()}`,
        };

        try {
            const response = await axios.post(`${this.baseUrl}/api/llm`, payload, {
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

    private async resolveAuthToken(): Promise<string> {
        if (this.authToken) {
            return this.authToken;
        }

        if (!this.allowGuestLogin) {
            return '';
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

    private buildUserPrompt(docData: DocumentData, mode: PlannerMode): string {
        const compactSlides = docData.slides.map((slide, idx) => ({
            index: idx + 1,
            title: slide.title,
            bullets: slide.bullets,
            level: slide.level || 1,
            breadcrumb: slide.breadcrumb || '',
        }));

        const rules = [
            'Task: build a slide-by-slide planning JSON for a PPT generator.',
            'Must keep slide order and hierarchy. Do NOT merge, split, or reorder slides.',
            'Each output slide index must match input index.',
            ...this.buildModeRules(mode),
            'layout must be one of: image_overlay, image_only.',
            'imagePrompt must be <= 220 chars, concise, visual, no text/watermark.',
            'Output schema:',
            '{"title":"string","slides":[{"index":1,"title":"string","summary":"string","bullets":["..."],"layout":"image_overlay","imageIntent":"string","imagePrompt":"string"}]}',
        ].join('\n');

        return `${rules}\n\nDocument:\n${JSON.stringify({ title: docData.title, slides: compactSlides }, null, 2)}`;
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

        const choiceText = payload.choices?.[0]?.message?.content;
        if (typeof choiceText === 'string') return choiceText;
        return '';
    }

    private parsePlannedDocument(raw: string): PlannedDocument | null {
        const jsonText = this.extractJsonBlock(raw);
        if (!jsonText) return null;

        try {
            const parsed = JSON.parse(jsonText);
            if (!parsed || !Array.isArray(parsed.slides)) return null;

            const slides: PlannedSlide[] = parsed.slides
                .map((slide: any) => this.normalizePlannedSlide(slide))
                .filter((slide: PlannedSlide | null): slide is PlannedSlide => Boolean(slide));

            if (slides.length === 0) return null;
            return {
                title: typeof parsed.title === 'string' ? this.cleanText(parsed.title, 80) : undefined,
                slides,
            };
        } catch {
            return null;
        }
    }

    private normalizePlannedSlide(input: any): PlannedSlide | null {
        const index = Number(input?.index);
        if (!Number.isFinite(index) || index < 1) {
            return null;
        }

        const title = this.cleanText(input?.title, 60);
        const summary = this.cleanText(input?.summary, 120);
        const bullets = this.normalizeBullets(Array.isArray(input?.bullets) ? input.bullets : []);
        const layout = this.normalizeLayout(input?.layout);
        const imageIntent = this.cleanText(input?.imageIntent, 160);
        const imagePrompt = this.cleanText(input?.imagePrompt, 220);

        return {
            index,
            title: title || `Slide ${index}`,
            summary: this.selectSummary(summary || title || `Slide ${index}`, bullets, title),
            bullets,
            layout,
            imageIntent: imageIntent || title,
            imagePrompt: imagePrompt || this.cleanText(`${title}. ${bullets.slice(0, 2).join(' ')}`, 220),
        };
    }

    private normalizeLayout(layout: any): SlideLayoutType {
        return layout === 'image_only' ? 'image_only' : 'image_overlay';
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

    private buildHeuristicPlan(docData: DocumentData, mode: PlannerMode): DocumentData {
        const rawSlides = docData.slides.map((slide, idx) => {
            const title = this.cleanText(slide.title, 60) || `Slide ${idx + 1}`;
            const bullets = this.normalizeBullets(slide.bullets);
            const summary = this.buildHeuristicSummary(title, bullets, mode);
            const layout = this.heuristicLayout(title, bullets);
            const imageIntent = this.cleanText(
                [slide.breadcrumb, title, bullets.slice(0, 2).join(' ')].filter(Boolean).join(' | '),
                160,
            );
            const imagePrompt = this.heuristicImagePrompt(title, bullets, slide.breadcrumb, mode);

            return {
                ...slide,
                title,
                bullets,
                summary,
                layout,
                imageIntent,
                imagePrompt,
                sourceIndex: idx + 1,
            } as SlideContent;
        });

        const slides = this.ensureUniqueTitles(rawSlides);
        return {
            title: this.cleanText(docData.title, 80) || 'Presentation',
            slides,
        };
    }

    private mergePlan(base: DocumentData, llmPlan: PlannedDocument): DocumentData {
        const byIndex = new Map<number, PlannedSlide>();
        llmPlan.slides.forEach((slide) => byIndex.set(slide.index, slide));

        const mergedSlides = base.slides.map((baseSlide, idx) => {
            const planned = byIndex.get(idx + 1);
            if (!planned) {
                return baseSlide;
            }

            const bullets = planned.bullets.length > 0 ? planned.bullets : baseSlide.bullets;
            const title = planned.title || baseSlide.title;
            const summary = this.selectSummary(planned.summary || baseSlide.summary || '', bullets, title);
            return {
                ...baseSlide,
                title,
                summary,
                bullets,
                layout: planned.layout || baseSlide.layout,
                imageIntent: planned.imageIntent || baseSlide.imageIntent,
                imagePrompt: planned.imagePrompt || baseSlide.imagePrompt,
                sourceIndex: idx + 1,
            } as SlideContent;
        });

        const slides = this.ensureUniqueTitles(mergedSlides);
        return {
            title: llmPlan.title || base.title,
            slides,
        };
    }

    private normalizeBullets(bullets: string[]): string[] {
        const unique = new Set<string>();
        const normalized: string[] = [];

        for (const raw of bullets) {
            const text = this.cleanText(raw, 80);
            if (!text) continue;
            const key = this.normalizeForCompare(text);
            if (!key) continue;
            if (unique.has(key)) continue;
            unique.add(key);
            normalized.push(text);
        }

        return normalized.slice(0, 5);
    }

    private buildHeuristicSummary(title: string, bullets: string[], mode: PlannerMode): string {
        if (bullets.length === 0) {
            return title;
        }

        const base = title.replace(/[：:]\s*.*$/, '').trim() || title;
        const summary = `${base}的关键变化与影响`;
        const finalSummary =
            mode === 'creative'
                ? this.cleanText(`${title}. ${bullets.slice(0, 2).join(' ')}`, 120) || summary
                : summary;
        return this.selectSummary(finalSummary, bullets, title);
    }

    private selectSummary(summary: string, bullets: string[], title: string): string {
        const cleaned = this.cleanText(summary, 120);
        if (!cleaned) return '';
        if (this.isRedundantSummary(cleaned, bullets, title)) return '';
        return cleaned;
    }

    private isRedundantSummary(summary: string, bullets: string[], title: string): boolean {
        const summaryNormalized = this.normalizeForCompare(summary);
        if (!summaryNormalized) {
            return true;
        }

        const stripped = this.normalizeForCompare(
            summary.replace(new RegExp(`^\\s*${this.escapeRegExp(title)}\\s*[:：,，。\\-]*\\s*`), ''),
        );
        const candidates = [summaryNormalized, stripped].filter(Boolean);
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
        return this.cleanText(text, 300)
            .toLowerCase()
            .replace(/[\s\p{P}\p{S}]+/gu, '');
    }

    private escapeRegExp(input: string): string {
        return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    private heuristicLayout(title: string, bullets: string[]): SlideLayoutType {
        const textSize = title.length + bullets.reduce((sum, b) => sum + b.length, 0);
        if (bullets.length > 4 || textSize > 170) {
            return 'image_only';
        }
        return 'image_overlay';
    }

    private heuristicImagePrompt(
        title: string,
        bullets: string[],
        breadcrumb?: string,
        mode: PlannerMode = 'strict',
    ): string {
        const combined = [title, ...bullets, breadcrumb || ''].join(' ');
        if (this.isSensitiveEntity(combined)) {
            const context = this.cleanText(this.stripSensitiveTerms(`${breadcrumb || ''} ${title}`), 52);
            const era = /互联网|搜索|电商|邮件|聊天|平台/i.test(combined)
                ? 'internet platform ecosystem'
                : /PC|微机|处理器|台式|大型机|硬件/i.test(combined)
                  ? 'early personal computer era'
                  : 'technology industry development';

            const prompt = [
                `${mode === 'creative' ? 'Cinematic' : 'Professional'} 16:9 illustration of ${era}.`,
                context ? `Theme keywords: ${context}.` : '',
                mode === 'creative'
                    ? 'Visual elements: iconic devices, team collaboration, layered data network, sense of momentum and innovation.'
                    : 'Visual elements: computer devices, team collaboration, data network, innovation.',
                'No text, no logo, no watermark.',
            ]
                .filter(Boolean)
                .join(' ');

            return this.cleanText(prompt, 220);
        }

        const ingredients = [title, ...bullets.slice(0, 2)].filter(Boolean).join('; ');
        const prompt = [
            `${mode === 'creative' ? 'Cinematic' : 'Professional'} 16:9 slide illustration about ${ingredients}.`,
            breadcrumb ? `Context: ${this.cleanText(breadcrumb, 60)}.` : '',
            mode === 'creative'
                ? 'Modern, layered composition with clear focal point, polished lighting, no text, no watermark.'
                : 'Modern, clean composition, no text, no watermark.',
        ]
            .filter(Boolean)
            .join(' ');

        return this.cleanText(prompt, 220);
    }

    private isSensitiveEntity(text: string): boolean {
        const patterns = [
            /MITS/i,
            /Altair/i,
            /苹果|微软|联想|百度|阿里巴巴|网易|腾讯|谷歌/,
            /乔布斯|比尔盖茨|柳传志|李彦宏|马云|马化腾|丁磊|罗伯茨/,
            /qq聊天|basic[:：]/i,
        ];
        return patterns.some((pattern) => pattern.test(text));
    }

    private stripSensitiveTerms(text: string): string {
        return text
            .replace(/MITS|Altair|苹果|微软|联想|百度|阿里巴巴|网易|腾讯|谷歌/gi, '科技企业')
            .replace(/乔布斯|比尔盖茨|柳传志|李彦宏|马云|马化腾|丁磊|罗伯茨/gi, '代表人物')
            .replace(/qq聊天|basic[:：]?/gi, '核心业务')
            .replace(/\s+/g, ' ')
            .trim();
    }

    private ensureUniqueTitles(slides: SlideContent[]): SlideContent[] {
        const seen = new Set<string>();
        const result: SlideContent[] = [];

        slides.forEach((slide, index) => {
            const base = this.cleanText(slide.title, 60) || `Slide ${index + 1}`;
            let candidate = base;
            let probe = candidate.toLowerCase();
            let dup = 1;

            while (seen.has(probe)) {
                const tail = this.lastBreadcrumbSegment(slide.breadcrumb) || `L${slide.level || 1}`;
                candidate = `${base} - ${tail}${dup > 1 ? `-${dup}` : ''}`;
                probe = candidate.toLowerCase();
                dup += 1;
            }

            seen.add(probe);
            result.push({ ...slide, title: candidate });
        });

        return result;
    }

    private lastBreadcrumbSegment(breadcrumb?: string): string {
        if (!breadcrumb) return '';
        const parts = breadcrumb.split('/').map((p) => p.trim()).filter(Boolean);
        if (parts.length === 0) return '';
        return this.cleanText(parts[parts.length - 1], 24);
    }

    private resolvePlannerMode(mode?: PlannerMode): PlannerMode {
        if (mode === 'creative' || mode === 'strict') {
            return mode;
        }

        return this.defaultMode;
    }

    private buildModeDirective(mode: PlannerMode): string {
        if (mode === 'creative') {
            return 'Use source-grounded creativity: lightly polish phrasing, sharpen takeaways, and enrich visual direction without changing facts.';
        }

        return 'Use maximum source fidelity: stay close to the source wording and meaning, with no unsupported additions.';
    }

    private buildModeRules(mode: PlannerMode): string[] {
        if (mode === 'creative') {
            return [
                'Mode: creative. You may lightly polish wording, sharpen takeaways, and add small connective phrasing for better presentation flow.',
                'Do not introduce unsupported factual claims, dates, statistics, or named entities.',
                'Bullets: keep 2-5 concise bullets grounded in the source, but you may rewrite them to be more presentation-ready.',
                'summary should be presentation-oriented, vivid, and still fully grounded in the source.',
            ];
        }

        return [
            'Mode: strict. Stay as close as possible to the source content and wording.',
            'Do not introduce new facts, examples, interpretations, dates, or named entities.',
            'Bullets: keep 2-5 concise bullets extracted directly from source meaning.',
            'summary should be a faithful paraphrase with no extra interpretation.',
        ];
    }

    private cleanText(input: any, maxLength: number): string {
        if (typeof input !== 'string') return '';

        let text = input
            .replace(/\r?\n+/g, ' ')
            .replace(/\s+/g, ' ')
            .replace(/[“”]/g, '"')
            .replace(/[‘’]/g, "'")
            .replace(/[\u0000-\u001f]/g, '')
            .trim();

        if (text.length > maxLength) {
            text = `${text.slice(0, maxLength - 3).trim()}...`;
        }

        return text;
    }
}
