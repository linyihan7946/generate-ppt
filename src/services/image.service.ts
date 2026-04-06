import axios from 'axios';
import { SlideContent } from '../types';

export class ImageService {
    private apiKey: string | undefined;
    private baseUrl: string;
    private cache = new Map<string, string>();

    constructor() {
        this.apiKey = process.env.IMAGE_API_KEY;
        this.baseUrl = process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn';
        console.log('ImageService initialized with baseUrl:', this.baseUrl);
    }

    async enrichSlidesWithGeneratedImages(slides: SlideContent[], concurrency = 2): Promise<void> {
        const jobs = slides.map((slide) => async () => {
            if (slide.images.length > 0) return;

            const prompt = this.buildPrompt(slide);
            const imageData = await this.generateImage(prompt);
            if (imageData) {
                slide.images.push(imageData);
            }
        });

        await this.runWithConcurrency(jobs, concurrency);
    }

    async generateImage(prompt: string): Promise<string | null> {
        const cacheKey = prompt.trim();
        if (this.cache.has(cacheKey)) {
            return this.cache.get(cacheKey) || null;
        }

        const primaryImage = await this.generateByPrimaryApi(prompt);
        if (primaryImage) {
            this.cache.set(cacheKey, primaryImage);
            return primaryImage;
        }

        console.log('图片生成失败，尝试使用简化提示词重试...');
        const safePrompt = `A professional abstract presentation background about technology. Minimalist style, high quality, 4k. No text.`;
        const safeImage = await this.generateByPrimaryApi(safePrompt);
        if (safeImage) {
            this.cache.set(cacheKey, safeImage);
            return safeImage;
        }

        const fallbackImage = await this.generateByFallback(prompt);
        if (fallbackImage) {
            this.cache.set(cacheKey, fallbackImage);
            return fallbackImage;
        }

        return null;
    }

    private async generateByPrimaryApi(prompt: string): Promise<string | null> {
        try {
            const headers: Record<string, string> = {
                'Content-Type': 'application/json',
            };

            if (this.apiKey) {
                headers.Authorization = `Bearer ${this.apiKey}`;
            }

            const response = await axios.post(
                `${this.baseUrl}/api/image/direct-edit`,
                {
                    prompt,
                    images: [],
                    model: 'gemini-3.1-flash-image-preview',
                    aspect_ratio: '16:9',
                    resolution: '2K',
                },
                {
                    headers,
                    timeout: 120000,
                    proxy: false,
                },
            );

            const rawResult = response.data?.data?.data?.[0];
            const imagePayload =
                typeof rawResult === 'string'
                    ? rawResult
                    : rawResult?.url || rawResult?.b64_json || rawResult?.base64;

            if (!imagePayload) {
                return null;
            }

            return await this.normalizeImagePayload(imagePayload);
        } catch (error: any) {
            console.error('Primary image API failed:', error?.message || error);
            if (error?.response?.data) {
                console.error('Primary image API response:', JSON.stringify(error.response.data));
            }
            return null;
        }
    }

    private async generateByFallback(prompt: string): Promise<string | null> {
        // Fallback keeps the "auto image" behavior available even when the primary image API is unstable.
        const seed = this.hashPrompt(prompt);
        const candidates = [
            `https://picsum.photos/seed/${seed}/1600/900`,
            `https://dummyimage.com/1600x900/0f172a/e2e8f0.png&text=AI+Illustration+${seed}`,
        ];

        for (const url of candidates) {
            const image = await this.downloadAsDataUrl(url);
            if (image) {
                return image;
            }
        }

        return this.localPlaceholderImage();
    }

    private async normalizeImagePayload(payload: string): Promise<string | null> {
        if (payload.startsWith('data:image')) {
            return payload;
        }

        if (payload.startsWith('http://') || payload.startsWith('https://')) {
            return await this.downloadAsDataUrl(payload);
        }

        if (this.looksLikeBase64(payload)) {
            return `data:image/png;base64,${payload}`;
        }

        return null;
    }

    private looksLikeBase64(value: string): boolean {
        return /^[A-Za-z0-9+/=\r\n]+$/.test(value) && value.length > 100;
    }

    private async downloadAsDataUrl(url: string): Promise<string | null> {
        try {
            const response = await axios.get(url, {
                responseType: 'arraybuffer',
                timeout: 45000,
                proxy: false,
            });
            const contentType = (response.headers['content-type'] as string) || 'image/png';
            const base64 = Buffer.from(response.data).toString('base64');
            return `data:${contentType};base64,${base64}`;
        } catch (error: any) {
            console.error('Failed to download image:', error?.message || error);
            return null;
        }
    }

    private buildPrompt(slide: SlideContent): string {
        const cleanTitle = slide.title.replace(/\(Cont\..*\)/, '');
        const context = slide.breadcrumb ? `Context: ${slide.breadcrumb}.` : '';
        // Clean up bullets to remove potential sensitive words or complex characters
        const cleanedBullets = slide.bullets.slice(0, 2).map(b => b.replace(/[^\w\s\u4e00-\u9fa5,.，。]/g, ''));
        const keyPoints = cleanedBullets.join(', ');

        return [
            `A professional and modern presentation slide illustration about: ${cleanTitle}.`,
            context,
            keyPoints ? `Key concepts: ${keyPoints}.` : '',
            'Minimalist style, tech-oriented, clean background, high quality, 4k. No text, no people.',
        ]
            .filter(Boolean)
            .join(' ');
    }

    private hashPrompt(prompt: string): string {
        let hash = 2166136261;
        for (let i = 0; i < prompt.length; i++) {
            hash ^= prompt.charCodeAt(i);
            hash +=
                (hash << 1) +
                (hash << 4) +
                (hash << 7) +
                (hash << 8) +
                (hash << 24);
        }
        return Math.abs(hash >>> 0).toString(36);
    }

    private localPlaceholderImage(): string {
        // 1x1 neutral pixel as last-resort fallback so slides never miss an image slot.
        return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6WQ9QAAAAASUVORK5CYII=';
    }

    private async runWithConcurrency(jobs: Array<() => Promise<void>>, concurrency: number): Promise<void> {
        if (jobs.length === 0) return;

        const safeConcurrency = Math.max(1, Math.min(concurrency, jobs.length));
        let cursor = 0;

        const workers = Array.from({ length: safeConcurrency }, async () => {
            while (true) {
                const jobIndex = cursor++;
                if (jobIndex >= jobs.length) {
                    break;
                }
                await jobs[jobIndex]();
            }
        });

        await Promise.all(workers);
    }
}
