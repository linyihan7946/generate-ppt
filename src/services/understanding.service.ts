import { DocumentData, UnderstandingResult, UnderstandingTopic } from '../types';

export class UnderstandingService {
    analyze(docData: DocumentData): UnderstandingResult {
        const chapterTitles = this.collectChapterTitles(docData);
        const topics = this.collectTopics(docData);
        const timelineSignals = this.collectPatternMatches(docData, /\b(18|19|20)\d{2}\b|年|阶段|历程|演进|发展史|timeline|history/gi);
        const comparisonSignals = this.collectPatternMatches(docData, /对比|比较|区别|差异|优势|劣势|vs\b|versus|compare/gi);
        const processSignals = this.collectPatternMatches(docData, /流程|步骤|方法|实施|推进|落地|step\b|process|workflow/gi);
        const keyNumbers = this.collectPatternMatches(docData, /\b\d+(?:\.\d+)?%?\b/g);
        const thesis = this.buildThesis(docData, topics);

        return {
            thesis,
            chapterTitles,
            topics,
            timelineSignals,
            comparisonSignals,
            processSignals,
            keyNumbers,
        };
    }

    private collectChapterTitles(docData: DocumentData): string[] {
        const seen = new Set<string>();
        const chapters: string[] = [];

        docData.slides.forEach((slide) => {
            const title = this.cleanText(slide.title);
            if (!title) return;
            if ((slide.level || 1) > 2 && chapters.length > 0) return;
            const key = title.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            chapters.push(title);
        });

        if (chapters.length > 0) {
            return chapters.slice(0, 8);
        }

        return docData.slides
            .map((slide) => this.cleanText(slide.title))
            .filter(Boolean)
            .slice(0, 6);
    }

    private collectTopics(docData: DocumentData): UnderstandingTopic[] {
        return docData.slides
            .map((slide, index) => {
                const title = this.cleanText(slide.title);
                const bulletWeight = slide.bullets.filter((bullet) => this.cleanText(bullet).length > 0).length;
                if (!title) {
                    return null;
                }

                return {
                    title,
                    sourceRefs: [slide.sourceIndex || index + 1],
                    importance: Math.max(1, 6 - Math.min(5, slide.level || 1)) + Math.min(3, bulletWeight),
                } as UnderstandingTopic;
            })
            .filter((topic): topic is UnderstandingTopic => Boolean(topic))
            .sort((left, right) => right.importance - left.importance)
            .slice(0, 10);
    }

    private collectPatternMatches(docData: DocumentData, pattern: RegExp): string[] {
        const text = docData.slides
            .map((slide) => [slide.title, slide.summary || '', slide.bullets.join(' ')].join(' '))
            .join(' ');
        const matches = text.match(pattern) || [];
        const unique = Array.from(new Set(matches.map((item) => this.cleanText(item)).filter(Boolean)));
        return unique.slice(0, 12);
    }

    private buildThesis(docData: DocumentData, topics: UnderstandingTopic[]): string {
        const firstSummary = docData.slides.find((slide) => this.cleanText(slide.summary || '').length > 0)?.summary || '';
        if (this.cleanText(firstSummary)) {
            return this.cleanText(firstSummary);
        }

        const firstBullet = docData.slides.flatMap((slide) => slide.bullets).map((bullet) => this.cleanText(bullet)).find(Boolean);
        if (firstBullet) {
            return firstBullet;
        }

        return topics[0]?.title || this.cleanText(docData.title) || 'Core presentation thesis';
    }

    private cleanText(input: any): string {
        if (typeof input !== 'string') return '';
        return input.replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
    }
}
