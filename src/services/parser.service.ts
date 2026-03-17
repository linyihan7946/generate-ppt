import * as fs from 'fs';
import { marked } from 'marked';
import mammoth from 'mammoth';
import pdfParse from 'pdf-parse';
import { DocumentData, SlideContent } from '../types';

export class ParserService {
    async parseMarkdown(filePath: string): Promise<DocumentData> {
        const content = fs.readFileSync(filePath, 'utf-8');
        const tokens = marked.lexer(content);
        
        const slides: SlideContent[] = [];
        let currentSlide: SlideContent | null = null;

        tokens.forEach(token => {
            if (token.type === 'heading') {
                if (currentSlide) slides.push(currentSlide);
                currentSlide = { title: token.text, bullets: [], images: [] };
            } else if (token.type === 'list') {
                token.items.forEach(item => {
                    currentSlide?.bullets.push(item.text);
                });
            } else if (token.type === 'paragraph') {
                if (currentSlide) {
                    currentSlide.bullets.push(token.text);
                } else {
                    currentSlide = { title: token.text, bullets: [], images: [] };
                }
            } else if (token.type === 'image') {
                currentSlide?.images.push(token.href);
            }
        });

        if (currentSlide) slides.push(currentSlide);

        return { title: 'Markdown Presentation', slides };
    }

    async parseDocx(filePath: string): Promise<DocumentData> {
        const buffer = fs.readFileSync(filePath);
        
        const { value: html } = await mammoth.convertToHtml({ buffer }, {
            convertImage: mammoth.images.imgElement(async (image) => {
                const base64 = await image.readAsBase64String();
                return { src: `data:${image.contentType};base64,${base64}` };
            })
        });
        
        let slides: SlideContent[] = [];

        // Check if document has heading tags
        if (/<h[1-6]>/i.test(html)) {
            slides = this.parseByHeadings(html);
        } 
        // Check if document has list structure
        else if (/<ol>/i.test(html) || /<ul>/i.test(html)) {
            // Try to parse by list, but if fails or returns single slide with too much content, fallback
            const listSlides = this.parseByNestedList(html);
            if (listSlides.length > 0 && listSlides[0].bullets.length < 20) {
                slides = listSlides;
            } else {
                 slides = this.parseByParagraphs(html);
            }
        }
        // Fallback: parse by paragraphs
        else {
            slides = this.parseByParagraphs(html);
        }
        
        // If we still have only one slide with many bullets, force split it
        if (slides.length === 1 && slides[0].bullets.length > 8) {
             const originalSlide = slides[0];
             const newSlides: SlideContent[] = [];
             const bullets = originalSlide.bullets;
             const ITEMS_PER_SLIDE = 6;
             
             for (let i = 0; i < bullets.length; i += ITEMS_PER_SLIDE) {
                 const chunk = bullets.slice(i, i + ITEMS_PER_SLIDE);
                 const title = i === 0 ? originalSlide.title : `${originalSlide.title} (Cont. ${Math.floor(i/ITEMS_PER_SLIDE) + 1})`;
                 newSlides.push({
                     title,
                     bullets: chunk,
                     images: i === 0 ? originalSlide.images : []
                 });
             }
             slides = newSlides;
        }

        // Extract title from first slide if exists
        const title = slides.length > 0 ? slides[0].title : 'Docx Presentation';

        return { title, slides };
    }

    private parseByHeadings(html: string): SlideContent[] {
        const slides: SlideContent[] = [];
        // Split by headings h1-h6
        const parts = html.split(/(<h[1-6][^>]*>.*?<\/h[1-6]>)/i);
        
        let currentSlide: SlideContent | null = null;
        
        for (let i = 0; i < parts.length; i++) {
            const part = parts[i];
            if (!part.trim()) continue;
            
            // Check if this part is a heading
            const headingMatch = part.match(/<h([1-6])[^>]*>(.*?)<\/h\1>/i);
            
            if (headingMatch) {
                // If we have a current slide, push it
                if (currentSlide) {
                    slides.push(currentSlide);
                }
                
                // Start a new slide
                const title = headingMatch[2].replace(/<[^>]+>/g, '').trim();
                currentSlide = { title, bullets: [], images: [] };
            } else if (currentSlide) {
                // Content for the current slide
                const bullets = this.extractBulletsFromHtml(part);
                const images = this.extractImagesFromHtml(part);
                
                currentSlide.bullets.push(...bullets);
                currentSlide.images.push(...images);
            } else {
                // Content before the first heading - create a default slide or ignore
                // For now, let's create a title slide if it looks like a title
                const text = part.replace(/<[^>]+>/g, '').trim();
                if (text) {
                     currentSlide = { title: text.substring(0, 50), bullets: [], images: [] };
                }
            }
        }
        
        if (currentSlide) {
            slides.push(currentSlide);
        }

        return slides;
    }
    
    private parseByParagraphs(html: string): SlideContent[] {
        const slides: SlideContent[] = [];
        // Split by paragraphs
        const paragraphs = html.split(/<p[^>]*>/i);
        
        let currentSlide: SlideContent | null = null;
        let slideLineCount = 0;
        const LINES_PER_SLIDE = 8;
        
        for (const p of paragraphs) {
            const cleanText = p.replace(/<\/p>/i, '').replace(/<[^>]+>/g, '').trim();
            if (!cleanText) continue;
            
            // Heuristic: If text is short and bold/strong, or looks like a title, start new slide
            // For now, just simple length check or if no slide exists
            const isTitle = !currentSlide || (cleanText.length < 50 && slideLineCount >= LINES_PER_SLIDE);
            
            if (isTitle) {
                if (currentSlide) slides.push(currentSlide);
                currentSlide = { title: cleanText, bullets: [], images: [] };
                slideLineCount = 0;
            } else if (currentSlide) {
                currentSlide.bullets.push(cleanText);
                slideLineCount++;
                
                // Extract images from this paragraph
                const images = this.extractImagesFromHtml(p);
                currentSlide.images.push(...images);
            }
        }
        
        if (currentSlide) slides.push(currentSlide);
        
        return slides;
    }

    private parseByNestedList(html: string): SlideContent[] {
        const slides: SlideContent[] = [];
        
        // Find top-level list items (first level <li> inside <ol> or <ul>)
        // Use regex to extract top-level list items
        const topLevelListMatch = html.match(/<(ol|ul)>([\s\S]*?)<\/\1>/i);
        
        if (!topLevelListMatch) {
            return slides;
        }

        const listContent = topLevelListMatch[2];
        
        // Split by top-level <li> tags, but need to handle nested lists
        const items = this.splitTopLevelItems(listContent);
        
        for (const item of items) {
            // Extract title: text before first nested <ol> or <ul>, or entire text
            const titleMatch = item.match(/^([^<]+|<[^o][^>]*>)*?(?=<ol|<ul|$)/i);
            let title = '';
            
            if (titleMatch) {
                title = titleMatch[0].replace(/<[^>]+>/g, '').trim();
            }
            
            if (!title) {
                // Try to get first text content
                const textMatch = item.match(/>([^<]+)</);
                title = textMatch ? textMatch[1].trim() : 'Untitled';
            }

            // Extract bullets from nested lists
            const bullets: string[] = [];
            this.extractNestedBullets(item, bullets, 0);

            // Extract images
            const images = this.extractImagesFromHtml(item);

            slides.push({ title: title.substring(0, 100), bullets, images });
        }

        return slides;
    }

    private splitTopLevelItems(html: string): string[] {
        const items: string[] = [];
        let depth = 0;
        let currentItem = '';
        let i = 0;

        while (i < html.length) {
            if (html.substring(i, i + 4).toLowerCase() === '<li>') {
                if (depth === 0) {
                    if (currentItem.trim()) {
                        items.push(currentItem.trim());
                    }
                    currentItem = '';
                }
                depth++;
                currentItem += '<li>';
                i += 4;
            } else if (html.substring(i, i + 5).toLowerCase() === '</li>') {
                depth--;
                currentItem += '</li>';
                i += 5;
            } else {
                currentItem += html[i];
                i++;
            }
        }

        if (currentItem.trim()) {
            items.push(currentItem.trim());
        }

        return items;
    }

    private extractNestedBullets(html: string, bullets: string[], level: number): void {
        // Find all list items at current level
        const listMatch = html.match(/<(ol|ul)>([\s\S]*?)<\/\1>/gi);
        
        if (!listMatch) {
            return;
        }

        for (const list of listMatch) {
            const items = this.splitTopLevelItems(list.replace(/<\/?(ol|ul)>/gi, ''));
            
            for (const item of items) {
                // Get text content (excluding nested lists)
                const textContent = item
                    .replace(/<(ol|ul)>[\s\S]*<\/\1>/gi, '')
                    .replace(/<[^>]+>/g, '')
                    .trim();
                
                if (textContent && level < 3) { // Limit depth to avoid too many bullets
                    const indent = '  '.repeat(level);
                    bullets.push(indent + textContent);
                }

                // Recursively extract nested bullets
                const nestedListMatch = item.match(/<(ol|ul)>[\s\S]*<\/\1>/i);
                if (nestedListMatch) {
                    this.extractNestedBullets(nestedListMatch[0], bullets, level + 1);
                }
            }
        }
    }

    private extractBulletsFromHtml(html: string): string[] {
        const bullets: string[] = [];
        const liMatches = html.match(/<li>(.*?)<\/li>/gi);
        
        if (liMatches) {
            liMatches.forEach(li => {
                const text = li.replace(/<[^>]+>/g, '').trim();
                if (text) bullets.push(text);
            });
        }
        
        if (bullets.length === 0) {
            const plainText = html.replace(/<[^>]+>/g, '').trim();
            if (plainText) {
                bullets.push(plainText);
            }
        }

        return bullets;
    }

    private extractImagesFromHtml(html: string): string[] {
        const images: string[] = [];
        const imgMatches = html.match(/<img src="([^"]+)"/g);
        
        if (imgMatches) {
            imgMatches.forEach(match => {
                const src = match.match(/src="([^"]+)"/)?.[1];
                if (src) images.push(src);
            });
        }

        return images;
    }

    async parsePdf(filePath: string): Promise<DocumentData> {
        const dataBuffer = fs.readFileSync(filePath);
        const data = await (pdfParse as any)(dataBuffer);
        
        const sections = data.text.split(/\n{3,}/);
        
        const slides: SlideContent[] = sections
            .map((section: string) => section.trim())
            .filter((section: string) => section.length > 0)
            .map((section: string, index: number) => {
                const lines = section.split('\n').map((l: string) => l.trim()).filter((l: string) => l.length > 0);
                const title = lines[0] || `Slide ${index + 1}`;
                const bullets = lines.slice(1);
                
                return { title: title.substring(0, 100), bullets, images: [] };
            });

        return { title: 'PDF Presentation', slides };
    }
}