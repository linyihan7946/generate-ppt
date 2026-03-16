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
            slides = this.parseByNestedList(html);
        }
        // Fallback: treat as single slide
        else {
            slides = [{ 
                title: 'Document', 
                bullets: html.replace(/<[^>]+>/g, '\n').split('\n').filter(l => l.trim()), 
                images: [] 
            }];
        }

        // Extract title from first slide if exists
        const title = slides.length > 0 ? slides[0].title : 'Docx Presentation';

        return { title, slides };
    }

    private parseByHeadings(html: string): SlideContent[] {
        const slides: SlideContent[] = [];
        const sections = html.split(/<(h[1-6])>/i);
        
        for (let i = 1; i < sections.length; i += 2) {
            const tag = sections[i];
            const content = sections[i + 1];
            const titlePart = content.split(`</${tag}>`)[0];
            const title = titlePart.replace(/<[^>]+>/g, '');
            const rest = content.split(`</${tag}>`)[1] || '';
            
            const bullets = this.extractBulletsFromHtml(rest);
            const images = this.extractImagesFromHtml(rest);

            slides.push({ title, bullets, images });
        }

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