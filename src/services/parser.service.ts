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
        
        const slides: SlideContent[] = [];
        const sections = html.split(/<(h[1-6])>/i);
        
        for (let i = 1; i < sections.length; i += 2) {
            const tag = sections[i];
            const content = sections[i + 1];
            const titlePart = content.split(`</${tag}>`)[0];
            const title = titlePart.replace(/<[^>]+>/g, '');
            const rest = content.split(`</${tag}>`)[1] || '';
            
            const bullets = rest.match(/<li>(.*?)<\/li>/g)?.map(li => li.replace(/<[^>]+>/g, '')) || [];
            if (bullets.length === 0 && rest.trim()) {
                const plainText = rest.replace(/<[^>]+>/g, '').trim();
                if (plainText) bullets.push(plainText);
            }

            const sectionImages: string[] = [];
            const imgMatches = rest.match(/<img src="([^"]+)"/g);
            if (imgMatches) {
                imgMatches.forEach(match => {
                    const src = match.match(/src="([^"]+)"/)?.[1];
                    if (src) sectionImages.push(src);
                });
            }

            slides.push({ title, bullets, images: sectionImages });
        }

        if (slides.length === 0) {
            slides.push({ title: 'Untitled', bullets: html.replace(/<[^>]+>/g, '\n').split('\n').filter(l => l.trim()), images: [] });
        }

        return { title: 'Docx Presentation', slides };
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
