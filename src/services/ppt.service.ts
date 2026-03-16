import pptxgen from 'pptxgenjs';
import { DocumentData } from '../types';
import * as fs from 'fs';
import path from 'path';

export class PPTService {
    async generate(data: DocumentData, outputPath: string): Promise<string> {
        const pres = new pptxgen();

        // 1. Title Slide
        const titleSlide = pres.addSlide();
        titleSlide.addText(data.title, {
            x: '10%',
            y: '40%',
            w: '80%',
            fontSize: 44,
            align: 'center',
            color: '363636',
            bold: true,
        });

        // 2. Content Slides
        for (const slideData of data.slides) {
            const slide = pres.addSlide();
            
            // Add Title
            slide.addText(slideData.title, {
                x: 0.5,
                y: 0.5,
                w: '90%',
                h: 1,
                fontSize: 32,
                color: '0088CC',
                bold: true,
            });

            // Add Bullets
            const bulletList = slideData.bullets.map(b => ({ text: b, options: { bullet: true, fontSize: 18 } }));
            slide.addText(bulletList, {
                x: 0.5,
                y: 1.5,
                w: slideData.images.length > 0 ? '50%' : '90%',
                h: 4,
                align: 'left',
                valign: 'top',
            });

            // Add Images if present
            if (slideData.images.length > 0) {
                // For simplicity, take the first image
                const img = slideData.images[0];
                if (img.startsWith('http')) {
                    slide.addImage({ path: img, x: '55%', y: 1.5, w: 4, h: 3 });
                } else if (img.startsWith('data:image')) {
                    slide.addImage({ data: img, x: '55%', y: 1.5, w: 4, h: 3 });
                }
            }
        }

        await pres.writeFile({ fileName: outputPath });
        return outputPath;
    }
}
