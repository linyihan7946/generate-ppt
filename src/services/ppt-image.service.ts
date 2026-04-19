import pptxgen from 'pptxgenjs';
import { DocumentData } from '../types';
import { SlideRendererService } from './slide-renderer.service';
import { ScreenshotService } from './screenshot.service';

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;

/**
 * 方案A: HTML → 高清PNG → PPT
 * 每页幻灯片先渲染为 HTML，用 Puppeteer 截图为高清 PNG，
 * 再将图片作为全页背景放入 PPT
 */
export class PPTImageService {
    private renderer = new SlideRendererService();
    private screenshotService = new ScreenshotService();

    async generate(data: DocumentData, outputPath: string): Promise<string> {
        console.log('PPTImageService: Rendering HTML slides...');
        const htmlPages = this.renderer.renderAll(data);
        console.log(`PPTImageService: ${htmlPages.length} HTML pages generated`);

        console.log('PPTImageService: Capturing screenshots...');
        const screenshots = await this.screenshotService.captureSlides(htmlPages);
        console.log(`PPTImageService: ${screenshots.length} screenshots captured`);

        console.log('PPTImageService: Building PPTX...');
        const pres = new pptxgen();
        pres.layout = 'LAYOUT_WIDE';
        pres.author = 'generate-ppt';
        pres.company = 'generate-ppt';
        pres.subject = 'Generated presentation (HTML-rendered)';
        pres.title = data.title;

        for (const base64 of screenshots) {
            const slide = pres.addSlide();
            slide.addImage({
                data: base64,
                x: 0,
                y: 0,
                w: SLIDE_WIDTH,
                h: SLIDE_HEIGHT,
            });
        }

        await pres.writeFile({ fileName: outputPath });
        await this.screenshotService.close();

        console.log(`PPTImageService: PPTX written to ${outputPath}`);
        return outputPath;
    }
}
