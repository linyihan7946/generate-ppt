import puppeteer, { Browser } from 'puppeteer';
import fs from 'fs';
import path from 'path';

/**
 * 使用 Puppeteer 将 HTML 字符串渲染为高清 PNG 图片
 * 分辨率：1920×1080 viewport + 2x deviceScaleFactor = 3840×2160 输出
 */
export class ScreenshotService {
    private browser: Browser | null = null;

    /**
     * 将多个 HTML 页面截图为 PNG base64 字符串数组
     */
    async captureSlides(htmlPages: string[], outputDir?: string): Promise<string[]> {
        const browser = await this.getBrowser();
        const results: string[] = [];

        // 创建临时输出目录
        const tempDir = outputDir || path.join(process.cwd(), 'output', '.slide-screenshots');
        fs.mkdirSync(tempDir, { recursive: true });

        const page = await browser.newPage();
        await page.setViewport({
            width: 1920,
            height: 1080,
            deviceScaleFactor: 2,
        });

        for (let i = 0; i < htmlPages.length; i++) {
            const html = htmlPages[i];
            await page.setContent(html, { waitUntil: 'domcontentloaded', timeout: 30000 });

            const screenshotPath = path.join(tempDir, `slide-${String(i).padStart(3, '0')}.png`);
            await page.screenshot({
                path: screenshotPath,
                type: 'png',
                fullPage: false,
                clip: { x: 0, y: 0, width: 1920, height: 1080 },
            });

            // 读取为 base64 供 pptxgenjs 使用
            const imageBuffer = fs.readFileSync(screenshotPath);
            const base64 = `data:image/png;base64,${imageBuffer.toString('base64')}`;
            results.push(base64);

            console.log(`Screenshot ${i + 1}/${htmlPages.length} captured`);
        }

        await page.close();
        return results;
    }

    private async getBrowser(): Promise<Browser> {
        if (!this.browser || !this.browser.connected) {
            this.browser = await puppeteer.launch({
                headless: true,
                args: [
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu',
                    '--font-render-hinting=none',
                ],
            });
        }
        return this.browser;
    }

    async close(): Promise<void> {
        if (this.browser) {
            await this.browser.close();
            this.browser = null;
        }
    }
}
