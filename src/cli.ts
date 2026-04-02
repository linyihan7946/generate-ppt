import path from 'path';
import fs from 'fs';
import dotenv from 'dotenv';

import { ParserService } from './services/parser.service';
import { ImageService } from './services/image.service';
import { PPTService } from './services/ppt.service';
import { DocumentData } from './types';

dotenv.config();

function getArgValue(flag: string): string | undefined {
    const index = process.argv.findIndex((arg) => arg === flag);
    if (index < 0 || index + 1 >= process.argv.length) {
        return undefined;
    }
    return process.argv[index + 1];
}

async function run(): Promise<void> {
    const inputArg = getArgValue('--input');
    if (!inputArg) {
        throw new Error('Missing --input argument. Example: --input input/计算机发展史.docx');
    }

    const inputPath = path.resolve(process.cwd(), inputArg);
    if (!fs.existsSync(inputPath)) {
        throw new Error(`Input file not found: ${inputPath}`);
    }

    const ext = path.extname(inputPath).toLowerCase();
    const parserService = new ParserService();
    const imageService = new ImageService();
    const pptService = new PPTService();

    let docData: DocumentData;
    if (ext === '.md') {
        docData = await parserService.parseMarkdown(inputPath);
    } else if (ext === '.docx') {
        docData = await parserService.parseDocx(inputPath);
    } else if (ext === '.pdf') {
        docData = await parserService.parsePdf(inputPath);
    } else {
        throw new Error(`Unsupported file type: ${ext}`);
    }

    const enableAiImages = process.env.ENABLE_AI_IMAGES !== 'false';
    const imageConcurrency = Number(process.env.IMAGE_CONCURRENCY || 2);
    if (enableAiImages) {
        await imageService.enrichSlidesWithGeneratedImages(docData.slides, imageConcurrency);
    }

    const outputArg = getArgValue('--output');
    const outputDir = path.resolve(process.cwd(), 'output');
    fs.mkdirSync(outputDir, { recursive: true });

    const defaultOutputName = `${path.basename(inputPath, ext)}-${Date.now()}.pptx`;
    const outputPath = outputArg
        ? path.resolve(process.cwd(), outputArg)
        : path.join(outputDir, defaultOutputName);

    await pptService.generate(docData, outputPath);

    console.log(`Generated PPT: ${outputPath}`);
    console.log(`Title: ${docData.title}`);
    console.log(`Slides: ${docData.slides.length}`);
}

run().catch((error) => {
    console.error('CLI generation failed:', error instanceof Error ? error.message : error);
    process.exit(1);
});
