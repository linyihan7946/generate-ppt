import path from 'path';
import fs from 'fs';
import dotenv from 'dotenv';

import { ParserService } from './services/parser.service';
import { ImageService } from './services/image.service';
import { PPTService } from './services/ppt.service';
import { PlannerService } from './services/planner.service';
import { EvaluatorService } from './services/evaluator.service';
import { DeckAudience, DeckFocus, DeckFormat, DeckLength, DeckStyle, DocumentData, PlannerMode } from './types';

dotenv.config();

function getArgValue(flag: string): string | undefined {
    const index = process.argv.findIndex((arg) => arg === flag);
    if (index < 0 || index + 1 >= process.argv.length) {
        return undefined;
    }
    return process.argv[index + 1];
}

function normalizePlannerMode(input?: string): PlannerMode | undefined {
    if (input === 'strict' || input === 'creative') {
        return input;
    }

    return undefined;
}

function normalizeDeckFormat(input?: string): DeckFormat | undefined {
    if (input === 'presenter' || input === 'detailed') {
        return input;
    }
    return undefined;
}

function normalizeAudience(input?: string): DeckAudience | undefined {
    if (input === 'general' || input === 'beginner' || input === 'executive' || input === 'student' || input === 'technical') {
        return input;
    }
    return undefined;
}

function normalizeFocus(input?: string): DeckFocus | undefined {
    if (input === 'overview' || input === 'timeline' || input === 'argument' || input === 'process' || input === 'comparison') {
        return input;
    }
    return undefined;
}

function normalizeStyle(input?: string): DeckStyle | undefined {
    if (input === 'professional' || input === 'minimal' || input === 'bold' || input === 'educational') {
        return input;
    }
    return undefined;
}

function normalizeLength(input?: string): DeckLength | undefined {
    if (input === 'short' || input === 'default' || input === 'long') {
        return input;
    }
    return undefined;
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
    const plannerService = new PlannerService();
    const imageService = new ImageService();
    const pptService = new PPTService();
    const evaluatorService = new EvaluatorService();

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

    const plannerModeArg = getArgValue('--planner-mode');
    const plannerMode = normalizePlannerMode(plannerModeArg);
    if (plannerModeArg && !plannerMode) {
        throw new Error(`Invalid --planner-mode value: ${plannerModeArg}. Use strict or creative.`);
    }
    const deckFormatArg = getArgValue('--deck-format');
    const audienceArg = getArgValue('--audience');
    const focusArg = getArgValue('--focus');
    const styleArg = getArgValue('--style');
    const lengthArg = getArgValue('--length');

    const deckFormat = normalizeDeckFormat(deckFormatArg);
    const audience = normalizeAudience(audienceArg);
    const focus = normalizeFocus(focusArg);
    const style = normalizeStyle(styleArg);
    const length = normalizeLength(lengthArg);

    if (deckFormatArg && !deckFormat) {
        throw new Error(`Invalid --deck-format value: ${deckFormatArg}. Use presenter or detailed.`);
    }
    if (audienceArg && !audience) {
        throw new Error(`Invalid --audience value: ${audienceArg}. Use general, beginner, executive, student, or technical.`);
    }
    if (focusArg && !focus) {
        throw new Error(`Invalid --focus value: ${focusArg}. Use overview, timeline, argument, process, or comparison.`);
    }
    if (styleArg && !style) {
        throw new Error(`Invalid --style value: ${styleArg}. Use professional, minimal, bold, or educational.`);
    }
    if (lengthArg && !length) {
        throw new Error(`Invalid --length value: ${lengthArg}. Use short, default, or long.`);
    }

    docData = await plannerService.planDocument(docData, {
        mode: plannerMode,
        deckFormat,
        audience,
        focus,
        style,
        length,
    });

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

    const enableEvaluation = process.env.ENABLE_EVALUATION !== 'false';
    if (enableEvaluation) {
        const report = await evaluatorService.evaluate(docData, outputPath);
        const reportPaths = evaluatorService.saveReport(report, outputPath);
        console.log(`Quality Score: ${report.overallScore} (${report.grade})`);
        console.log(`Quality JSON: ${reportPaths.jsonPath}`);
        console.log(`Quality Markdown: ${reportPaths.markdownPath}`);
    }

    console.log(`Generated PPT: ${outputPath}`);
    console.log(`Title: ${docData.title}`);
    console.log(`Slides: ${docData.slides.length}`);
    if (docData.brief) {
        console.log(`Deck Format: ${docData.brief.deckFormat}`);
        console.log(`Audience: ${docData.brief.audience}`);
        console.log(`Focus: ${docData.brief.focus}`);
    }
}

run().catch((error) => {
    console.error('CLI generation failed:', error instanceof Error ? error.message : error);
    process.exit(1);
});
