import fs from 'fs';
import path from 'path';
import dotenv from 'dotenv';

import { ParserService } from '../src/services/parser.service';
import { ImageService } from '../src/services/image.service';
import { PPTService } from '../src/services/ppt.service';
import { PlannerService } from '../src/services/planner.service';
import { EvaluatorService } from '../src/services/evaluator.service';
import {
    DeckAudience,
    DeckFocus,
    DeckFormat,
    DeckLength,
    DeckStyle,
    DocumentData,
    PlannerMode,
    PlannerOptions,
    QualityReport,
} from '../src/types';

dotenv.config({ path: path.resolve(__dirname, '../.env') });

type SupportedInputExt = '.md' | '.markdown' | '.docx' | '.pdf';

interface BatchResult {
    inputPath: string;
    outputPath?: string;
    reportJsonPath?: string;
    reportMarkdownPath?: string;
    status: 'success' | 'failed';
    title?: string;
    slideCount?: number;
    score?: number;
    grade?: string;
    durationMs: number;
    error?: string;
}

interface BatchSummary {
    generatedAt: string;
    inputDir: string;
    outputDir: string;
    totalFiles: number;
    successCount: number;
    failedCount: number;
    averageScore: number;
    useImages: boolean;
    plannerOptions: PlannerOptions;
    results: BatchResult[];
}

const SUPPORTED_EXTENSIONS = new Set<SupportedInputExt>(['.md', '.markdown', '.docx', '.pdf']);
const VALUE_FLAGS = new Set([
    '--input-dir',
    '--output-dir',
    '--planner-mode',
    '--deck-format',
    '--audience',
    '--focus',
    '--style',
    '--length',
    '--with-images',
]);

function getArgValue(flag: string): string | undefined {
    const index = process.argv.findIndex((arg) => arg === flag);
    if (index < 0 || index + 1 >= process.argv.length) {
        return undefined;
    }
    return process.argv[index + 1];
}

function hasArg(flag: string): boolean {
    return process.argv.includes(flag);
}

function getPositionalArgs(): string[] {
    const args = process.argv.slice(2);
    const positional: string[] = [];

    for (let i = 0; i < args.length; i += 1) {
        const current = args[i];
        if (VALUE_FLAGS.has(current)) {
            i += 1;
            continue;
        }
        if (current.startsWith('--')) {
            continue;
        }
        positional.push(current);
    }

    return positional;
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

function parseBooleanArg(flag: string, defaultValue: boolean): boolean {
    const value = getArgValue(flag);
    if (!value) {
        return defaultValue;
    }

    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === '1' || normalized === 'yes') {
        return true;
    }
    if (normalized === 'false' || normalized === '0' || normalized === 'no') {
        return false;
    }

    throw new Error(`Invalid value for ${flag}: ${value}. Use true or false.`);
}

function formatTimestamp(date = new Date()): string {
    const year = String(date.getFullYear());
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}${month}${day}-${hours}${minutes}${seconds}`;
}

function walkSupportedDocuments(rootDir: string): string[] {
    const results: string[] = [];

    function visit(currentDir: string): void {
        const entries = fs.readdirSync(currentDir, { withFileTypes: true });
        entries.forEach((entry) => {
            const fullPath = path.join(currentDir, entry.name);
            if (entry.isDirectory()) {
                visit(fullPath);
                return;
            }

            const ext = path.extname(entry.name).toLowerCase() as SupportedInputExt;
            if (SUPPORTED_EXTENSIONS.has(ext)) {
                results.push(fullPath);
            }
        });
    }

    visit(rootDir);
    return results.sort((left, right) => left.localeCompare(right, 'zh-CN'));
}

async function parseDocument(parserService: ParserService, inputPath: string): Promise<DocumentData> {
    const ext = path.extname(inputPath).toLowerCase() as SupportedInputExt;

    if (ext === '.md' || ext === '.markdown') {
        return parserService.parseMarkdown(inputPath);
    }
    if (ext === '.docx') {
        return parserService.parseDocx(inputPath);
    }
    if (ext === '.pdf') {
        return parserService.parsePdf(inputPath);
    }

    throw new Error(`Unsupported file type: ${ext}`);
}

function toOutputPptPath(inputDir: string, outputDir: string, inputPath: string): string {
    const relativePath = path.relative(inputDir, inputPath);
    const ext = path.extname(relativePath);
    const relativeWithoutExt = relativePath.slice(0, relativePath.length - ext.length);
    const outputPath = path.join(outputDir, `${relativeWithoutExt}.pptx`);
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    return outputPath;
}

function toMarkdown(summary: BatchSummary): string {
    const lines: string[] = [
        '# Batch PPT Test Summary',
        '',
        `- Generated At: ${summary.generatedAt}`,
        `- Input Dir: ${summary.inputDir}`,
        `- Output Dir: ${summary.outputDir}`,
        `- Total Files: ${summary.totalFiles}`,
        `- Success Count: ${summary.successCount}`,
        `- Failed Count: ${summary.failedCount}`,
        `- Average Score: **${summary.averageScore}**`,
        `- Use Images: ${summary.useImages}`,
        '',
        '## Results',
        '',
        '| File | Status | Score | Grade | Slides | Duration(s) |',
        '|---|---|---:|---:|---:|---:|',
    ];

    summary.results.forEach((result) => {
        lines.push(
            `| ${path.relative(summary.inputDir, result.inputPath)} | ${result.status} | ${result.score ?? '-'} | ${result.grade ?? '-'} | ${result.slideCount ?? '-'} | ${(result.durationMs / 1000).toFixed(1)} |`,
        );
    });

    const failed = summary.results.filter((result) => result.status === 'failed');
    if (failed.length > 0) {
        lines.push('', '## Failed Files', '');
        failed.forEach((result) => {
            lines.push(`- ${path.relative(summary.inputDir, result.inputPath)}: ${result.error || 'Unknown error'}`);
        });
    }

    lines.push('');
    return lines.join('\n');
}

function printSummary(summary: BatchSummary): void {
    console.log('');
    console.log('Batch PPT Results');
    console.log('=================');

    summary.results.forEach((result, index) => {
        const relativePath = path.relative(summary.inputDir, result.inputPath);
        if (result.status === 'success') {
            console.log(
                `${index + 1}. ${relativePath} -> ${result.score} (${result.grade}) | slides=${result.slideCount} | ${(result.durationMs / 1000).toFixed(1)}s`,
            );
        } else {
            console.log(`${index + 1}. ${relativePath} -> FAILED | ${(result.durationMs / 1000).toFixed(1)}s`);
            console.log(`   ${result.error}`);
        }
    });

    console.log('');
    console.log(`Average Score: ${summary.averageScore}`);
    console.log(`Succeeded: ${summary.successCount}/${summary.totalFiles}`);
    console.log(`Failed: ${summary.failedCount}/${summary.totalFiles}`);
}

async function run(): Promise<void> {
    const positionalArgs = getPositionalArgs();
    const plannerModeArg = getArgValue('--planner-mode');
    const deckFormatArg = getArgValue('--deck-format');
    const audienceArg = getArgValue('--audience');
    const focusArg = getArgValue('--focus');
    const styleArg = getArgValue('--style');
    const lengthArg = getArgValue('--length');

    const plannerMode = normalizePlannerMode(plannerModeArg);
    const deckFormat = normalizeDeckFormat(deckFormatArg);
    const audience = normalizeAudience(audienceArg);
    const focus = normalizeFocus(focusArg);
    const style = normalizeStyle(styleArg);
    const length = normalizeLength(lengthArg);

    if (plannerModeArg && !plannerMode) {
        throw new Error(`Invalid --planner-mode value: ${plannerModeArg}. Use strict or creative.`);
    }
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

    const plannerOptions: PlannerOptions = {
        mode: plannerMode,
        deckFormat,
        audience,
        focus,
        style,
        length,
    };

    const inputDir = path.resolve(process.cwd(), getArgValue('--input-dir') || positionalArgs[0] || 'input');
    if (!fs.existsSync(inputDir)) {
        throw new Error(`Input directory not found: ${inputDir}`);
    }

    const outputDir = path.resolve(
        process.cwd(),
        getArgValue('--output-dir') || positionalArgs[1] || path.join('output', `batch-${formatTimestamp()}`),
    );
    fs.mkdirSync(outputDir, { recursive: true });

    const useImages = parseBooleanArg('--with-images', process.env.ENABLE_AI_IMAGES !== 'false');
    const failFast = hasArg('--fail-fast');
    const imageConcurrency = Number(process.env.IMAGE_CONCURRENCY || 2);

    const inputFiles = walkSupportedDocuments(inputDir);
    if (inputFiles.length === 0) {
        throw new Error(`No supported files found in ${inputDir}. Supported: ${Array.from(SUPPORTED_EXTENSIONS).join(', ')}`);
    }

    const parserService = new ParserService();
    const plannerService = new PlannerService();
    const imageService = new ImageService();
    const pptService = new PPTService();
    const evaluatorService = new EvaluatorService();

    const results: BatchResult[] = [];

    for (const inputPath of inputFiles) {
        const startedAt = Date.now();
        const outputPath = toOutputPptPath(inputDir, outputDir, inputPath);
        const relativePath = path.relative(inputDir, inputPath);

        console.log(`Processing: ${relativePath}`);

        try {
            let docData = await parseDocument(parserService, inputPath);
            docData = await plannerService.planDocument(docData, plannerOptions);

            if (useImages) {
                await imageService.enrichSlidesWithGeneratedImages(docData.slides, imageConcurrency);
            }

            await pptService.generate(docData, outputPath);
            const report: QualityReport = await evaluatorService.evaluate(docData, outputPath);
            const reportPaths = evaluatorService.saveReport(report, outputPath);

            results.push({
                inputPath,
                outputPath,
                reportJsonPath: reportPaths.jsonPath,
                reportMarkdownPath: reportPaths.markdownPath,
                status: 'success',
                title: report.title,
                slideCount: docData.slides.length,
                score: report.overallScore,
                grade: report.grade,
                durationMs: Date.now() - startedAt,
            });

            console.log(`  -> Score: ${report.overallScore} (${report.grade})`);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            results.push({
                inputPath,
                outputPath,
                status: 'failed',
                durationMs: Date.now() - startedAt,
                error: message,
            });

            console.error(`  -> Failed: ${message}`);

            if (failFast) {
                break;
            }
        }
    }

    const succeeded = results.filter((result) => result.status === 'success' && typeof result.score === 'number');
    const averageScore =
        succeeded.length === 0
            ? 0
            : Math.round((succeeded.reduce((sum, result) => sum + (result.score || 0), 0) / succeeded.length) * 10) / 10;

    const summary: BatchSummary = {
        generatedAt: new Date().toISOString(),
        inputDir,
        outputDir,
        totalFiles: results.length,
        successCount: succeeded.length,
        failedCount: results.length - succeeded.length,
        averageScore,
        useImages,
        plannerOptions,
        results,
    };

    const summaryJsonPath = path.join(outputDir, 'batch-summary.json');
    const summaryMarkdownPath = path.join(outputDir, 'batch-summary.md');

    fs.writeFileSync(summaryJsonPath, JSON.stringify(summary, null, 2), 'utf-8');
    fs.writeFileSync(summaryMarkdownPath, toMarkdown(summary), 'utf-8');

    printSummary(summary);
    console.log(`Summary JSON: ${summaryJsonPath}`);
    console.log(`Summary Markdown: ${summaryMarkdownPath}`);
}

run().catch((error) => {
    console.error('Batch generation failed:', error instanceof Error ? error.message : error);
    process.exit(1);
});
