import './polyfills';
import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import dotenv from 'dotenv';
import cors from 'cors';

import { ParserService } from './services/parser.service';
import { PPTService } from './services/ppt.service';
import { ImageService } from './services/image.service';
import { PlannerService } from './services/planner.service';
import { EvaluatorService } from './services/evaluator.service';
import { DeckAudience, DeckFocus, DeckFormat, DeckLength, DeckStyle, DocumentData, PlannerMode } from './types';

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Set up storage for uploaded files
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = path.join(__dirname, 'uploads');
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}`);
    },
});

const upload = multer({ storage });

const parserService = new ParserService();
const plannerService = new PlannerService();
const pptService = new PPTService();
const imageService = new ImageService();
const evaluatorService = new EvaluatorService();

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

app.post('/generate-ppt', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).send('No file uploaded.');
        }

        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();

        let docData: DocumentData;

        // 1. Parse Document
        console.log(`Parsing ${ext} file...`);
        if (ext === '.md') {
            docData = await parserService.parseMarkdown(filePath);
        } else if (ext === '.docx') {
            docData = await parserService.parseDocx(filePath);
        } else if (ext === '.pdf') {
            docData = await parserService.parsePdf(filePath);
        } else {
            return res.status(400).send('Unsupported file format.');
        }

        const plannerModeArg = typeof req.body?.plannerMode === 'string' ? req.body.plannerMode : undefined;
        const deckFormatArg = typeof req.body?.deckFormat === 'string' ? req.body.deckFormat : undefined;
        const audienceArg = typeof req.body?.audience === 'string' ? req.body.audience : undefined;
        const focusArg = typeof req.body?.focus === 'string' ? req.body.focus : undefined;
        const styleArg = typeof req.body?.style === 'string' ? req.body.style : undefined;
        const lengthArg = typeof req.body?.length === 'string' ? req.body.length : undefined;

        const plannerMode = normalizePlannerMode(plannerModeArg);
        const deckFormat = normalizeDeckFormat(deckFormatArg);
        const audience = normalizeAudience(audienceArg);
        const focus = normalizeFocus(focusArg);
        const style = normalizeStyle(styleArg);
        const length = normalizeLength(lengthArg);

        if (plannerModeArg && !plannerMode) {
            return res.status(400).send('Invalid plannerMode. Use strict or creative.');
        }
        if (deckFormatArg && !deckFormat) {
            return res.status(400).send('Invalid deckFormat. Use presenter or detailed.');
        }
        if (audienceArg && !audience) {
            return res.status(400).send('Invalid audience.');
        }
        if (focusArg && !focus) {
            return res.status(400).send('Invalid focus.');
        }
        if (styleArg && !style) {
            return res.status(400).send('Invalid style.');
        }
        if (lengthArg && !length) {
            return res.status(400).send('Invalid length.');
        }

        console.log('Planning presentation...');
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
            console.log('Generating AI images (if enabled)...');
            await imageService.enrichSlidesWithGeneratedImages(docData.slides, imageConcurrency);
        }

        // 3. Generate PPT
        const outputFilename = `presentation-${Date.now()}.pptx`;
        
        // Use the output directory in the project root
        const outputDir = path.join(__dirname, '../output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        const outputPath = path.join(outputDir, outputFilename);
        
        console.log('Generating PPT...');
        await pptService.generate(docData, outputPath);

        const enableEvaluation = process.env.ENABLE_EVALUATION !== 'false';
        if (enableEvaluation) {
            const report = await evaluatorService.evaluate(docData, outputPath);
            const reportPaths = evaluatorService.saveReport(report, outputPath);
            res.setHeader('X-PPT-Quality-Score', String(report.overallScore));
            res.setHeader('X-PPT-Quality-Grade', report.grade);
            res.setHeader('X-PPT-Quality-Json', reportPaths.jsonPath);
            res.setHeader('X-PPT-Quality-Markdown', reportPaths.markdownPath);
        }

        res.download(outputPath, (err) => {
            if (err) {
                console.error('Error sending file:', err);
            }
        });

    } catch (error) {
        console.error('Error generating PPT:', error);
        res.status(500).send('An error occurred during PPT generation.');
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
