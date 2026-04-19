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

import { ChatService, ChatMessage } from './services/chat.service';

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));
app.use('/output', express.static(path.join(__dirname, '../output')));

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
const chatService = new ChatService();

// 新增对话生成 PPT 接口
app.post('/api/chat', upload.array('files', 5), async (req: any, res: any) => {
    try {
        console.log(' req.files:', req.files);
        console.log(' req.body:', JSON.stringify(req.body).substring(0, 200));
        
        const files: Express.Multer.File[] = req.files || [];
        const text = req.body.text || '';
        const messagesRaw = req.body.messages;
        
        let messages: ChatMessage[] = [];
        if (Array.isArray(messagesRaw)) {
            messages = messagesRaw;
        } else if (typeof messagesRaw === 'string') {
            try {
                messages = JSON.parse(messagesRaw);
            } catch {
                messages = [];
            }
        }

        console.log('Received chat request, messages count:', messages.length, ', files count:', files?.length || 0);

        let docContent = '';
        
        if (files && files.length > 0) {
            console.log('Processing uploaded files...');
            const parsedDocs: DocumentData[] = [];
            
            for (const file of files) {
                const ext = path.extname(file.originalname).toLowerCase();
                console.log(`Parsing file: ${file.originalname}, ext: ${ext}`);
                
                try {
                    if (ext === '.md') {
                        const doc = await parserService.parseMarkdown(file.path);
                        parsedDocs.push(doc);
                    } else if (ext === '.docx') {
                        const doc = await parserService.parseDocx(file.path);
                        parsedDocs.push(doc);
                    } else if (ext === '.pdf') {
                        const doc = await parserService.parsePdf(file.path);
                        parsedDocs.push(doc);
                    } else if (['.png', '.jpg', '.jpeg'].includes(ext)) {
                        const base64 = fs.readFileSync(file.path).toString('base64');
                        const mimeType = ext === '.png' ? 'image/png' : 'image/jpeg';
                        docContent += `\n\n[用户上传了图片: ${file.originalname}]\n图片数据: data:${mimeType};base64,${base64.substring(0, 100)}...`;
                    }
                    
                    // Clean up temp file
                    fs.unlinkSync(file.path);
                } catch (parseErr) {
                    console.error(`Error parsing file ${file.originalname}:`, parseErr);
                }
            }
            
            if (parsedDocs.length > 0) {
                const primaryDoc = parsedDocs[0];
                docContent = `\n\n=== 用户上传的文档内容 ===\n文档标题: ${primaryDoc.title}\n`;
                parsedDocs.forEach((doc, idx) => {
                    if (idx > 0) docContent += `\n--- 文档 ${idx + 1} ---\n`;
                    docContent += `标题: ${doc.title}\n`;
                    doc.slides.forEach((slide, slideIdx) => {
                        docContent += `\n## ${slide.title}\n`;
                        if (slide.bullets.length > 0) {
                            docContent += slide.bullets.map(b => `- ${b}`).join('\n') + '\n';
                        }
                    });
                });
            }
        }

        const chatResponse = await chatService.chatAndGenerate(messages, text, docContent);

        let downloadUrl = undefined;

        if (chatResponse.pptData) {
            console.log('LLM generated PPT data, processing images and PPTX...');
            
            const enableAiImages = process.env.ENABLE_AI_IMAGES !== 'false';
            const imageConcurrency = Number(process.env.IMAGE_CONCURRENCY || 2);
            if (enableAiImages) {
                await imageService.enrichSlidesWithGeneratedImages(chatResponse.pptData.slides, imageConcurrency);
            }

            const outputFilename = `presentation-chat-${Date.now()}.pptx`;
            const outputDir = path.join(__dirname, '../output');
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }
            const outputPath = path.join(outputDir, outputFilename);
            
            await pptService.generate(chatResponse.pptData, outputPath);
            
            downloadUrl = `/output/${outputFilename}`;
        }

        res.json({
            reply: chatResponse.reply,
            downloadUrl: downloadUrl,
            outlineData: chatResponse.outlineData || undefined,
        });

    } catch (error: any) {
        console.error('Chat API Error:', error);
        res.status(500).json({ error: error.message || 'An error occurred during chat.' });
    }
});

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
