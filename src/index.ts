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
import { PPTImageService } from './services/ppt-image.service';

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
const pptImageService = new PPTImageService();

// 文档原始图片的会话级缓存（按上传时的文件名哈希存储，10分钟过期自动清理）
interface ImageCacheEntry {
    titleMap: Map<string, string[]>;  // slideTitle -> images[]
    ordered: string[];                // 按顺序全部图片
    createdAt: number;
}
const docImageCache = new Map<string, ImageCacheEntry>();
const IMAGE_CACHE_TTL = 10 * 60 * 1000; // 10 minutes

function cleanExpiredImageCache() {
    const now = Date.now();
    for (const [key, entry] of docImageCache) {
        if (now - entry.createdAt > IMAGE_CACHE_TTL) {
            docImageCache.delete(key);
        }
    }
}

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
        // 保存 parser 解析出的文档原始图片，用于后续回填到 LLM 生成的 slides 中
        const docOriginalImages: Map<string, string[]> = new Map(); // slideTitle -> images[]
        const docOrderedImages: string[] = []; // 按顺序收集所有图片
        // 用于缓存的 key（文档文件名）
        let imageCacheKey = '';
        
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
                        const dataUrl = `data:${mimeType};base64,${base64}`;
                        docContent += `\n\n[用户上传了图片: ${file.originalname}]`;
                        docOrderedImages.push(dataUrl);
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
                        // 收集文档中的原始图片
                        if (slide.images && slide.images.length > 0) {
                            docOriginalImages.set(slide.title.trim().toLowerCase(), [...slide.images]);
                            docOrderedImages.push(...slide.images);
                        }
                    });
                });
                if (docOrderedImages.length > 0) {
                    console.log(`Extracted ${docOrderedImages.length} original images from uploaded documents`);
                }
                // 缓存图片供后续请求（确认生成阶段）使用
                if (docOrderedImages.length > 0) {
                    imageCacheKey = primaryDoc.title.trim().toLowerCase();
                    cleanExpiredImageCache();
                    docImageCache.set(imageCacheKey, {
                        titleMap: new Map(docOriginalImages),
                        ordered: [...docOrderedImages],
                        createdAt: Date.now(),
                    });
                    console.log(`Cached ${docOrderedImages.length} images under key: "${imageCacheKey}"`);
                }
            }
        }
        
        // 如果当前请求没有上传文件（确认生成阶段），尝试从缓存恢复图片
        if (docOrderedImages.length === 0 && docImageCache.size > 0) {
            // 从聊天历史中提取可能的文档标题
            const allText = messages.map(m => m.content).join(' ') + ' ' + text;
            for (const [key, entry] of docImageCache) {
                if (Date.now() - entry.createdAt > IMAGE_CACHE_TTL) continue;
                // 模糊匹配：缓存 key 出现在对话内容中，或者取最近的缓存
                if (allText.toLowerCase().includes(key) || docImageCache.size === 1) {
                    for (const [title, imgs] of entry.titleMap) {
                        docOriginalImages.set(title, imgs);
                    }
                    docOrderedImages.push(...entry.ordered);
                    console.log(`Restored ${entry.ordered.length} cached images from key: "${key}"`);
                    break;
                }
            }
        }

        const chatResponse = await chatService.chatAndGenerate(messages, text, docContent);

        let downloadUrl = undefined;

        if (chatResponse.pptData) {
            console.log('LLM generated PPT data, processing PPTX...');

            // 回填文档原始图片到 LLM 生成的 slides
            if (docOrderedImages.length > 0 && chatResponse.pptData.slides) {
                let backfilledCount = 0;
                const slides = chatResponse.pptData.slides;
                
                // 策略1: 按标题匹配回填
                for (const slide of slides) {
                    if (slide.images && slide.images.length > 0) continue;
                    const titleKey = slide.title.trim().toLowerCase();
                    const matched = docOriginalImages.get(titleKey);
                    if (matched && matched.length > 0) {
                        slide.images = [...matched];
                        backfilledCount += matched.length;
                    }
                }
                
                // 策略2: 未匹配到的空 slide，按顺序轮流分配剩余图片
                const usedImages = new Set(slides.flatMap(s => s.images || []));
                const unusedImages = docOrderedImages.filter(img => !usedImages.has(img));
                if (unusedImages.length > 0) {
                    let imgIdx = 0;
                    for (const slide of slides) {
                        if (slide.images && slide.images.length > 0) continue;
                        if (imgIdx >= unusedImages.length) break;
                        slide.images = [unusedImages[imgIdx]];
                        imgIdx++;
                        backfilledCount++;
                    }
                }
                
                if (backfilledCount > 0) {
                    console.log(`Backfilled ${backfilledCount} original document images into slides`);
                }
            }
            
            const outputFilename = `presentation-chat-${Date.now()}.pptx`;
            const outputDir = path.join(__dirname, '../output');
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }
            const outputPath = path.join(outputDir, outputFilename);

            const useHtmlMode = process.env.PPT_RENDER_MODE === 'html';
                        
            if (useHtmlMode) {
                console.log('Using HTML\u2192PNG\u2192PPT rendering pipeline...');
                const enableAiImages = process.env.ENABLE_AI_IMAGES !== 'false';
                const imageConcurrency = Number(process.env.IMAGE_CONCURRENCY || 2);
                if (enableAiImages) {
                    console.log('Generating AI images for HTML slides...');
                    await imageService.enrichSlidesWithGeneratedImages(chatResponse.pptData.slides, imageConcurrency);
                }
                await pptImageService.generate(chatResponse.pptData, outputPath);
            } else {
                console.log('Using native pptxgenjs rendering pipeline...');
                const enableAiImages = process.env.ENABLE_AI_IMAGES !== 'false';
                const imageConcurrency = Number(process.env.IMAGE_CONCURRENCY || 2);
                if (enableAiImages) {
                    await imageService.enrichSlidesWithGeneratedImages(chatResponse.pptData.slides, imageConcurrency);
                }
                await pptService.generate(chatResponse.pptData, outputPath);
            }
            
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
        const useHtmlMode = process.env.PPT_RENDER_MODE === 'html';
        if (useHtmlMode) {
            console.log('Using HTML\u2192PNG\u2192PPT rendering pipeline...');
            await pptImageService.generate(docData, outputPath);
        } else {
            console.log('Using native pptxgenjs rendering pipeline...');
            await pptService.generate(docData, outputPath);
        }

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
