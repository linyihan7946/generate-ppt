import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import dotenv from 'dotenv';
import cors from 'cors';

import { ParserService } from './services/parser.service';
import { PPTService } from './services/ppt.service';
import { ImageService } from './services/image.service';
import { DocumentData } from './types';

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
const pptService = new PPTService();
const imageService = new ImageService();

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

        // 2. Optional: Generate AI Images for slides that have none
        console.log('Generating AI images (if enabled)...');
        for (const slide of docData.slides) {
            if (slide.images.length === 0) {
                const aiImg = await imageService.generateImage(slide.title);
                if (aiImg) {
                    slide.images.push(aiImg);
                }
            }
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

        // 4. Send File
        res.download(outputPath, (err) => {
            if (err) {
                console.error('Error sending file:', err);
            }
            // Cleanup: remove uploaded and generated files
            // fs.unlinkSync(filePath);
            // fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error('Error generating PPT:', error);
        res.status(500).send('An error occurred during PPT generation.');
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
