/**
 * 对比测试：旧方案（pptxgenjs）vs 新方案（HTML→PNG→PPT）
 * 使用相同的文档数据，分别生成 PPT 并评分
 */
import path from 'path';
import fs from 'fs';
import { ParserService } from './src/services/parser.service';
import { PPTService } from './src/services/ppt.service';
import { PPTImageService } from './src/services/ppt-image.service';
import { EvaluatorService } from './src/services/evaluator.service';
import { ChatService } from './src/services/chat.service';
import { DocumentData } from './src/types';
import dotenv from 'dotenv';

dotenv.config();

const outputDir = path.join(__dirname, 'output', 'comparison-test');

async function main() {
    fs.mkdirSync(outputDir, { recursive: true });

    const parserService = new ParserService();
    const pptService = new PPTService();
    const pptImageService = new PPTImageService();
    const evaluatorService = new EvaluatorService();
    const chatService = new ChatService();

    // Step 1: 获取文档数据
    console.log('=== Step 1: Parsing document ===');
    const docxPath = path.join(__dirname, 'input', '计算机发展史.docx');
    if (!fs.existsSync(docxPath)) {
        console.error('Test document not found:', docxPath);
        process.exit(1);
    }
    const parsedDoc = await parserService.parseDocx(docxPath);
    console.log(`Parsed: ${parsedDoc.title}, ${parsedDoc.slides.length} slides`);

    // Step 2: 用 LLM 生成 PPT 数据
    console.log('\n=== Step 2: Generating PPT data via LLM ===');
    
    // 构建 docContent
    let docContent = `\n\n=== 用户上传的文档内容 ===\n文档标题: ${parsedDoc.title}\n`;
    parsedDoc.slides.forEach((slide) => {
        docContent += `\n## ${slide.title}\n`;
        if (slide.bullets.length > 0) {
            docContent += slide.bullets.map(b => `- ${b}`).join('\n') + '\n';
        }
    });

    // 直接用 outline→confirmed 流程获取最终 PPT JSON
    const fakeOutlineHistory = [
        { role: 'user' as const, content: '根据这个文档生成科技风格的PPT' },
        { role: 'assistant' as const, content: '好的，大纲已生成。\n```outline\n标题：计算机发展史\n```' }
    ];
    
    const chatResponse = await chatService.chatAndGenerate(
        fakeOutlineHistory,
        '确认生成，请开始生成PPT',
        docContent
    );

    if (!chatResponse.pptData) {
        console.error('LLM did not return PPT data');
        console.log('Reply:', chatResponse.reply?.substring(0, 300));
        process.exit(1);
    }

    const pptData = chatResponse.pptData;
    console.log(`PPT data: ${pptData.title}, ${pptData.slides.length} slides`);

    // Step 3: 用旧方案生成
    console.log('\n=== Step 3: Generating PPT (Legacy: pptxgenjs) ===');
    const legacyPath = path.join(outputDir, 'legacy.pptx');
    const t1 = Date.now();
    await pptService.generate(pptData, legacyPath);
    const legacyTime = Date.now() - t1;
    console.log(`Legacy PPT generated in ${legacyTime}ms`);

    // Step 4: 用新方案生成
    console.log('\n=== Step 4: Generating PPT (New: HTML→PNG→PPT) ===');
    const htmlPath = path.join(outputDir, 'html-rendered.pptx');
    const t2 = Date.now();
    await pptImageService.generate(pptData, htmlPath);
    const htmlTime = Date.now() - t2;
    console.log(`HTML-rendered PPT generated in ${htmlTime}ms`);

    // Step 5: 评分对比
    console.log('\n=== Step 5: Evaluating both PPTs ===');
    
    const legacyReport = await evaluatorService.evaluate(pptData, legacyPath);
    const legacySaved = evaluatorService.saveReport(legacyReport, legacyPath);
    
    const htmlReport = await evaluatorService.evaluate(pptData, htmlPath);
    const htmlSaved = evaluatorService.saveReport(htmlReport, htmlPath);

    // Step 6: 输出对比结果
    console.log('\n' + '='.repeat(60));
    console.log('  COMPARISON RESULTS');
    console.log('='.repeat(60));
    
    console.log(`\n${'Metric'.padEnd(30)} ${'Legacy'.padStart(10)} ${'HTML→PNG'.padStart(10)} ${'Diff'.padStart(10)}`);
    console.log('-'.repeat(60));
    
    const printRow = (name: string, legacy: number, html: number) => {
        const diff = html - legacy;
        const sign = diff > 0 ? '+' : '';
        console.log(`${name.padEnd(30)} ${legacy.toFixed(1).padStart(10)} ${html.toFixed(1).padStart(10)} ${(sign + diff.toFixed(1)).padStart(10)}`);
    };

    printRow('Overall Score', legacyReport.overallScore, htmlReport.overallScore);
    printRow('Grade', 0, 0); // placeholder
    console.log(`${'Grade'.padEnd(30)} ${legacyReport.grade.padStart(10)} ${htmlReport.grade.padStart(10)}`);
    
    console.log('\n--- Dimension Scores ---');
    const dims = ['logic', 'layout', 'imageSemantics', 'contentRichness', 'audienceFit', 'consistency', 'sourceUnderstanding'] as const;
    for (const dim of dims) {
        printRow(
            dim,
            legacyReport.dimensions[dim].score,
            htmlReport.dimensions[dim].score
        );
    }

    console.log('\n--- Generation Time ---');
    console.log(`${'Legacy'.padEnd(30)} ${legacyTime}ms`);
    console.log(`${'HTML→PNG'.padEnd(30)} ${htmlTime}ms`);

    console.log('\n--- File Size ---');
    const legacySize = fs.statSync(legacyPath).size;
    const htmlSize = fs.statSync(htmlPath).size;
    console.log(`${'Legacy'.padEnd(30)} ${(legacySize / 1024).toFixed(0)} KB`);
    console.log(`${'HTML→PNG'.padEnd(30)} ${(htmlSize / 1024).toFixed(0)} KB`);

    console.log('\n--- Key Findings (Legacy) ---');
    legacyReport.keyFindings.slice(0, 5).forEach(f => console.log(`  • ${f}`));
    console.log('\n--- Key Findings (HTML→PNG) ---');
    htmlReport.keyFindings.slice(0, 5).forEach(f => console.log(`  • ${f}`));

    console.log('\nReports saved to:');
    console.log(`  Legacy: ${legacySaved.markdownPath}`);
    console.log(`  HTML:   ${htmlSaved.markdownPath}`);
    
    console.log('\n=== Comparison test complete ===');
}

main().catch(e => {
    console.error('Test failed:', e);
    process.exit(1);
});
