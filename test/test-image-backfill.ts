/**
 * 测试文档原始图片回填功能
 * 
 * 测试流程:
 * 1. 创建带图片引用的 MD 文件
 * 2. 上传到 /api/chat (outline阶段)
 * 3. 确认生成 (confirmed阶段)
 * 4. 检查生成的 PPTX 是否包含图片
 */
import axios from 'axios';
import FormData from 'form-data';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';

const BASE_URL = 'http://localhost:3000';

// 创建一个带有 base64 内联图片的测试 Markdown
function createTestMarkdownWithImages(): string {
    // 1x1 红色像素 PNG
    const redPixel = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==';
    // 1x1 蓝色像素 PNG  
    const bluePixel = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==';

    return `# 人工智能技术概览

## 机器学习基础
- 监督学习是最常见的机器学习范式
- 通过标注数据训练模型识别模式
- 分类和回归是两大核心任务

![机器学习示意图](data:image/png;base64,${redPixel})

## 深度学习与神经网络
- 多层神经网络能自动提取特征
- CNN在图像领域表现卓越
- RNN处理序列数据效果显著

![深度学习架构](data:image/png;base64,${bluePixel})

## 自然语言处理
- Transformer架构革新了NLP
- 大语言模型展现通用智能
- 文本生成和理解能力大幅提升

## 计算机视觉
- 图像分类准确率已超过人类
- 目标检测实时性不断提升
- 图像生成技术日新月异
`;
}

async function runTest() {
    console.log('====================================');
    console.log('  文档图片回填功能测试');
    console.log('====================================\n');

    // 0. 健康检查
    try {
        await axios.get(BASE_URL);
        console.log('✅ 服务器运行正常\n');
    } catch {
        console.error('❌ 服务器未启动，请先 npm start');
        process.exit(1);
    }

    // 1. 创建测试 MD 文件
    const testMdPath = path.join(__dirname, 'test-image-backfill.md');
    const mdContent = createTestMarkdownWithImages();
    fs.writeFileSync(testMdPath, mdContent, 'utf-8');
    console.log('✅ 测试 Markdown 文件已创建（含 2 张内联图片）\n');

    // 2. Phase 1: 上传文件 → 获取大纲
    console.log('--- Phase 1: 上传文档获取大纲 ---');
    let outlineReply = '';
    let outlineData: any = null;

    const form1 = new FormData();
    form1.append('text', '根据这份文档帮我生成PPT');
    form1.append('messages', JSON.stringify([]));
    form1.append('files', fs.createReadStream(testMdPath), {
        filename: 'test-image-backfill.md',
        contentType: 'text/markdown',
    });

    try {
        const res1 = await axios.post(`${BASE_URL}/api/chat`, form1, {
            headers: form1.getHeaders(),
            timeout: 120000,
        });
        outlineReply = res1.data.reply || '';
        outlineData = res1.data.outlineData;
        console.log('✅ 大纲阶段返回成功');
        console.log(`  reply: ${outlineReply.substring(0, 80)}...`);
        console.log(`  outlineData: ${outlineData ? `${outlineData.slides?.length} slides` : 'null'}`);
    } catch (err: any) {
        console.error('❌ 大纲阶段失败:', err.response?.data || err.message);
        cleanup(testMdPath);
        process.exit(1);
    }

    // 3. Phase 2: 确认生成 PPT（不重新上传文件，测试缓存回填）
    console.log('\n--- Phase 2: 确认生成 PPT ---');
    
    const messages = [
        { role: 'user', content: '根据这份文档帮我生成PPT' },
        { role: 'assistant', content: outlineReply + (outlineData ? `\n\`\`\`outline\n${formatOutline(outlineData)}\n\`\`\`` : '') },
    ];

    let downloadUrl = '';
    try {
        const res2 = await axios.post(`${BASE_URL}/api/chat`, {
            text: '确认生成，请开始生成PPT',
            messages,
        }, {
            headers: { 'Content-Type': 'application/json' },
            timeout: 300000, // AI图片生成可能较慢
        });
        downloadUrl = res2.data.downloadUrl || '';
        console.log('✅ PPT 生成成功');
        console.log(`  reply: ${(res2.data.reply || '').substring(0, 80)}...`);
        console.log(`  downloadUrl: ${downloadUrl}`);
    } catch (err: any) {
        console.error('❌ PPT 生成失败:', err.response?.data || err.message);
        cleanup(testMdPath);
        process.exit(1);
    }

    // 4. 下载并检查 PPTX
    if (!downloadUrl) {
        console.error('\n❌ 无下载链接，无法验证图片');
        cleanup(testMdPath);
        process.exit(1);
    }

    console.log('\n--- Phase 3: 验证 PPTX 图片 ---');
    try {
        const pptRes = await axios.get(`${BASE_URL}${downloadUrl}`, { responseType: 'arraybuffer' });
        const pptxBuffer = Buffer.from(pptRes.data);
        console.log(`  PPTX 大小: ${(pptxBuffer.length / 1024).toFixed(1)} KB`);
        
        const zip = await JSZip.loadAsync(pptxBuffer);
        
        // 检查 ppt/media 目录中的图片数量
        const mediaFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/media/'));
        console.log(`  PPTX 内嵌图片数: ${mediaFiles.length}`);
        mediaFiles.forEach(f => {
            console.log(`    - ${path.basename(f)}`);
        });

        // 检查 slide 的 rels 文件中引用的图片
        const slideRels = Object.keys(zip.files).filter(name => 
            /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(name)
        );
        let slidesWithImages = 0;
        for (const relFile of slideRels) {
            const relContent = await zip.files[relFile].async('string');
            if (relContent.includes('/media/')) {
                slidesWithImages++;
            }
        }
        console.log(`  有图片引用的幻灯片数: ${slidesWithImages}/${slideRels.length}`);

        // 判定测试结果
        console.log('\n====================================');
        if (mediaFiles.length > 0) {
            console.log('✅ 测试通过 — PPTX 包含图片！');
            console.log(`  共 ${mediaFiles.length} 张图片嵌入到 ${slidesWithImages} 张幻灯片`);
        } else {
            console.log('❌ 测试失败 — PPTX 中没有发现图片');
        }
        console.log('====================================');

    } catch (err: any) {
        console.error('❌ PPTX 验证失败:', err.message);
    }

    cleanup(testMdPath);
}

function formatOutline(data: any): string {
    let text = `标题：${data.title}\n`;
    if (data.slides) {
        data.slides.forEach((s: any, i: number) => {
            text += `${i + 1}. ${s.title}\n`;
            if (s.bullets) {
                s.bullets.forEach((b: string) => { text += `   - ${b}\n`; });
            }
        });
    }
    return text;
}

function cleanup(testMdPath: string) {
    try { fs.unlinkSync(testMdPath); } catch {}
}

runTest().catch(err => {
    console.error('测试异常:', err);
    process.exit(1);
});
