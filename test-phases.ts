import axios from 'axios';
import FormData from 'form-data';
import fs from 'fs';
import path from 'path';

const BASE_URL = 'http://localhost:3000';

async function testPhases() {
    console.log('=== Phase 1: First message (gathering) ===');
    const r1 = await axios.post(`${BASE_URL}/api/chat`, {
        messages: [],
        text: '我想做一个关于AI技术的培训PPT'
    }, { headers: { 'Content-Type': 'application/json' } });
    console.log('Reply:', r1.data.reply?.substring(0, 200));
    console.log('Has outline:', !!r1.data.outlineData);
    console.log('Has download:', !!r1.data.downloadUrl);
    console.log('Phase 1 PASS:', !r1.data.outlineData && !r1.data.downloadUrl ? '✅' : '⚠️ unexpected data');

    console.log('\n=== Phase 2: Second message → should produce outline ===');
    const msgs2 = [
        { role: 'user', content: '我想做一个关于AI技术的培训PPT' },
        { role: 'assistant', content: r1.data.reply }
    ];
    const r2 = await axios.post(`${BASE_URL}/api/chat`, {
        messages: msgs2,
        text: '受众是技术团队，风格偏科技感，大概10页左右'
    }, { headers: { 'Content-Type': 'application/json' } });
    console.log('Reply:', r2.data.reply?.substring(0, 200));
    console.log('Has outline:', !!r2.data.outlineData);
    console.log('Has download:', !!r2.data.downloadUrl);
    if (r2.data.outlineData) {
        console.log('Outline title:', r2.data.outlineData.title);
        console.log('Outline slides:', r2.data.outlineData.slides?.length);
    }
    console.log('Phase 2 PASS:', r2.data.outlineData && !r2.data.downloadUrl ? '✅' : '⚠️');

    console.log('\n=== Phase 3: Confirm outline → should generate PPT ===');
    const msgs3 = [
        ...msgs2,
        { role: 'user', content: '受众是技术团队，风格偏科技感，大概10页左右' },
        { role: 'assistant', content: r2.data.reply + '\n```outline\n标题：AI Training\n```' }
    ];
    const r3 = await axios.post(`${BASE_URL}/api/chat`, {
        messages: msgs3,
        text: '确认生成，请开始生成PPT'
    }, { headers: { 'Content-Type': 'application/json' }, timeout: 180000 });
    console.log('Reply:', r3.data.reply?.substring(0, 200));
    console.log('Has outline:', !!r3.data.outlineData);
    console.log('Has download:', !!r3.data.downloadUrl);
    console.log('Phase 3 PASS:', r3.data.downloadUrl ? '✅' : '⚠️ no downloadUrl');

    console.log('\n=== Phase 4: File upload first message → should ask questions, not generate ===');
    const formData = new FormData();
    formData.append('text', '根据这个文档生成科技风格的PPT');
    formData.append('messages', JSON.stringify([]));
    const docxPath = path.join(__dirname, 'input', '计算机发展史.docx');
    if (fs.existsSync(docxPath)) {
        formData.append('files', fs.createReadStream(docxPath));
        const r4 = await axios.post(`${BASE_URL}/api/chat`, formData, {
            headers: formData.getHeaders(),
            timeout: 120000
        });
        console.log('Reply:', r4.data.reply?.substring(0, 200));
        console.log('Has outline:', !!r4.data.outlineData);
        console.log('Has download:', !!r4.data.downloadUrl);
        console.log('Phase 4 PASS:', !r4.data.downloadUrl ? '✅ no premature generation' : '⚠️ generated too early!');
    }

    console.log('\n=== All phase tests complete ===');
}

testPhases().catch(e => console.error('Error:', e.response?.data || e.message));
