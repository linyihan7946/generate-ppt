import axios from 'axios';
import FormData from 'form-data';
import fs from 'fs';
import path from 'path';

const BASE_URL = 'http://localhost:3000';

async function testChatWithFile() {
    console.log('=== Test 1: Chat API Health Check ===');
    try {
        const response = await axios.get(BASE_URL);
        console.log('✅ Server is running:', response.status === 200);
    } catch (err) {
        console.error('❌ Server is not running:', err.message);
        console.log('Please start the server with: npm start');
        return;
    }

    console.log('\n=== Test 2: Chat with text message ===');
    try {
        const response = await axios.post(`${BASE_URL}/api/chat`, {
            messages: [
                { role: 'user', content: '我想做一个关于AI技术的工作汇报PPT' }
            ]
        }, {
            headers: { 'Content-Type': 'application/json' }
        });
        
        console.log('✅ Response received:', response.data.reply ? 'YES' : 'NO');
        console.log('Reply preview:', response.data.reply?.substring(0, 100) + '...');
    } catch (err) {
        console.error('❌ Chat with text failed:', err.response?.data || err.message);
    }

    console.log('\n=== Test 3: Chat with file upload (DOCX) ===');
    try {
        const formData = new FormData();
        formData.append('text', '请根据这份文档生成PPT');
        formData.append('messages', JSON.stringify([]));
        
        const docxPath = path.join(__dirname, 'input', '计算机发展史.docx');
        if (fs.existsSync(docxPath)) {
            formData.append('files', fs.createReadStream(docxPath), {
                filename: '计算机发展史.docx',
                contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            
            const response = await axios.post(`${BASE_URL}/api/chat`, formData, {
                headers: formData.getHeaders()
            });
            
            console.log('✅ File uploaded successfully');
            console.log('Reply:', response.data.reply?.substring(0, 150) + '...');
            console.log('Download URL:', response.data.downloadUrl || 'NOT_PROVIDED');
        } else {
            console.log('⚠️ Test file not found:', docxPath);
        }
    } catch (err) {
        console.error('❌ Chat with file failed:', err.response?.data || err.message);
    }

    console.log('\n=== Test 4: Chat with Markdown file ===');
    try {
        const formData = new FormData();
        formData.append('text', '请根据这份文档生成PPT');
        formData.append('messages', JSON.stringify([]));
        
        const mdPath = path.join(__dirname, 'test.md');
        if (fs.existsSync(mdPath)) {
            formData.append('files', fs.createReadStream(mdPath), {
                filename: 'test.md',
                contentType: 'text/markdown'
            });
            
            const response = await axios.post(`${BASE_URL}/api/chat`, formData, {
                headers: formData.getHeaders()
            });
            
            console.log('✅ Markdown file uploaded successfully');
            console.log('Reply:', response.data.reply?.substring(0, 150) + '...');
            console.log('Download URL:', response.data.downloadUrl || 'NOT_PROVIDED');
        } else {
            console.log('⚠️ Test file not found:', mdPath);
        }
    } catch (err) {
        console.error('❌ Chat with Markdown file failed:', err.response?.data || err.message);
    }

    console.log('\n=== Test 5: Generate PPT endpoint (direct file upload) ===');
    try {
        const formData = new FormData();
        const docxPath = path.join(__dirname, 'input', '计算机发展史.docx');
        
        if (fs.existsSync(docxPath)) {
            formData.append('file', fs.createReadStream(docxPath), {
                filename: '计算机发展史.docx',
                contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            formData.append('plannerMode', 'creative');
            
            const response = await axios.post(`${BASE_URL}/generate-ppt`, formData, {
                headers: formData.getHeaders(),
                responseType: 'arraybuffer'
            });
            
            console.log('✅ PPT generated successfully');
            console.log('Content-Type:', response.headers['content-type']);
            console.log('Content-Length:', response.data.byteLength, 'bytes');
            
            const outputPath = path.join(__dirname, 'test-output.pptx');
            fs.writeFileSync(outputPath, Buffer.from(response.data));
            console.log('✅ PPT saved to:', outputPath);
        }
    } catch (err) {
        console.error('❌ Generate PPT failed:', err.response?.data || err.message);
    }

    console.log('\n=== All tests completed ===');
}

testChatWithFile();