import axios from 'axios';

async function test() {
    console.log('Testing Phase 2 only...');
    
    // Simulate Phase 1 reply
    const fakeP1Reply = '您好！我需要了解一些信息。';
    
    const r = await axios.post('http://localhost:3000/api/chat', {
        messages: [
            { role: 'user', content: '我想做一个关于AI技术的培训PPT' },
            { role: 'assistant', content: fakeP1Reply }
        ],
        text: '受众是技术团队，风格偏科技感，大概10页左右'
    }, { headers: { 'Content-Type': 'application/json' }, timeout: 60000 });
    
    console.log('Full response reply:');
    console.log(r.data.reply);
    console.log('\n---');
    console.log('Has outline:', !!r.data.outlineData);
    console.log('Has download:', !!r.data.downloadUrl);
    if (r.data.outlineData) {
        console.log('Outline:', JSON.stringify(r.data.outlineData, null, 2).substring(0, 500));
    }
}

test().catch(e => console.error('Error:', e.response?.data || e.message));
