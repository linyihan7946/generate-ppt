require('dotenv').config();
const axios = require('axios');

async function test() {
    const baseUrl = process.env.PLANNER_API_BASE_URL || 'https://www.aigenimage.cn';
    const apiKey = process.env.PLANNER_AUTH_TOKEN;
    console.log('Testing with baseUrl:', baseUrl);
    
    try {
        const response = await axios.post(
            `${baseUrl}/api/llm/direct`,
            {
                model: process.env.PLANNER_MODEL || 'gemini-3.1-pro-preview',
                prompt: 'USER:\nHello\n\nASSISTANT:\n',
                temperature: 0.7,
                stream: false
            },
            {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${apiKey}`
                }
            }
        );
        console.log('Success:', response.data);
    } catch (err) {
        console.error('Error:', err.response ? err.response.data : err.message);
    }
}
test();