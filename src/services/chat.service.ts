import axios from 'axios';
import { DocumentData } from '../types';

export interface ChatMessage {
    role: 'system' | 'user' | 'assistant';
    content: string;
}

export interface ChatResponse {
    reply: string;
    pptData?: DocumentData; // 如果大模型决定生成PPT，则返回解析后的数据
}

export class ChatService {
    private apiKey: string | undefined;
    private baseUrl: string;

    constructor() {
        this.apiKey = process.env.PLANNER_AUTH_TOKEN || process.env.IMAGE_API_KEY; 
        this.baseUrl = process.env.PLANNER_API_BASE_URL || process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn';
    }

    async chatAndGenerate(messages: ChatMessage[]): Promise<ChatResponse> {
        const systemPrompt: ChatMessage = {
            role: 'system',
            content: `你是一个资深的商业演示文稿（PPT）咨询专家。你的任务是通过对话了解客户的需求，并最终为他们生成一份高质量的PPT内容。

请按照以下流程与用户交互：
1. **引导提问**：如果用户提供的信息不足，你需要主动引导用户，了解以下核心要素（可以分次自然地提问，不要像机器一样一次性抛出所有问题）：
   - 这个PPT的用途是？（如：汇报、路演、培训）
   - 受众是谁？（如：管理层、客户、技术人员）
   - 风格偏好？（如：简洁、商务、故事化）
   - 信息密度？（如：精简大纲、详细解释）
   - 具体要讲述的核心业务或主题内容是什么？
2. **确认并生成**：当你认为收集到了足够的信息，或者用户明确要求“现在开始生成PPT”时，你需要根据收集到的所有信息，构思一份完整的PPT大纲和内容。
3. **输出格式**：在你回复的话术最后，**必须**输出一段 JSON 格式的数据，用 \`\`\`json 和 \`\`\` 包裹起来。后端系统会自动解析这段 JSON 去生成真实的 PPT 文件。

JSON 结构必须完全符合以下格式：
\`\`\`json
{
  "title": "演示文稿的大标题",
  "brief": {
    "deckGoal": "一句话描述这个PPT的最终目标",
    "audience": "目标受众",
    "focus": "核心焦点",
    "style": "视觉和演讲风格"
  },
  "slides": [
    {
      "title": "幻灯片页面标题 (如：项目背景)",
      "slideRole": "content", 
      "keyMessage": "这一页要传达的核心观点或金句",
      "bullets": ["要点1：详细说明...", "要点2：详细说明...", "要点3：详细说明..."]
    }
  ]
}
\`\`\`
注意：slideRole 可以是 "agenda"(目录), "content"(正文), "comparison"(对比), "summary"(总结), "next_step"(下一步计划)。
请发挥你的专业知识，为用户填充丰满、专业的 bullets 内容。如果还没聊完，不要输出 JSON，继续聊天。`
        };

        const payloadMessages = [systemPrompt, ...messages];
        const promptString = payloadMessages.map(m => `${m.role.toUpperCase()}:\n${m.content}`).join('\n\n') + '\n\nASSISTANT:\n';

        try {
            const response = await axios.post(
                `${this.baseUrl}/api/llm/direct`,
                {
                    model: process.env.PLANNER_MODEL || 'gemini-3.1-pro-preview',
                    prompt: promptString,
                    temperature: 0.7,
                    stream: false
                },
                {
                    headers: {
                        'Content-Type': 'application/json',
                        ...(this.apiKey && { Authorization: `Bearer ${this.apiKey}` }),
                    },
                    timeout: 60000,
                    validateStatus: () => true,
                }
            );

            if (response.status !== 200 || response.data?.success === false) {
                console.error(`LLM API failed: status=${response.status}, message=${response.data?.message || 'unknown'}`);
                throw new Error('API return error');
            }

            let replyContent = '';
            const payload = response.data;
            if (typeof payload.data === 'string') replyContent = payload.data;
            else if (payload.data?.choices?.[0]?.text) replyContent = payload.data.choices[0].text;
            else if (payload.data?.choices?.[0]?.message?.content) replyContent = payload.data.choices[0].message.content;
            else if (payload.data?.reply) replyContent = String(payload.data.reply);
            else if (payload.data?.text) replyContent = String(payload.data.text);
            else if (payload.data?.content) replyContent = String(payload.data.content);
            else replyContent = JSON.stringify(payload);
            
            // 尝试从回复中解析 JSON
            const jsonMatch = replyContent.match(/```json\n([\s\S]*?)\n```/);
            let pptData: DocumentData | undefined = undefined;
            let finalReply = replyContent;

            if (jsonMatch && jsonMatch[1]) {
                try {
                    pptData = JSON.parse(jsonMatch[1]);
                    // 从回复中剔除 json，只保留大模型的寒暄文本给用户看
                    finalReply = replyContent.replace(/```json\n[\s\S]*?\n```/, '').trim();
                    if (!finalReply) {
                        finalReply = "我已经为您生成了PPT内容，正在处理成文件，请稍候下载...";
                    }
                } catch (e) {
                    console.error("解析模型返回的 JSON 失败:", e);
                }
            }

            return {
                reply: finalReply,
                pptData
            };
        } catch (error: any) {
            console.error('LLM Chat Error:', error.response?.data || error.message);
            throw new Error('与AI助手通信失败，请稍后再试。');
        }
    }
}