import axios from 'axios';
import { DocumentData } from '../types';

export interface ChatMessage {
    role: 'system' | 'user' | 'assistant';
    content: string;
}

export interface ChatResponse {
    reply: string;
    pptData?: DocumentData;
    outlineData?: OutlineData;
}

export interface OutlineData {
    title: string;
    brief: {
        deckGoal: string;
        audience: string;
        focus: string;
        style: string;
    };
    slides: Array<{
        title: string;
        slideRole: string;
        keyMessage: string;
        bullets: string[];
    }>;
}

export class ChatService {
    private apiKey: string | undefined;
    private baseUrl: string;

    constructor() {
        this.apiKey = process.env.PLANNER_AUTH_TOKEN || process.env.IMAGE_API_KEY; 
        this.baseUrl = process.env.PLANNER_API_BASE_URL || process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn';
    }

    async chatAndGenerate(messages: ChatMessage[], userText: string = '', docContent: string = ''): Promise<ChatResponse> {
        const phase = this.detectPhase(messages, userText, docContent);
        console.log('Chat phase detected:', phase);

        const systemContent = this.buildSystemPrompt(phase, docContent);

        let promptString: string;
        
        if (phase === 'outline') {
            // outline 阶段：构建聚焦 prompt，不传原始对话历史
            const requirementsSummary = this.extractRequirements(messages, userText, docContent);
            // 用 response priming 技巧：在 prompt 末尾预填回复开头，引导 LLM 进入输出模式
            promptString = `${systemContent}\n\n---\n以下是客户已提供的需求信息：\n${requirementsSummary}\n\n---\n现在请你根据以上信息，直接输出PPT大纲的JSON。\n\n好的，根据您的需求，我为您设计了以下PPT大纲：\n\n\`\`\`json\n`;
        } else {
            const systemPrompt: ChatMessage = { role: 'system', content: systemContent };
            const payloadMessages = [systemPrompt, ...messages];
            if (userText) {
                payloadMessages.push({ role: 'user', content: userText });
            }
            promptString = payloadMessages.map(m => `${m.role.toUpperCase()}:\n${m.content}`).join('\n\n') + '\n\nASSISTANT:\n';
        }

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

            return this.parseResponse(replyContent, phase);
        } catch (error: any) {
            console.error('LLM Chat Error:', error.response?.data || error.message);
            throw new Error('与AI助手通信失败，请稍后再试。');
        }
    }

    /**
     * 根据对话历史判断当前所处阶段
     * - gathering: 需求收集阶段
     * - outline: 应该输出大纲供用户确认
     * - confirmed: 用户已确认大纲，应输出最终 JSON
     */
    private detectPhase(messages: ChatMessage[], userText: string, docContent: string = ''): 'gathering' | 'outline' | 'confirmed' {
        // 用户明确确认大纲的信号
        const confirmPatterns = /确认生成|开始生成|就这样生成|可以生成|没问题.*生成|同意.*生成|好的.*生成|确认大纲|大纲没问题|大纲可以|就按这个|按这个生成/i;
        
        // 检查历史中是否已经出过大纲（assistant 消息中含有 ```outline 标记）
        const hasOutline = messages.some(m => 
            m.role === 'assistant' && m.content.includes('```outline')
        );

        if (hasOutline && confirmPatterns.test(userText)) {
            return 'confirmed';
        }

        // 用户主动请求生成大纲
        const generateNowPatterns = /生成.*大纲|出.*大纲|看看大纲|给我大纲|先出大纲/i;
        if (generateNowPatterns.test(userText)) {
            return 'outline';
        }

        // 历史中的 user 消息数 + 当前这条 userText（如果非空）
        const historyUserCount = messages.filter(m => m.role === 'user').length;
        const totalUserCount = historyUserCount + (userText ? 1 : 0);

        // 有文档上传时，第二轮就可以出大纲；无文档时需要至少2轮
        if (docContent && totalUserCount >= 1) {
            return 'outline';
        }
        if (totalUserCount >= 2) {
            return 'outline';
        }

        return 'gathering';
    }

    /**
     * 从对话历史中提取用户需求摘要（用于 outline 阶段的聚焦 prompt）
     */
    private extractRequirements(messages: ChatMessage[], userText: string, docContent: string): string {
        const parts: string[] = [];
        
        // 提取所有用户消息
        const userMessages = messages
            .filter(m => m.role === 'user')
            .map(m => m.content);
        if (userText) userMessages.push(userText);
        
        if (userMessages.length > 0) {
            parts.push('用户的原始需求：');
            userMessages.forEach((msg, i) => parts.push(`  ${i + 1}. ${msg}`));
        }
        
        if (docContent) {
            // 截取文档前2000字符防止过长
            const truncated = docContent.length > 2000 
                ? docContent.substring(0, 2000) + '\n...(文档内容已截断)'
                : docContent;
            parts.push(`\n用户上传的文档内容：\n${truncated}`);
        }
        
        return parts.join('\n');
    }

    private buildSystemPrompt(phase: 'gathering' | 'outline' | 'confirmed', docContent: string): string {
        const jsonSpec = `JSON 结构必须完全符合以下格式：
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
      "title": "幻灯片页面标题",
      "slideRole": "content",
      "keyMessage": "这一页的核心观点",
      "bullets": ["要点1：详细说明...", "要点2：详细说明...", "要点3：详细说明..."],
      "imagePrompt": "用英文描述这一页需要的配图内容，要具体、与本页主题强相关，例如：A futuristic microprocessor chip glowing with blue circuits on a dark motherboard"
    }
  ]
}
\`\`\`
slideRole 可以是 "agenda"(目录), "content"(正文), "comparison"(对比), "summary"(总结), "next_step"(下一步计划)。
imagePrompt 必须用英文撰写，具体描述与该页主题强相关的配图画面，避免通用描述，需要体现该页独特的内容场景。`;

        const docSection = docContent ? `\n\n=== 用户上传的文档内容 ===\n${docContent}\n===\n` : '';

        if (phase === 'gathering') {
            return `你是一个资深的商业演示文稿（PPT）咨询专家。你的任务是通过自然的对话，逐步了解客户的需求。
${docSection}
【当前阶段：需求收集】
你需要了解以下核心要素（分次自然地提问，每次只问1-2个问题，不要一次性全问）：
- PPT的用途：汇报、路演、培训、分享、教学等
- 目标受众：管理层、客户、技术人员、学生等
- 风格偏好：简洁、商务、科技、故事化、学术等
- 期望的PPT页数或信息密度
- 需要重点突出的内容或章节
${docContent ? '\n用户已上传了文档，你可以先简要概括文档内容，然后询问用户对PPT的具体要求（用途、受众、风格等）。' : ''}

【严格规则】
- 这个阶段绝对不要输出任何 \`\`\`json 或 \`\`\`outline 代码块
- 保持友好专业的对话语气
- 每次回复控制在3-5句话以内`;
        }

        if (phase === 'outline') {
            return `你的身份：资深商业PPT策划专家。
${docSection}
你的任务：根据下方客户需求，输出一份PPT结构大纲。直接输出，不要提问。

输出要求：在回复中包含一个 \`\`\`json 代码块，格式如下：

\`\`\`json
{
  "title": "演示文稿大标题",
  "brief": {
    "deckGoal": "这个PPT的目标",
    "audience": "目标受众",
    "focus": "核心焦点",
    "style": "视觉风格"
  },
  "slides": [
    {
      "title": "第1页标题",
      "slideRole": "agenda",
      "keyMessage": "核心观点",
      "bullets": ["要点1", "要点2", "要点3"],
      "imagePrompt": "English description of a scene closely related to this slide's topic"
    }
  ]
}
\`\`\`

slideRole 可选值：agenda(目录)、content(正文)、comparison(对比)、summary(总结)、next_step(下一步)。
imagePrompt 必须用英文，描述与该页内容直接相关的具体画面场景。

规则：
- 必须输出 \`\`\`json 代码块
- 不要向用户提问，直接根据已有信息生成
- 页数通常8-12页，每页至少3个bullets
- JSON之外可以加一句简短的说明文字`;
        }

        // phase === 'confirmed'
        return `你是一个资深的商业演示文稿（PPT）咨询专家。用户已经确认了大纲，现在需要生成最终的PPT数据。
${docSection}
【当前阶段：最终生成】
请根据之前对话中你输出的大纲和用户的反馈修改意见，生成最终版本的PPT数据。

在回复的最后，**必须**输出一段 JSON 格式的数据，用 \`\`\`json 和 \`\`\` 包裹起来。

${jsonSpec}

【严格规则】
- 必须输出 \`\`\`json 代码块
- JSON 内容要基于之前确认的大纲，bullets 内容要丰满专业
- 回复开头可以简短说一句"正在为您生成PPT..."之类的话
- 发挥你的专业知识，充实每一页的 bullets 内容
- 每一页都必须包含 imagePrompt 字段，用英文描述与该页内容密切相关的配图场景，要具体生动，避免抽象通用`;
    }

    private parseResponse(replyContent: string, phase: string): ChatResponse {
        console.log(`parseResponse phase=${phase}, has json=${replyContent.includes('```json')}, has outline=${replyContent.includes('```outline')}, response length=${replyContent.length}`);
        
        // outline 阶段可能用了 response priming，回复直接以 JSON 内容开始
        let processedReply = replyContent;
        if (phase === 'outline' && !replyContent.includes('```json') && !replyContent.includes('```outline')) {
            // 尝试检测是否是裸 JSON（可能以 { 开头或含有 JSON 结构）
            const trimmed = replyContent.trim();
            if (trimmed.startsWith('{')) {
                // 回复直接就是 JSON（priming 把 ```json\n 已经加到 prompt 中了）
                const closingIndex = trimmed.lastIndexOf('}');
                if (closingIndex > 0) {
                    const jsonPart = trimmed.substring(0, closingIndex + 1);
                    const afterJson = trimmed.substring(closingIndex + 1).replace(/^[\s\n]*```[\s\n]*/, '').trim();
                    processedReply = '```json\n' + jsonPart + '\n```\n' + afterJson;
                }
            }
        }
        
        // 检查是否包含 JSON 块
        const jsonMatch = processedReply.match(/```json\n([\s\S]*?)\n```/);
        if (jsonMatch && jsonMatch[1]) {
            try {
                const parsed = JSON.parse(jsonMatch[1]);
                
                // 大纲阶段：即使 LLM 返回了 json，也转为 outlineData 而非生成 PPT
                if (phase === 'outline' && parsed.slides) {
                    console.log('Outline phase: converting json to outlineData');
                    const outlineData: OutlineData = {
                        title: parsed.title || '演示文稿',
                        brief: parsed.brief || { deckGoal: '', audience: '', focus: '', style: '' },
                        slides: parsed.slides.map((s: any) => ({
                            title: s.title || '',
                            slideRole: s.slideRole || 'content',
                            keyMessage: s.keyMessage || '',
                            bullets: s.bullets || [],
                        })),
                    };
                    const finalReply = processedReply.replace(/```json\n[\s\S]*?\n```/, '').trim()
                        || '以下是为您生成的PPT大纲，请确认是否满意：';
                    return { reply: finalReply, outlineData };
                }
                
                // confirmed 阶段：正常生成 PPT
                const pptData = parsed as DocumentData;
                if (pptData && pptData.slides) {
                    pptData.slides = pptData.slides.map(slide => ({
                        ...slide,
                        images: slide.images || [],
                        bullets: slide.bullets || [],
                        imagePrompt: slide.imagePrompt || '',
                    }));
                }
                const finalReply = processedReply.replace(/```json\n[\s\S]*?\n```/, '').trim()
                    || '正在为您生成PPT，请稍候...';
                return { reply: finalReply, pptData };
            } catch (e) {
                console.error("解析模型返回的 JSON 失败:", e);
            }
        }

        // 检查是否包含大纲预览
        const outlineMatch = processedReply.match(/```outline\n([\s\S]*?)\n```/);
        if (outlineMatch && outlineMatch[1]) {
            const outlineData = this.parseOutline(outlineMatch[1]);
            const finalReply = processedReply.replace(/```outline\n[\s\S]*?\n```/, '').trim();
            return { reply: finalReply, outlineData };
        }

        // outline 阶段但 LLM 未返回任何结构化数据：添加提示信息
        if (phase === 'outline') {
            return { reply: processedReply + '\n\n_(系统提示：大纲正在生成中，如果内容不完整请再次发送消息)_' };
        }

        return { reply: processedReply };
    }

    private parseOutline(outlineText: string): OutlineData {
        const lines = outlineText.split('\n');
        let title = '';
        let deckGoal = '';
        let audience = '';
        let style = '';
        const slides: OutlineData['slides'] = [];
        let currentSlide: OutlineData['slides'][0] | null = null;

        for (const line of lines) {
            const trimmed = line.trim();
            if (trimmed.startsWith('标题：') || trimmed.startsWith('标题:')) {
                title = trimmed.replace(/^标题[：:]/, '').trim();
            } else if (trimmed.startsWith('目标：') || trimmed.startsWith('目标:')) {
                deckGoal = trimmed.replace(/^目标[：:]/, '').trim();
            } else if (trimmed.startsWith('受众：') || trimmed.startsWith('受众:')) {
                audience = trimmed.replace(/^受众[：:]/, '').trim();
            } else if (trimmed.startsWith('风格：') || trimmed.startsWith('风格:')) {
                style = trimmed.replace(/^风格[：:]/, '').trim();
            } else if (/^第\d+页/.test(trimmed)) {
                if (currentSlide) slides.push(currentSlide);
                const slideTitle = trimmed.replace(/^第\d+页\s*/, '').replace(/\(.*?\)/, '').trim();
                const roleMatch = trimmed.match(/\(类型:\s*(.*?)\)/);
                const roleMap: Record<string, string> = {
                    '目录': 'agenda', '正文': 'content', '对比': 'comparison',
                    '总结': 'summary', '下一步': 'next_step', '时间线': 'timeline',
                };
                currentSlide = {
                    title: slideTitle,
                    slideRole: roleMap[roleMatch?.[1]?.trim() || ''] || 'content',
                    keyMessage: '',
                    bullets: [],
                };
            } else if (trimmed.startsWith('核心观点：') || trimmed.startsWith('核心观点:')) {
                if (currentSlide) {
                    currentSlide.keyMessage = trimmed.replace(/^核心观点[：:]/, '').trim();
                }
            } else if (trimmed.startsWith('- ')) {
                if (currentSlide) {
                    currentSlide.bullets.push(trimmed.substring(2).trim());
                }
            }
        }
        if (currentSlide) slides.push(currentSlide);

        return {
            title: title || '演示文稿',
            brief: { deckGoal, audience, focus: '', style },
            slides,
        };
    }
}