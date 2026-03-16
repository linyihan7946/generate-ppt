# Document to PPT Generator

根据 Markdown、DOCX、PDF 文档自动生成 PowerPoint 演示文稿的后端服务。

## 功能特性

- ✅ 支持 Markdown (.md)、Word (.docx)、PDF (.pdf) 文件格式
- ✅ 自动提取文档层级结构（标题 -> 幻灯片标题）
- ✅ 保留原始文档顺序和层级
- ✅ 提取文档中的配图（支持 DOCX 内嵌图片）
- ✅ 支持 AI 自动生成配图（需配置 OpenAI API Key）
- ✅ Web 界面上传文件

## 快速开始

### 1. 安装依赖

```bash
npm install
```

### 2. 配置环境变量

复制 `.env` 文件并配置：

```env
OPENAI_API_KEY=your_openai_api_key_here
PORT=3000
```

> 如果不配置 `OPENAI_API_KEY`，AI 配图功能将被跳过。

### 3. 启动服务

```bash
# 开发模式
npm start

# 生产模式
npm run build
npm run serve
```

服务将在 http://localhost:3000 启动。

## API 使用

### POST `/generate-ppt`

上传文档并生成 PPT。

**请求：**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: 
  - `file`: 文档文件（支持 .md, .docx, .pdf）

**响应：**
- Content-Type: `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- 返回生成的 .pptx 文件

**示例 (cURL)：**

```bash
curl -X POST http://localhost:3000/generate-ppt \
  -F "file=@/path/to/document.docx" \
  --output presentation.pptx
```

**示例 (JavaScript)：**

```javascript
const formData = new FormData();
formData.append('file', fileInput.files[0]);

const response = await fetch('/generate-ppt', {
    method: 'POST',
    body: formData
});

const blob = await response.blob();
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'presentation.pptx';
a.click();
```

## Web 界面

访问 http://localhost:3000 可使用 Web 界面上传文件。

## 文档解析规则

### Markdown
- `# 标题` → 幻灯片标题
- 列表项 → 幻灯片要点
- 段落 → 幻灯片要点
- 图片链接 → 幻灯片配图

### DOCX
- 标题样式（Heading 1-6）→ 幻灯片标题
- 正文段落 → 幻灯片要点
- 列表项 → 幻灯片要点
- 内嵌图片 → 幻灯片配图

### PDF
- 段落分隔（三个以上换行）→ 幻灯片分隔
- 第一行 → 幻灯片标题
- 后续行 → 幻灯片要点

## 项目结构

```
src/
├── index.ts              # 入口文件，Express 服务
├── types.ts              # 类型定义
├── services/
│   ├── parser.service.ts # 文档解析服务
│   ├── ppt.service.ts    # PPT 生成服务
│   └── image.service.ts  # AI 配图服务
└── uploads/              # 上传文件存储目录
```

## 技术栈

- **后端框架**: Express + TypeScript
- **PPT 生成**: pptxgenjs
- **文档解析**:
  - Markdown: marked
  - DOCX: mammoth
  - PDF: pdf-parse
- **AI 配图**: OpenAI DALL-E API

## 注意事项

1. **PDF 解析限制**: PDF 是扁平化文本，可能无法完美还原层级结构
2. **图片提取**: 目前仅支持 DOCX 内嵌图片，PDF 图片提取需要额外处理
3. **AI 配图**: 需要配置有效的 OpenAI API Key

## 开发

```bash
# 监听模式开发
npm start

# 编译
npm run build

# 运行编译后的代码
npm run serve
```