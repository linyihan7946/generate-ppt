# AGENTS.md - 项目开发指南

本文档面向 AI Agent 和开发者，提供项目的核心架构、关键模块和开发指南。

## 项目概述

这是一个"文档转 PPT"生成器，核心价值在于将原始文档重构成更适合演示表达的幻灯片，而非简单的格式转换。

**核心流程**：
```
原始文档 → 结构化解析 → 演示化规划 → 图片增强 → PPT渲染 → 质量评估
```

**一句话心智模型**：
```
ParserService → UnderstandingService → PlannerService → ImageService → PPTService → EvaluatorService
```

## 目录结构

```
src/
  cli.ts                         CLI 入口（本地文件生成）
  index.ts                       Web 服务入口（Express API）
  types.ts                       核心领域类型定义
  polyfills.ts                   运行时兼容补丁
  services/
    parser.service.ts            文档解析（Markdown/DOCX/PDF）
    understanding.service.ts     轻量语义理解与主题抽取
    planner.service.ts           演示结构规划与叙事增强 ⭐最关键
    image.service.ts             图片补全与生成
    ppt.service.ts               PPT 渲染（pptxgenjs）
    ppt-image.service.ts         HTML→PNG→PPT 渲染管道
    evaluator.service.ts         质量评估与报告输出
    chat.service.ts              对话式生成服务
    screenshot.service.ts        截图服务
    slide-renderer.service.ts    幻灯片渲染器
public/
  index.html                     简单调试页面
output/                          生成的 PPT 与质量报告
docs/
  ARCHITECTURE.md                详细架构说明
  AGENTS-PLANNER.md              PlannerService 详细文档 ⭐
```

## 核心领域模型

### DocumentData（核心中间态）

```typescript
interface DocumentData {
    title: string;
    slides: SlideContent[];
    brief?: DeckBrief;           // 全局约束和目标
    understanding?: UnderstandingResult;  // 语义信号
}
```

### SlideContent（页面描述对象）

```typescript
interface SlideContent {
    title: string;
    bullets: string[];
    images: string[];
    level?: number;
    breadcrumb?: string;
    summary?: string;
    layout?: SlideLayoutType;
    imageIntent?: string;
    imagePrompt?: string;
    slideRole?: SlideRole;       // ⭐渲染层关键桥梁
    keyMessage?: string;
    speakerNotes?: string[];
    sourceRefs?: number[];
}
```

### SlideRole（页面角色类型）

```typescript
type SlideRole =
    | 'content'          // 正文页
    | 'agenda'           // 目录页
    | 'section_divider'  // 章节分隔页
    | 'key_insight'      // 关键洞察页
    | 'timeline'         // 时间线页
    | 'comparison'       // 对比页
    | 'process'          // 流程页
    | 'data_highlight'   // 数据高亮页
    | 'summary'          // 总结页
    | 'next_step';       // 下一步页
```

**重要**：`slideRole` 一旦判断错误，后续渲染布局通常也会跟着偏掉，是质量敏感字段。

## 主要模块简介

### 1. ParserService

**职责**：解析原始文档，保留结构信息

**支持格式**：
- Markdown：以标题层级拆分 slide，识别图片
- DOCX：使用 mammoth 转 HTML，抽取内嵌图片，多套兜底解析策略
- PDF：通过 pdf-parse 抽文本，语义保真度相对最低

**推荐输入**：DOCX 和 Markdown 是更优输入，PDF 适合兼容输入

### 2. UnderstandingService

**职责**：从页面集合提取 deck 级别语义信号

**提供信息**：
- 章节标题、核心主题
- 时间线信号、对比信号、流程信号
- 关键数字、thesis

**特点**：轻量、可解释、无需外部模型也能工作

### 3. PlannerService ⭐最关键模块

**职责**：决定页面顺序、内容重组、补充演示页、生成关键字段

**为什么最关键**：
- 最终 PPT 的"观感问题"大多源于规划器前面就把页面目标定义错了
- 该做时间线的内容被当成普通 bullets
- 该做对比页的内容没有抽成双列结构
- 收尾没有 summary，导致演示缺闭环

**详细文档**：参见 [AGENTS-PLANNER.md](./docs/AGENTS-PLANNER.md)

### 4. ImageService

**职责**：给缺图页面生成图片

**设计理念**："增强而不阻塞"

**失败回退策略**：
1. 用更安全、更泛化的 prompt 重试
2. 下载占位图
3. 本地极小像素图兜底

**注意**：图片生成不保证语义完全精准，更像"提升观感"的补充层

### 5. PPTService

**职责**：使用 pptxgenjs 生成真实 .pptx

**特点**：
- 按 `slideRole` 分发到不同页面模板
- 长内容自动分页
- 封面页从图片中挑选可用素材
- 过滤不应展示给观众的 presenter artifact 文本

**本质**：多模板渲染引擎，而非纯模板填充层

### 6. EvaluatorService

**职责**：启发式质量评估，输出 JSON 和 Markdown 报告

**评估维度**：
- logic（结构逻辑）
- layout（布局质量）
- imageSemantics（图片语义对齐）
- contentRichness（内容丰富度）
- audienceFit（受众匹配度）
- consistency（一致性）
- sourceUnderstanding（源文档理解）

**特点**：直接解析已生成的 .pptx 内部 XML，检查最终可见文本

### 7. ChatService

**职责**：对话式生成 PPT

**工作流程**：
- gathering：需求收集阶段
- outline：输出大纲供用户确认
- confirmed：用户确认后生成最终 PPT

**特点**：通过自然对话逐步了解需求，分阶段生成

## 运行入口

### Web 入口（src/index.ts）

- 启动 Express 服务（默认端口 3000）
- 接收上传文件，调用完整生成链路
- 返回生成后的 .pptx，附带质量分数

**适合**：手工上传测试、接 UI 或外部系统、在线服务化

### CLI 入口（src/cli.ts）

- 从本地文件路径读取输入文档
- 调用完整生成链路
- 将 .pptx 和质量报告输出到 output/

**适合**：本地调试、批量生成、回归验证、A/B 对比

## 关键配置

### Planner 相关

```env
ENABLE_PLANNER=true
PLANNER_MODE=strict              # strict 或 creative
PLANNER_MODEL=gemini-3.1-pro-preview
PLANNER_AUTH_TOKEN=              # 或使用 LLM_AUTH_TOKEN
PLANNER_API_BASE_URL=https://www.aigenimage.cn:3001
PLANNER_USE_WORKER_PROXY=false   # 默认关闭
PLANNER_CONTENT_MODE=strict      # strict 或 creative
PLANNER_EXPAND_SPARSE_CONTENT=true
```

### 图片相关

```env
ENABLE_AI_IMAGES=true
IMAGE_API_KEY=your_api_key
IMAGE_API_BASE_URL=https://www.aigenimage.cn
IMAGE_CONCURRENCY=2
IMAGE_MODEL=gemini-3.1-flash-image-preview
IMAGE_RESOLUTION=2K
```

### 渲染相关

```env
PPT_TEMPLATE_STYLE=true
PPT_KEEP_TEXT=true
PPT_IMAGE_ONLY_MODE=false
PPT_MAX_BULLETS_PER_SLIDE=5
PPT_RENDER_MODE=native          # native 或 html
```

### 评估相关

```env
ENABLE_EVALUATION=true
```

## 开发建议

### 新开发者阅读顺序

1. readme.md
2. src/types.ts
3. src/cli.ts
4. src/services/parser.service.ts
5. src/services/planner.service.ts ⭐
6. src/services/ppt.service.ts
7. src/services/evaluator.service.ts

### 调试排障指南

**PPT 逻辑不顺**：优先看 PlannerService、slideRole、DeckBrief

**页面文字怪异**：优先看规划器 prompt 输出、sanitizePresentationLanguage

**图片效果差**：优先看 imagePrompt 是否过泛、页面标题和 bullets 是否具体

**PPT 打开正常但观感差**：优先看 role 判定是否合理、是否发生自动分页

### 模型切换指南

**优先关注**：接口契约和输出稳定性，而非只改 model name

**最容易出问题的地方**：
- 返回 JSON 不稳定
- 字段缺失或类型不对
- slideRole 漂移
- 标题变成"分析口吻"而非"演示口吻"
- 中文 deck 混入英文说明语
- 幻觉扩写，脱离原文事实

**推荐切换步骤**：
1. 保持 ParserService 和 PPTService 不动
2. 只替换 PlannerService 的模型调用实现
3. 用同一份输入文档做前后对比
4. 检查 DocumentData 的关键字段是否稳定
5. 生成真实 PPT，结合 EvaluatorService 比较分数

## 项目设计原则

### 原则 1：优先保证链路可完成

先生成出可交付的 PPT，再逐步提升质量。

### 原则 2：启发式兜底必须存在

不要把系统完全绑定在某个模型供应商上。

### 原则 3：中间态要可解释

DocumentData / SlideContent / DeckBrief 这些结构让问题能定位到具体阶段。

### 原则 4：最终质量要看成品

评估最终 .pptx 的可见结果，远比只看模型返回值更有意义。

### 原则 5：面向观众，而不是面向模型

任何调试信息、prompt 说明、任务元数据都不应出现在用户最终看到的页面里。

## 已知边界

**更适合的场景**：
- 结构化较清晰的 Markdown / DOCX
- 教学讲解、汇报提纲、知识综述类内容

**相对薄弱的场景**：
- 排版复杂的 PDF
- 强视觉叙事型品牌发布稿
- 需要非常精准图表复刻的商业报告
- 对动画、母版、企业品牌规范要求极高的场景

## 给 AI Agent 的提示

如果你是后续接手这个仓库的 AI，请优先记住：

- 先理解 DocumentData，再改任何服务
- 先检查 PlannerService，再怀疑 PPTService
- 遇到"页面看起来像 prompt"的问题，先查语言清洗和 rendered-text 评估
- 更换模型时，先保 schema 稳定，再谈风格优化
- 不要轻易删除启发式兜底逻辑，它们是系统稳定性的关键
- PlannerService 是最复杂、最关键、最值得优先阅读的模块，详见 [AGENTS-PLANNER.md](./docs/AGENTS-PLANNER.md)

## 相关文档

- [ARCHITECTURE.md](./ARCHITECTURE.md) - 详细架构说明
- [AGENTS-PLANNER.md](./docs/AGENTS-PLANNER.md) - PlannerService 详细文档 ⭐
- [readme.md](./readme.md) - 用户使用说明