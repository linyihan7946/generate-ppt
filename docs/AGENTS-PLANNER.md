# AGENTS-PLANNER.md - PlannerService 详细文档

本文档详细说明 PlannerService 的设计、实现和维护要点。这是整个项目最复杂、最关键、最值得优先阅读的模块。

## 核心定位

PlannerService 是整个项目的"大脑"，负责将原始文档内容重构成适合演示的幻灯片结构。

**一句话职责**：决定页面顺序、内容重组、补充演示页、生成关键字段。

**为什么最关键**：
- 最终 PPT 的"观感问题"大多源于规划器前面就把页面目标定义错了
- 该做时间线的内容被当成普通 bullets
- 该做对比页的内容没有抽成双列结构
- 收尾没有 summary，导致演示缺闭环
- 标题像原文摘录，而不是演示标题
- 生成了给模型看的说明，而不是给观众看的文案

## 主流程架构

### planDocument 主方法

```typescript
async planDocument(docData: DocumentData, options: PlannerOptions = {}): Promise<DocumentData>
```

**执行流程**：
1. 解析规划模式和偏好设置
2. 构建启发式初稿（buildHeuristicPlan）
3. 如果启用 LLM，调用 Gemini 生成高质量规划
4. 合并启发式初稿和 LLM 规划
5. 稀疏页扩写（expandSparseSlidesIfNeeded）
6. 增强叙事连续性（strengthenNarrativeContinuity）
7. 唯一化标题（ensureUniqueTitles）
8. 清理演示语言（sanitizePresentationLanguage）

**关键设计**：双路径策略 - 启发式兜底 + 可选 LLM 增强

### 双路径策略详解

**启发式路径**：
- 无需外部模型，基于规则和算法
- 保证系统不会因模型不可用而完全失效
- 提供稳定的基础输出

**LLM 路径**：
- 调用 Gemini 生成高质量规划
- 提升内容组织、叙事感与表达质量
- 失败时自动回退到启发式结果

**合并策略**：
- LLM 输出与本地初稿智能合并
- 保留启发式的稳定性和 LLM 的创造性
- 根据模式（strict/creative）调整合并权重

## 关键方法详解

### 1. buildHeuristicPlan（启发式规划）

**职责**：构建无需 LLM 的基础规划

**流程**：
1. 标准化源文档幻灯片
2. 推断每个幻灯片的 slideRole
3. 根据角色丰富幻灯片内容
4. 分析文档理解（UnderstandingService）
5. 构建 DeckBrief（全局约束）
6. 判断是否需要添加 Agenda 页
7. 确保收尾页（summary/next_step）
8. 唯一化标题

**关键点**：
- slideRole 推断是质量敏感环节
- Agenda 页判断基于幻灯片数量和格式偏好
- 收尾页确保演示闭环

### 2. inferSlideRole（角色推断）

**职责**：根据内容特征推断幻灯片角色

**推断逻辑**：
```typescript
if (preferences.focus === 'timeline' || looksLikeTimeline(slide)) → 'timeline'
if (preferences.focus === 'comparison' || looksLikeComparison(slide)) → 'comparison'
if (preferences.focus === 'process' || looksLikeProcess(slide)) → 'process'
if (looksLikeDataHighlight(slide)) → 'data_highlight'
if (isSectionDividerCandidate(slide, index, allSlides)) → 'section_divider'
if (preferences.deckFormat === 'presenter' && bullets.length <= 3) → 'key_insight'
default → 'content'
```

**重要**：slideRole 一旦判断错误，后续渲染布局通常也会跟着偏掉。

### 3. generatePlanWithGemini（LLM 规划）

**职责**：调用 Gemini API 生成高质量规划

**流程**：
1. 检查是否可用 Worker Proxy
2. 解析认证 Token
3. 构建 Planner Prompt
4. 调用 LLM API
5. 提取模型文本
6. 解析 JSON 结构
7. 标准化幻灯片数据

**Prompt 构建要点**：
- 明确角色定位："professional presentation strategist"
- 强调源文档约束："stay grounded in the provided source slides"
- 指定输出格式：严格 JSON schema
- 提供模式指令：strict 或 creative
- 包含偏好信息：audience、focus、style、length

### 4. expandSparseSlidesIfNeeded（稀疏页扩写）

**职责**：处理只有一两条 bullet、内容不够撑起一页的问题

**价值**：
- 避免 PPT 看起来像半成品
- 让段落之间更连贯
- 给渲染器提供更稳定的文本材料

**触发条件**：
- `PLANNER_EXPAND_SPARSE_CONTENT=true`（默认启用）
- 幻灯片 bullets 数量少于阈值

**扩写策略**：
- 调用 LLM 生成补充内容
- 保持与原文的关联性
- 不引入未支持的事实

### 5. sanitizePresentationLanguage（语言清洗）

**职责**：清理不应出现在用户可见页面里的文本

**清洗对象**：
- Planner artifact 文本（如 "help understand"、"presentation framing"）
- 长英文短语（中文 deck 中不应出现）
- 提示词式说明文案
- 元信息泄漏

**清洗方法**：
- `isPlannerArtifactText`：检测规划器生成的元信息
- `containsLongEnglishPhrase`：检测中文 deck 中的英文混杂
- `sanitizeRoleTitle`：标准化角色标题（如 "agenda" → "内容导航"）

**重要**：这是防止"页面看起来像 prompt"的关键防线。

## 核心数据结构

### PlannedSlide（规划输出）

```typescript
interface PlannedSlide {
    title: string;
    summary: string;
    bullets: string[];
    layout: SlideLayoutType;
    imageIntent: string;
    imagePrompt: string;
    slideRole: SlideRole;
    keyMessage: string;
    speakerNotes: string[];
    sourceRefs: number[];
}
```

### PlanningPreferences（规划偏好）

```typescript
interface PlanningPreferences {
    audience: DeckAudience;      // general | beginner | executive | student | technical
    focus: DeckFocus;            // overview | timeline | argument | process | comparison
    style: DeckStyle;            // professional | minimal | bold | educational
    deckFormat: DeckFormat;      // presenter | detailed
    length: DeckLength;          // short | default | long
}
```

## 配置选项详解

### 环境变量

```env
ENABLE_PLANNER=true                      # 启用规划器
PLANNER_MODE=strict                      # strict 或 creative
PLANNER_MODEL=gemini-3.1-pro-preview     # 使用的模型
PLANNER_AUTH_TOKEN=                      # 认证 Token
PLANNER_API_BASE_URL=https://www.aigenimage.cn:3001  # API 基础 URL
PLANNER_USE_WORKER_PROXY=false           # 是否使用 Worker Proxy
PLANNER_CONTENT_MODE=strict              # 内容模式
PLANNER_EXPAND_SPARSE_CONTENT=true       # 稀疏页扩写
PLANNER_USE_GUEST_LOGIN=false            # 是否使用访客登录
```

### PlannerMode（规划模式）

**strict 模式**：
- 尽量保持原文内容和措辞
- 只做最小化的结构调整
- 适合需要高度保真的场景

**creative 模式**：
- 允许 Gemini 润色措辞、提炼要点
- 使幻灯片更适合演示表达
- 不改变事实，但提升表达质量

### Worker Proxy 模式

**用途**：通过 Cloudflare Worker 转发请求到上游模型接口

**配置**：
```env
PLANNER_USE_WORKER_PROXY=true
CLOUDFLARE_WORKER_URL=https://your-worker.example.com
LLM_API_KEY=your_provider_key
GOOGLE_API_KEY=your_google_key
```

**注意**：
- Worker Proxy 默认关闭
- 当前实现仍需要真实 provider key
- 如果已配置项目统一中转接口，一般不需要开启

## 维护要点

### 1. 输出 JSON 稳定性

**检查项**：
- JSON 是否符合预期 schema
- 字段是否缺失或类型错误
- slideRole 是否漂移

**维护建议**：
- 定期用相同输入测试输出稳定性
- 建立 JSON schema 验证机制
- 监控 LLM 输出的异常模式

### 2. slideRole 判断准确性

**检查项**：
- 时间线页是否正确识别
- 对比页是否正确识别
- 流程页是否正确识别
- 章节分隔页是否合理

**维护建议**：
- 增强 inferSlideRole 的启发式规则
- 监控 role 判断错误的案例
- 调整 LLM prompt 中的 role 指令

### 3. 语言清洗彻底性

**检查项**：
- 是否仍有元信息泄漏
- 中文 deck 是否混入英文说明
- 标题是否像"分析口吻"而非"演示口吻"

**维护建议**：
- 扩展 isPlannerArtifactText 的检测模式
- 增强 sanitizeRoleTitle 的映射规则
- 监控 EvaluatorService 的 rendered-text 检查结果

### 4. 稀疏页扩写适度性

**检查项**：
- 扩写是否过度脑补
- 是否脱离原文事实
- 是否引入幻觉内容

**维护建议**：
- 调整扩写 prompt 的约束强度
- 监控扩写内容的源文档关联性
- 建立 sourceRefs 覆盖率检查

### 5. sourceRefs 溯源能力

**检查项**：
- sourceRefs 是否正确指向原文
- 溯源信息是否完整
- 是否保留原文关联性

**维护建议**：
- 确保 sourceRefs 在合并时不丢失
- 监控 EvaluatorService 的 sourceRef 覆盖率
- 增强源文档到规划结果的追溯链路

## 常见问题与解决方案

### 问题 1：页面看起来像 prompt

**症状**：
- 页面出现 "AI-Synthesized Deck"、"Content slides" 等元信息
- 中文 deck 中出现长英文说明
- 标题像任务描述而非演示标题

**排查步骤**：
1. 检查 sanitizePresentationLanguage 是否执行
2. 检查 isPlannerArtifactText 的检测模式
3. 检查 LLM 输出的原始文本
4. 检查 EvaluatorService 的 rendered-text 检查结果

**解决方案**：
- 扩展 artifact 文本检测模式
- 增强 sanitizeRoleTitle 映射
- 调整 LLM prompt 约束指令

### 问题 2：slideRole 判断错误

**症状**：
- 该做时间线的内容被当成普通 bullets
- 该做对比页的内容没有抽成双列结构
- 章节分隔页位置不合理

**排查步骤**：
1. 检查 inferSlideRole 的推断逻辑
2. 检查 looksLikeTimeline/Comparison/Process 的判断条件
3. 检查 LLM 输出的 slideRole 字段
4. 检查 preferences.focus 的设置

**解决方案**：
- 增强 role 推断的启发式规则
- 调整 LLM prompt 的 role 指令
- 增加人工标注的 role 判断训练数据

### 问题 3：标题不适合演示

**症状**：
- 标题像原文摘录而非演示标题
- 标题过长或过于学术化
- 标题缺乏吸引力和节奏感

**排查步骤**：
1. 检查 cleanText 的标题处理
2. 检查 LLM 输出的标题字段
3. 检查 sanitizeRoleTitle 的映射规则
4. 检查标题唯一化逻辑

**解决方案**：
- 增强 title 清洗规则
- 调整 LLM prompt 的标题指令
- 增加标题风格转换的启发式规则

### 问题 4：收尾页缺失或不合理

**症状**：
- PPT 缺少 summary 页
- PPT 缺少 next_step 页
- 收尾页内容与整体不连贯

**排查步骤**：
1. 检查 ensureClosingSlides 的执行
2. 检查 shouldAddAgenda 的判断条件
3. 检查 buildSummarySlide/buildNextStepSlide 的内容生成
4. 检查叙事连续性增强逻辑

**解决方案**：
- 调整收尾页添加的触发条件
- 增强收尾页内容与整体的关联性
- 监控 EvaluatorService 的 actionCueSlideCount

### 问题 5：LLM 输出不稳定

**症状**：
- JSON 结构不符合预期
- 字段缺失或类型错误
- 输出格式漂移

**排查步骤**：
1. 检查 extractJsonBlock 的解析逻辑
2. 检查 parsePlannedDocument 的标准化处理
3. 检查 LLM API 的响应格式
4. 检查 prompt 的输出格式指令

**解决方案**：
- 增强 JSON 解析的容错性
- 调整 prompt 的格式约束强度
- 建立 JSON schema 验证机制
- 监控 LLM 输出的异常模式

## 模型切换指南

### 切换前准备

1. 保持 ParserService 和 PPTService 不动
2. 只替换 PlannerService 的模型调用实现
3. 准备相同输入文档用于前后对比

### 切换步骤

1. 修改 generatePlanWithGemini 的 API 调用
2. 调整 prompt 格式以适配新模型
3. 测试 JSON 输出稳定性
4. 检查 DocumentData 的关键字段
5. 生成真实 PPT，结合 EvaluatorService 比较分数
6. 人工抽查关键页面（封面、议程、时间线、总结）

### 成功切换标准

- 不破坏 JSON schema
- 不降低 role 判断准确率
- 不引入可见元信息泄漏
- 稀疏页不再出现"半页空白"
- 中文 deck 的语言风格仍自然

### 切换后监控

- 定期运行回归测试
- 监控 EvaluatorService 的质量分数
- 监控用户反馈的观感问题
- 建立模型切换前后的质量基线

## 性能优化建议

### 1. 并行化处理

- 稀疏页扩写可并行处理多个幻灯片
- LLM 调用可使用批处理模式
- 图片 prompt 生成可并行化

### 2. 缓存机制

- 缓存相同输入的规划结果
- 缓存 LLM 调用的响应
- 缓存 UnderstandingService 的分析结果

### 3. 增量更新

- 对于相似文档，复用部分规划结果
- 对于修改后的文档，增量更新规划
- 避免完全重新规划相似内容

## 测试建议

### 单元测试

- inferSlideRole 的推断准确性
- sanitizePresentationLanguage 的清洗效果
- parsePlannedDocument 的 JSON 解析容错性
- normalizePlannedSlide 的标准化处理

### 集成测试

- 完整规划流程的端到端测试
- LLM 调用的稳定性测试
- 与其他服务的协作测试

### 回归测试

- 相同输入的输出稳定性测试
- 模型切换前后的质量对比测试
- 不同配置下的输出一致性测试

## 相关文档

- [AGENTS.md](./AGENTS.md) - 项目开发指南
- [ARCHITECTURE.md](./ARCHITECTURE.md) - 详细架构说明
- [src/types.ts](../src/types.ts) - 核心类型定义
- [src/services/planner.service.ts](../src/services/planner.service.ts) - 源代码