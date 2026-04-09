# 项目架构说明

本文档面向两类读者：

- 新接手项目的开发者
- 未来替换或接入其他大模型时，需要快速建立上下文的 AI / Agent

目标不是逐行解释代码，而是帮助读者快速理解这个项目的心智模型、端到端链路、关键模块职责、配置方式，以及真正影响生成质量的核心点与难点。

## 1. 项目定位

这是一个“文档转 PPT”生成器。输入可以是 `Markdown / DOCX / PDF`，输出为 `.pptx` 文件，并伴随一份质量评估报告。

项目不是简单地把原文分段塞进幻灯片，而是分成 5 个阶段：

1. 解析原始文档，尽量保留结构
2. 重组为适合演示的叙事型 PPT 大纲
3. 为页面补图或生成视觉素材
4. 渲染为真实 PPT
5. 对结果做启发式质量评估

这意味着项目的核心价值不在“格式转换”，而在“把文档内容重构成更适合演示表达的 deck”。

## 2. 一句话心智模型

可以把整个系统理解为：

`原始文档 -> 结构化 slide skeleton -> 演示化 narrative deck -> 图文渲染 -> 质量打分`

更具体一点：

`ParserService -> PlannerService -> ImageService -> PPTService -> EvaluatorService`

其中：

- `ParserService` 负责“读懂文档的外在结构”
- `PlannerService` 负责“把内容改写成适合演示的结构”
- `ImageService` 负责“补齐视觉素材”
- `PPTService` 负责“真正把 deck 画出来”
- `EvaluatorService` 负责“给输出质量做回归评估”

## 3. 目录与文件职责

```text
src/
  cli.ts                         CLI 入口
  index.ts                       Web 服务入口
  types.ts                       核心领域类型定义
  polyfills.ts                   运行时兼容补丁
  services/
    parser.service.ts            文档解析
    understanding.service.ts     轻量语义理解与主题抽取
    planner.service.ts           演示结构规划与叙事增强
    image.service.ts             图片补全与生成
    ppt.service.ts               PPT 渲染
    evaluator.service.ts         质量评估与报告输出
  uploads/                       Web 上传临时文件目录
public/
  index.html                     简单调试页面
output/                          生成的 PPT 与质量报告
readme.md                        面向使用者的说明
```

## 4. 运行入口

项目有两个主要入口。

### 4.1 Web 入口

文件：`src/index.ts`

职责：

- 启动 Express 服务
- 接收上传文件
- 调用完整生成链路
- 返回生成后的 `.pptx`
- 在响应头里附带质量分数与报告路径

适合：

- 手工上传测试
- 接 UI 或外部系统
- 做在线服务化

### 4.2 CLI 入口

文件：`src/cli.ts`

职责：

- 从本地文件路径读取输入文档
- 调用完整生成链路
- 将 `.pptx` 和质量报告输出到 `output/`

适合：

- 本地调试
- 批量生成
- 回归验证
- 更换模型后的 A/B 对比

## 5. 端到端数据流

### 第 1 步：解析文档

入口模块：`ParserService`

输入：

- `.md`
- `.docx`
- `.pdf`

输出：

- `DocumentData`

`DocumentData` 是整个项目最重要的中间对象，它代表“已经被结构化后的文档内容”，后续所有服务都围绕它工作。

### 第 2 步：做轻量语义理解

入口模块：`UnderstandingService`

职责：

- 提取章节标题
- 归纳主题与重要性
- 识别时间线、对比、流程、关键数字
- 提炼 thesis / core takeaways

这一步不是独立 API，而是给 `PlannerService` 提供启发式语义基础，帮助后续生成更像演示文稿而不是文档摘抄。

### 第 3 步：规划 deck

入口模块：`PlannerService`

职责：

- 决定页面顺序
- 判断哪些内容应该合并、拆分、重写
- 补充 agenda / summary / next step 等演示页
- 给每一页打上 `slideRole`
- 生成 `keyMessage / summary / speakerNotes / imagePrompt`
- 在需要时调用 LLM 做高质量规划

这一步决定最终 PPT 的上限，是整个项目最关键的模块。

### 第 4 步：补图

入口模块：`ImageService`

职责：

- 给缺图页面生成图片
- 使用缓存避免相同 prompt 重复出图
- 主路径失败时降级到安全 prompt
- 再失败时回退到占位图

图片生成是“增强项”，不是主流程的单点故障。

### 第 5 步：渲染 PPT

入口模块：`PPTService`

职责：

- 使用 `pptxgenjs` 生成真实 `.pptx`
- 根据 `slideRole` 选择不同模板
- 根据内容量自动分页
- 插入图片、标题、要点、页脚、引用等元素

这一层不是简单模板填充，而是“角色驱动的渲染系统”。

### 第 6 步：质量评估

入口模块：`EvaluatorService`

职责：

- 根据 `DocumentData` 和已生成的 `pptx` 做启发式评分
- 输出 JSON 和 Markdown 报告
- 检查结构逻辑、布局、图片语义、内容丰富度、受众匹配度、一致性
- 检查渲染后是否出现元信息泄漏、提示词式说明文案、明显中英文混杂

评估器不是单元测试替代品，但非常适合做“模型切换前后”的质量回归。

## 6. 核心领域模型

文件：`src/types.ts`

建议优先理解以下几个类型。

### 6.1 `DocumentData`

核心中间态，包含：

- `title`
- `slides`
- `brief`
- `understanding`

其中：

- `slides` 是页面级内容
- `brief` 是整套 deck 的高层摘要
- `understanding` 是从原文中抽取出的语义信号

### 6.2 `SlideContent`

每一页的核心结构，典型字段包括：

- `title`
- `bullets`
- `images`
- `level`
- `breadcrumb`
- `summary`
- `layout`
- `imageIntent`
- `imagePrompt`
- `slideRole`
- `keyMessage`
- `speakerNotes`
- `sourceRefs`

可以把它理解成“还没渲染成 PPT 的页面描述对象”。

### 6.3 `DeckBrief`

代表整套 PPT 的全局约束和目标，包括：

- `deckGoal`
- `audience`
- `focus`
- `style`
- `deckFormat`
- `desiredLength`
- `chapterTitles`
- `coreTakeaways`

它主要影响封面、议程页、收尾页，以及整体叙事风格。

### 6.4 `slideRole`

这是渲染层的关键桥梁。常见值包括：

- `content`
- `agenda`
- `section_divider`
- `timeline`
- `comparison`
- `process`
- `data_highlight`
- `key_insight`
- `summary`
- `next_step`

`slideRole` 一旦判断错，后续渲染布局通常也会跟着偏掉，所以它是质量敏感字段。

## 7. 各模块详细说明

### 7.1 ParserService

文件：`src/services/parser.service.ts`

支持三种输入格式。

### Markdown 解析

特点：

- 以标题层级拆分 slide
- 列表转 bullets
- 识别 Markdown 图片
- 在内容过少时做最小兜底

### DOCX 解析

特点：

- 使用 `mammoth` 转 HTML
- 抽取内嵌图片为 base64
- 优先尝试“顶层列表结构”解析
- 其次按标题解析
- 最后再按段落块兜底

这是项目里比较讲究的一块，因为很多教学文档或汇报材料在 Word 里并不靠标题，而是靠缩进列表表达层级。

### PDF 解析

特点：

- 通过 `pdf-parse` 抽文本
- 按段落与空行做启发式切分
- 语义保真度相对最低

结论：

- `DOCX` 和 `Markdown` 是更优输入
- `PDF` 更适合作为兼容输入，而不是最佳输入

### 7.2 UnderstandingService

文件：`src/services/understanding.service.ts`

作用是从“页面集合”里提取 deck 级别信号，而不是直接生成 PPT 内容。它会给规划器提供：

- 章节标题
- 核心主题
- 时间线信号
- 对比信号
- 流程信号
- 关键数字信号
- thesis

它是一个轻量、可解释、无需外部模型也能工作的语义增强层。

### 7.3 PlannerService

文件：`src/services/planner.service.ts`

这是项目最复杂、最关键、最值得优先阅读的文件。

### 规划器的真实职责

它不是“调用一下大模型生成 JSON”这么简单，而是一个带多层兜底的规划系统：

1. 先基于原始内容构造启发式初稿
2. 如果配置允许，再调用 LLM 做高质量规划
3. 将 LLM 输出与本地初稿合并
4. 对内容稀疏页面做二次扩写
5. 增强叙事连续性
6. 去重标题
7. 清理不应出现在用户可见页面里的提示词、调试信息、元信息

### 为什么它最关键

因为最终 PPT 的大多数“观感问题”并不是渲染器造成的，而是规划器前面就把页面目标定义错了，例如：

- 该做时间线的内容被当成普通 bullets
- 该做对比页的内容没有抽成双列结构
- 收尾没有 summary，导致演示缺闭环
- 标题像原文摘录，而不是演示标题
- 生成了给模型看的说明，而不是给观众看的文案

### 规划器的双路径策略

规划器默认采用“启发式 + 可选 LLM”的双路径：

- 没有 LLM 时，项目也能工作
- 有 LLM 时，重点提升内容组织、叙事感与表达质量

这种设计保证了系统不会因为模型不可用而完全失效。

### 稀疏页扩写是第二个关键点

规划器除了“主规划”外，还有一次“稀疏页扩写”阶段，用来处理只有一两条 bullet、内容不够撑起一页的问题。

它的价值在于：

- 避免 PPT 看起来像半成品
- 让段落之间更连贯
- 给渲染器提供更稳定的文本材料

### 与模型相关的接口

当前主要有两类调用方式：

- 直接调用项目配置的 LLM 中转接口，例如 `/api/llm/direct`
- 显式开启 `worker proxy` 后，通过 worker 再转发到上游模型接口

注意：

- `worker proxy` 现在是显式开关，默认关闭
- 当前 worker 路径本质上仍是“代转发 provider 请求”，不是“服务端完全代持 provider key”的免密模式
- 如果已经配置了项目自己的中转接口，一般不需要开启 worker proxy

### 规划器维护时最该关注什么

- 输出 JSON 是否稳定符合 schema
- `slideRole` 判断是否合理
- 中文 deck 的语言清洗是否彻底
- 稀疏页扩写是否过度脑补
- `sourceRefs` 是否保留原文溯源能力

### 7.4 ImageService

文件：`src/services/image.service.ts`

设计理念是“增强而不阻塞”。

主路径：

- 针对缺图页面生成图片 prompt
- 调用图片 API

失败后的回退：

1. 用更安全、更泛化的 prompt 重试
2. 下载占位图
3. 最后用本地极小像素图兜底

这让系统在图片 API 不稳定时仍能产出可打开的 PPT。

需要注意：

- 图片生成并不保证语义完全精准
- 更像“提升观感”的补充层
- 真正影响图片质量的是上游 `imagePrompt` 质量和页面 role 判断

### 7.5 PPTService

文件：`src/services/ppt.service.ts`

这是渲染层核心。

它会把 `SlideContent` 转成实际页面元素，包括：

- 标题
- 摘要
- bullets
- 图片
- 页脚
- 页码
- 来源引用

### 为什么渲染器不是纯模板层

因为它内部按 `slideRole` 分发到不同页面类型，例如：

- `agenda`
- `section_divider`
- `timeline`
- `comparison`
- `process`
- `data_highlight`
- `summary`
- `next_step`
- `key_insight`
- `content`

不同 role 的视觉结构差异很大，所以这里本质上是一个“多模板渲染引擎”。

### 渲染器的几个关键行为

- 长内容自动分页
- 封面页会从图片中挑选可用素材
- 议程页和封面页会读取 `DeckBrief`
- 会过滤不应展示给观众的 presenter artifact 文本
- 可按环境变量控制是否保留文本、是否显示引用、是否进入图片优先模式

### 渲染层的典型风险

- 规划器清洗不彻底，导致元信息被渲染出来
- 页面信息过多，自动分页后失去节奏
- 页面 role 和模板不匹配，导致版式别扭

### 7.6 EvaluatorService

文件：`src/services/evaluator.service.ts`

评估器的重要性经常被低估，但它其实是这个项目后续演进的“护栏”。

它会输出：

- 总分
- 等级
- 各维度分数
- 结构化指标
- Markdown 报告

### 当前重点评估维度

- `logic`
- `layout`
- `imageSemantics`
- `contentRichness`
- `audienceFit`
- `consistency`

### 它的一个重要特点

除了看 `DocumentData`，它还会直接解析已经生成的 `.pptx` 内部 XML，提取“最终可见文本”再做检查。

新版评估还会识别渲染后的图片覆盖情况，并对“整页图片化 / image-first”的 deck 放宽一部分纯文字启发式惩罚，避免高质量视觉型 PPT 因为可提取文字过少而被系统性低估。

这非常关键，因为有些问题只会在渲染后暴露，例如：

- `AI-Synthesized Deck`
- `Content slides`
- `Audience / Format / Focus / Style`
- 提示词式英文辅助说明
- 中文 deck 中夹杂不必要的长英文 narration

也就是说，它不是只评“计划稿”，而是在评“最终交付物”。

## 8. 运行方式

### 8.1 安装

```bash
npm install
```

### 8.2 本地启动 Web 服务

```bash
npm run start
```

默认服务地址：

- `http://localhost:3000`

### 8.3 CLI 生成

```bash
npx ts-node src/cli.ts --input ./example.docx
```

也可以带额外参数：

```bash
npx ts-node src/cli.ts --input ./example.docx --planner-mode creative --deck-format presenter --audience general --focus overview --style professional --length default
```

常用输出位置：

- `output/*.pptx`
- `output/*.quality.json`
- `output/*.quality.md`

### 8.4 构建生产版本

```bash
npm run build
npm run serve
```

## 9. 关键配置与环境变量

详细项可参考 `.env.example`，这里仅列最关键的分组。

### 9.1 Planner 相关

- `PLANNER_ENABLED`
- `PLANNER_MODE`
- `PLANNER_MODEL`
- `PLANNER_AUTH_TOKEN`
- `PLANNER_API_BASE_URL`
- `PLANNER_USE_WORKER_PROXY`
- `CLOUDFLARE_WORKER_URL`
- `GOOGLE_API_KEY`

理解要点：

- 推荐优先使用项目统一中转接口
- `worker proxy` 默认应关闭
- 如果开启 `worker proxy`，当前实现通常仍需要真实 provider key

### 9.2 图片相关

- `IMAGE_ENRICH_ENABLED`
- `IMAGE_API_KEY`
- `IMAGE_API_BASE_URL`

### 9.3 评估相关

- `EVALUATION_ENABLED`
- `QUALITY_REPORT_ENABLED`

### 9.4 渲染相关

- `PPT_TEMPLATE_STYLE`
- `PPT_IMAGE_ONLY_MODE`
- `PPT_KEEP_TEXT`
- `PPT_MAX_BULLETS_PER_SLIDE`
- `PPT_SHOW_SOURCE_REFS`

## 10. 模型切换指南

如果后续要把大模型换成别家，优先关注的是“接口契约”和“输出稳定性”，不是只改一个 model name。

### 10.1 先看规划器，不要先看渲染器

模型切换最先受影响的是：

- `PlannerService` 主规划
- 稀疏页扩写
- 语言清洗前的原始输出风格

渲染器通常是被动接收上游结构化结果。

### 10.2 最容易出问题的地方

- 返回 JSON 不稳定
- 字段缺失或类型不对
- `slideRole` 漂移
- 标题变成“分析口吻”而非“演示口吻”
- 中文 deck 混入英文说明语
- 幻觉扩写，脱离原文事实

### 10.3 推荐的切换步骤

1. 先保持 `ParserService` 和 `PPTService` 不动
2. 只替换 `PlannerService` 的模型调用实现
3. 用同一份输入文档做前后对比
4. 检查 `DocumentData` 的关键字段是否稳定
5. 生成真实 PPT，结合 `EvaluatorService` 比较分数与问题项
6. 人工抽查封面、议程页、时间线页、总结页

### 10.4 成功切换的最低标准

- 不破坏 JSON schema
- 不降低 role 判断准确率
- 不引入可见元信息泄漏
- 稀疏页不再出现“半页空白”
- 中文 deck 的语言风格仍自然

## 11. 项目里的真正难点

这部分很重要。新接手时，如果只看到“文档转 PPT”，很容易低估问题复杂度。

### 难点 1：原始文档结构并不天然适合演示

文档写作和 PPT 表达是两种不同媒介：

- 文档适合完整叙述
- PPT 适合高密度提炼、节奏控制和视觉对齐

所以不能只做“切段落”，必须做“叙事重组”。

### 难点 2：Word 结构信号不稳定

很多 DOCX 并不规范使用标题样式，而是靠：

- 缩进
- 列表层级
- 段落顺序

这也是 `ParserService` 里有多套兜底解析策略的原因。

### 难点 3：模型输出看似正确，但未必适合直接给观众看

LLM 经常会输出：

- 面向模型的解释语
- 面向提示词的结构说明
- 英文化的模板占位文案

这些内容语义上“看起来合理”，但一旦进 PPT 就会非常违和。

### 难点 4：问题有时只在渲染后才暴露

例如某些文本在中间态里不明显，但最终排版到封面或侧栏里就会显得非常刺眼。

这也是为什么评估器必须直接读取最终 `.pptx` 文本。

### 难点 5：必须允许失败但不能整体崩掉

这个项目依赖多个不稳定环节：

- 外部 LLM
- 图片 API
- 不同格式解析质量

因此架构上采用了大量 graceful degradation：

- LLM 不可用时回退启发式规划
- 图片失败时回退占位图
- 解析失败时继续走段落兜底

目标不是“每一步都完美”，而是“尽量始终交付一个能用的 PPT”。

## 12. 当前项目的核心设计原则

可以把项目总结成 5 条原则。

### 原则 1：优先保证链路可完成

先生成出可交付的 PPT，再逐步提升质量。

### 原则 2：启发式兜底必须存在

不要把系统完全绑定在某个模型供应商上。

### 原则 3：中间态要可解释

`DocumentData / SlideContent / DeckBrief` 这些结构存在的意义，就是让问题能定位到具体阶段，而不是全部藏在 prompt 里。

### 原则 4：最终质量要看成品，不只看中间 JSON

评估最终 `.pptx` 的可见结果，远比只看模型返回值更有意义。

### 原则 5：面向观众，而不是面向模型

任何看起来像：

- 调试信息
- prompt 说明
- 任务元数据
- 面向系统的占位文本

都不应该出现在用户最终看到的页面里。

## 13. 新开发者建议阅读顺序

如果你第一次进入项目，建议按这个顺序阅读：

1. `readme.md`
2. `src/types.ts`
3. `src/cli.ts`
4. `src/services/parser.service.ts`
5. `src/services/planner.service.ts`
6. `src/services/ppt.service.ts`
7. `src/services/evaluator.service.ts`

这样可以先建立主流程，再进入最复杂的规划与渲染细节。

## 14. 调试与排障建议

### 当 PPT 逻辑不顺时

优先看：

- `PlannerService`
- `slideRole`
- `DeckBrief`

### 当页面文字怪异或出现说明性文案时

优先看：

- 规划器 prompt 输出
- `sanitizePresentationLanguage`
- `PPTService` 是否仍有未过滤的 artifact 文本
- `EvaluatorService` 的 rendered text 检查结果

### 当图片效果差时

优先看：

- `imagePrompt` 是否过泛
- 页面标题和 bullets 是否足够具体
- 图片 API 是否走到了降级路径

### 当 PPT 打开正常但观感差时

优先看：

- role 判定是否合理
- 是否发生自动分页
- 是否本应做时间线 / 对比 / 流程却退化成普通内容页

## 15. 当前已知边界

项目当前更适合：

- 结构化较清晰的 Markdown / DOCX
- 教学讲解、汇报提纲、知识综述类内容

项目当前相对薄弱的场景：

- 排版复杂的 PDF
- 强视觉叙事型品牌发布稿
- 需要非常精准图表复刻的商业报告
- 对动画、母版、企业品牌规范要求极高的场景

另外，仓库当前没有完整自动化测试体系，质量控制更多依赖：

- 示例文档回归生成
- 质量报告
- 人工抽检关键页面

## 16. 给未来 AI / Agent 的简短提示

如果你是后续接手这个仓库的 AI，请优先记住下面几点：

- 先理解 `DocumentData`，再改任何服务
- 先检查 `PlannerService`，再怀疑 `PPTService`
- 遇到“页面看起来像 prompt”的问题，先查语言清洗和 rendered-text 评估
- 更换模型时，先保 schema 稳定，再谈风格优化
- 不要轻易删除启发式兜底逻辑，它们是系统稳定性的关键

## 17. 总结

这个项目本质上是一个“带多层兜底的演示文稿生成流水线”，而不是单纯的文件格式转换器。

真正决定质量的关键不只是模型强弱，而是以下几件事是否协同良好：

- 解析是否保留结构
- 规划是否具备演示叙事感
- 图片是否只做增强而不拖垮主链路
- 渲染是否按页面角色选择合适模板
- 评估是否能及时发现成品中的退化问题

如果后续要继续演进，最值得持续投入的方向通常不是“再多加几个模板”，而是：

- 提高规划器输出稳定性
- 提升 role 判断准确率
- 强化对最终可见文本的质量约束
- 用评估器建立模型切换前后的回归基线
