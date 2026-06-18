# AGENTS.md

本文件为在本仓库中工作的 AI agent / coding agent 提供项目约定。请优先遵循用户的直接要求，其次遵循本文件，再参考通用工程习惯。

## 项目概览

`generate-ppt` 是一个本地运行的技术文档转 PPT 服务。它使用 Flask 提供上传页面和生成接口，读取 PDF、Word、Markdown、TXT 文档，将内容转成内部 `DeckSpec` / `SlideSpec` 模型，再用 `python-pptx` 渲染 16:9 演示文稿，并向生成的 `.pptx` 注入逐步淡入动画。

当前内置模板是 `technical-no-image`，定位为技术类 PPT，不生成 AI 图片，不依赖外部模型。

## 关键目录与文件

- `app.py`：应用入口。默认从 `8765` 开始寻找可用端口，并自动打开浏览器。
- `generate_ppt/web.py`：Flask app、上传校验、下载接口、端口探测、浏览器打开逻辑。
- `generate_ppt/pipeline.py`：生成流水线，串联文档加载、模板构建、PPT 渲染和动画注入。
- `generate_ppt/document_loader.py`：支持 `.pdf`、`.docx`、`.md`、`.markdown`、`.txt` 的文本抽取与归一化。
- `generate_ppt/slide_model.py`：PPT 的内部数据模型和页面类型枚举。
- `generate_ppt/ppt_renderer.py`：使用 `python-pptx` 绘制页面、文字、卡片、流程图、关系图等。
- `generate_ppt/animation.py`：为形状名以 `anim_` 开头的对象注入点击淡入动画。
- `generate_ppt/templates/`：模板实现与注册表。
- `web_templates/index.html`：上传页面模板。
- `static/styles.css`：上传页面样式。
- `tests/test_technical_template.py`：当前核心回归测试，验证横版尺寸、页面数量、关键文字和动画密度。
- `workspace/`：运行时上传和输出目录，已被 `.gitignore` 忽略。

## 本地运行

安装依赖：

```powershell
pip install -r requirements.txt
```

启动服务：

```powershell
python app.py
```

也可以双击或运行：

```powershell
.\start.bat
```

服务会优先使用：

```text
http://127.0.0.1:8765
```

如果端口被占用，会在后续端口中查找可用端口。若不希望自动打开浏览器，可以设置：

```powershell
$env:NO_BROWSER = "1"
python app.py
```

## 测试与验证

运行测试：

```powershell
pytest
```

当前测试重点在 `tests/test_technical_template.py`：

- 生成的 PPT 必须是 13.333 x 7.5 英寸横版。
- 技术模板应生成足够数量的页面。
- 输出文本中应包含关键页面文案。
- 每页应存在 timing，且淡入动画数量达到预期密度。

修改 `ppt_renderer.py`、`animation.py`、`slide_model.py`、`pipeline.py` 或模板逻辑后，至少运行 `pytest`。如果调整了前端上传流程，也应手动启动服务并上传一个小 `.txt` 或 `.md` 文件验证下载链路。

## 开发约定

- 保持项目轻量，不引入外部 AI 服务、远程模型或网络依赖来完成 PPT 内容生成。
- 运行时文件只能写入 `workspace/uploads` 和 `workspace/outputs` 等工作目录，不要提交生成的 `.pptx`、上传样例或缓存。
- 新增文件解析能力时，在 `document_loader.py` 增加加载逻辑，并同步更新 `ALLOWED_EXTENSIONS`、前端 `accept` 属性和 README。
- 新增页面类型时，先在 `SlideKind` 中定义枚举，再让模板产出对应 `SlideSpec`，最后在 `PptRenderer._render_slide()` 中实现渲染分支。
- 渲染器中的动画目标依赖 shape name。需要点击逐步淡入的元素必须使用稳定的 `anim_` 前缀命名，并让命名顺序符合播放顺序。
- `animation.py` 默认直接修改 `.pptx` 内部 XML。改动时要用 zip/XML 结构验证输出仍可被 PowerPoint 打开，并保留测试中的 timing / animEffect 检查。
- `_safe_name()` 会限制输出文件名为 ASCII 安全字符。不要把用户上传文件名直接用于输出路径。
- Flask 接口应继续使用 `secure_filename()` 处理上传文件名，并保持扩展名白名单校验。

## 模板扩展指南

新增模板时：

1. 在 `generate_ppt/templates/` 下新增模板类，提供 `build(document: SourceDocument) -> DeckSpec`。
2. 尽量复用 `SlideSpec`、`SlideKind` 和现有渲染能力；只有确实需要新布局时再扩展 `PptRenderer`。
3. 在 `generate_ppt/templates/registry.py` 注册新的 `TemplateDefinition`。
4. 前端模板下拉和模板卡片会通过 `list_templates()` 自动展示注册内容。
5. 为新模板添加或扩展测试，至少覆盖能生成 `.pptx`、页面尺寸正确、核心文本存在、动画没有丢失。

## 中文与编码注意事项

仓库中包含大量中文界面文案、模板文案和测试断言。当前部分文件在某些终端中可能显示为乱码。处理这些内容时：

- 使用 UTF-8 保存文件。
- 不要为了“顺手修复显示”而大范围重写已有中文字符串，除非任务明确要求修复编码。
- 如果确实修复编码，需要同步更新源码、前端文案、README 和测试断言，并完整运行 `pytest`。
- 修改测试中中文断言前，先确认生成 PPT 的真实文本内容。

## 前端约定

前端是一个单页上传工具，不是营销页。保持界面直接服务于“选择模板、上传文档、生成并下载 PPT”这条主流程。

- 样式位于 `static/styles.css`，模板位于 `web_templates/index.html`。
- 保持暗色技术风格、8px 以内圆角和清晰的上传状态反馈。
- 不要把生成结果写进页面以外的持久位置；下载应继续走 `/download/<filename>`。
- 如果调整 HTML 文案、表单字段或接口路径，必须同步检查 `web.py` 中的请求处理逻辑。

## Git 与提交

- 不要提交 `workspace/`、缓存目录、虚拟环境或生成的 PPT。
- 如果需要生成提交说明、变更摘要、发布说明或 PR 文案，默认使用中文。
- 在已有未提交改动时，先确认改动来源；不要回退与当前任务无关的用户改动。
