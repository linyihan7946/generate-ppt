# Document to PPT Generator

将 `Word / Markdown / PDF` 文档自动生成 PPT，支持：

- 按文档层级生成内容页（优先保留原始结构与顺序）
- 自动排版（标题页 + 内容页 + 图文布局）
- 自动配图（优先 AI 生成，失败时自动降级兜底）

## 1. 安装

```bash
npm install
```

## 2. 环境变量

复制 `.env.example` 为 `.env`，并按需修改：

```env
IMAGE_API_KEY=your_api_key_here
IMAGE_API_BASE_URL=https://www.aigenimage.cn
PORT=3000
ENABLE_AI_IMAGES=true
IMAGE_CONCURRENCY=2
```

说明：

- `ENABLE_AI_IMAGES=false` 时跳过自动配图
- `IMAGE_CONCURRENCY` 控制并发配图数，建议 `2~4`

## 3. 启动 Web 服务

```bash
npm start
```

打开 [http://localhost:3000](http://localhost:3000) 上传文件。

## 4. 命令行直接生成（推荐做批量/调试）

```bash
npm run generate -- --input input/计算机发展史.docx --output output/计算机发展史-optimized.pptx
```

若不指定 `--output`，会自动输出到 `output/` 并带时间戳。

## 5. API

`POST /generate-ppt`

- `multipart/form-data`
- 字段：`file`（支持 `.md/.docx/.pdf`）
- 返回：生成好的 `.pptx`

示例：

```bash
curl -X POST http://localhost:3000/generate-ppt \
  -F "file=@input/计算机发展史.docx" \
  --output output/from-api.pptx
```

## 6. 解析策略（重点）

- `DOCX`：优先解析标题 + 多层列表，生成层级化 slide，附带 `breadcrumb`
- `Markdown`：按 heading/list/paragraph 解析
- `PDF`：按段落分块解析（PDF 原生结构有限）

## 7. 常见问题

1. 图片不是每页都来自 AI？
- 当主图 API 返回失败（如内容审核/模型拒绝）时，会自动使用兜底图源，确保每页仍有配图。

2. 为什么服务启动报 Node 兼容问题？
- 项目已内置 `Object.hasOwn` 兼容补丁；若仍异常，建议 Node `>=16`。

3. 为什么某些页只有标题没有正文？
- 源文档对应层级可能本身无下级条目，生成器会保留该节点，避免改变原始层级逻辑。
