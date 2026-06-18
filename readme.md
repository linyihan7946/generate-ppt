# generate-ppt

本地技术文档转 PPT 服务。

当前内置模板：

- 技术类ppt（不生成图片）

## 启动

双击 `start.bat`，或在命令行运行：

```powershell
python app.py
```

服务启动后会自动打开页面：

```text
http://127.0.0.1:8765
```

如果 8765 已被占用，服务会自动使用后续可用端口，并在终端打印实际地址。

## 支持输入

- PDF
- Word `.docx`
- Markdown `.md`
- 文本 `.txt`

## 输出特点

- 横版 16:9，适合技术讲解录屏
- 按原文档顺序生成内容
- 一页一个核心观点
- 自动生成流程图、关系图、对比图、问题-原因-方案-结果结构
- 自动预留截图区域
- 点击逐步淡入动画
- 不生成 AI 图片，不依赖外部模型

## 新增模板

在 `generate_ppt/templates/` 下新增模板类，然后在 `generate_ppt/templates/registry.py` 注册即可。
