# personal-style-ppt2html

[English](README.md) | 中文

`personal-style-ppt2html` 是一个 Agent Skill 和 Python PPTX 转 HTML 工具集，用来把用户已有 PowerPoint 中的个人风格、单位风格和学术汇报风格转成可复用的 HTML presentation。

它适合那些已经有自己常用 `.pptx` 模板、组会汇报风格、单位 logo 横条、论文图展示习惯的人。目标不是生成花哨网页模板，而是保留严肃、克制、可复用的个人学术表达系统。

它不只适用于 Codex，也可以用于 Claude Code，或者任何能够读取 `SKILL.md` 目录并运行本地 Python 脚本的 agent / runtime。`agents/openai.yaml` 是 Codex 的 UI 元数据；核心工作流在 `SKILL.md`、`references/`、`scripts/` 和 `assets/` 中。

## 这个仓库能做什么

- 把 `.pptx` 转成可翻页的 HTML slide preview。
- 保留 PPTX 中可复用的 slide-layout 元素，例如单位 logo、顶部横条、底部规则线、页眉页脚视觉标识。
- 保留 slide-master 元素，并检测那些没有放在 layout 里、但被手动复制到多页的重复视觉块。
- 生成 `asset-registry.json`，记录可复用视觉候选项的语义角色、来源层级、覆盖页码、精确坐标和置信度。
- 生成 `imagegen-briefs.json`，当当前 agent 具备生图工具时，可据此生成与 PPT 风格相似的背景、分隔页 motif 和内容 backplate。
- 从历史 PPTX 中提取字体、主题色、页面比例、标题位置、citation footer、图片密度和常见页面布局。
- 把 PPTX 内嵌媒体复制到本地 HTML assets。
- 根据提取出的风格画像生成 CSS design tokens。
- 增加克制的动效预设，适合现场汇报、录屏讲解和 demo 视频。
- 对生成 HTML 做输出优化，包括本地资源、图片 lazy loading、异步解码、reduced-motion 支持和 print-safe CSS。
- 提供严肃学术场景的 HTML 模板，适合组会、seminar、paper reading、答辩和科研记录。
- 对生成的 HTML 做基础结构和动效安全审计，检查 viewport、print CSS、slide 容器、动效降级等。

## 为什么需要它

市面上很多 HTML presentation 模板视觉效果很强，但并不适合学术汇报。学术汇报通常需要的是另一种质量：

- 安静、克制的视觉层级；
- 准确的论文图、公式、引用和单位；
- 有单位身份，但不能像营销页一样喧宾夺主；
- 方法图、结果图、citation 是内容核心；
- 组会、seminar、答辩之间保持稳定的个人风格。

这个项目把历史 PPTX 当作风格参考，读取其中的 Open XML 结构，然后把可复用设计信号转成 HTML。

## 实现方法

`.pptx` 本质上是一个包含 Office Open XML 的 ZIP 包。

转换器会读取：

- `ppt/slides/slideN.xml`：每页自己的文本、图片、形状和坐标；
- `ppt/slides/_rels/slideN.xml.rels`：每页的关系文件；
- `ppt/slideLayouts/slideLayoutN.xml`：复用的 layout 元素；
- `ppt/slideLayouts/_rels/slideLayoutN.xml.rels`：layout 层的 logo 和媒体资源；
- `ppt/slideMasters/slideMasterN.xml`：master 层的复用标识和规则线；
- `ppt/theme/theme1.xml`：主题色；
- `ppt/media/*`：PPTX 内嵌图片。

这一步很关键，因为 logo 和顶部横条经常不在每一页的 `slideN.xml` 里，而是在 `slideLayout` 或 `slideMaster` 里；还有一些重复视觉块是被手动复制到多页的。现在转换器会顺着 slide 到 layout 再到 master 的关系提取复用元素，视觉挖掘脚本也会聚类重复出现的 slide-local 元素。

生成的 HTML 还可以带 `data-motion` 动效预设。动效只允许短促的透明度和轻微位移动画，并且包含 reduced-motion 和 print 降级规则，所以既适合学术汇报，也适合录屏展示。

如果当前 agent runtime 暴露了生图能力，工作流还可以基于提取出的风格指纹生成非事实性的辅助视觉素材。仓库不会假设所有 agent 都有这个能力：它会先写出 `imagegen-briefs.json`，然后由具备生图能力的 agent 执行这些 brief，并把选中的输出保存到最终 HTML 项目的 `assets/generated/` 目录。

## 仓库结构

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
├── assets/
│   └── academic-html-template/
├── references/
│   ├── academic-style-rules.md
│   ├── animation-and-optimization.md
│   ├── html-layout-patterns.md
│   ├── imagegen-asset-guidelines.md
│   ├── institution-brand-rules.md
│   └── reusable-visual-mining.md
└── scripts/
    ├── audit_html_layout.py
    ├── build_theme_css.py
    ├── extract_pptx_style.py
    ├── make_asset_manifest.py
    ├── mine_reusable_visuals.py
    ├── pptx_common.py
    ├── pptx_to_academic_html.py
    ├── prepare_imagegen_briefs.py
    ├── run_pipeline.py
    └── test_academic_html_tools.py
```

## 快速开始

使用 Python 3。核心 XML 解析路径不依赖第三方包。

完整流程可以直接运行：

```bash
python scripts/run_pipeline.py your-deck.pptx -o work/pipeline --motion recording
```

如果需要手动控制每一步：

```bash
python scripts/extract_pptx_style.py your-deck.pptx -o work/style-profile.json
python scripts/make_asset_manifest.py your-deck.pptx -o work/reference-assets
python scripts/mine_reusable_visuals.py your-deck.pptx -o work/asset-registry.json
python scripts/prepare_imagegen_briefs.py work/style-profile.json --registry work/asset-registry.json -o work/imagegen-briefs.json
python scripts/build_theme_css.py work/style-profile.json -o work/theme.generated.css
python scripts/pptx_to_academic_html.py your-deck.pptx -o work/html-preview --profile work/style-profile.json --motion recording
python scripts/audit_html_layout.py work/html-preview/index.html
```

然后打开：

```text
work/html-preview/index.html
```

生成的 HTML 支持键盘翻页、reduced-motion 偏好，也支持打印或导出 PDF。

如果把 `theme.generated.css` 用在可复用模板里，应放在 `theme.css` 和 `components.css` 之后引入，这样提取出来的 token 才会覆盖基础默认值。

## 动效预设

使用 `--motion none` 可以得到最保守的学术输出。`--motion subtle` 适合现场浏览器汇报，`--motion recording` 适合录屏讲解，`--motion demo` 只建议用于更公开、更展示型的视频片段。

动效系统会避免循环装饰动画，保持 slide-layout 中的 logo 和横条稳定，并且在 reduced-motion 用户设置和打印输出中自动禁用动画。

## 可复用视觉 Registry

当你想判断哪些元素能沉淀成个人风格系统时，运行 `mine_reusable_visuals.py`。它会识别 layout 和 master 继承的对象，也会识别那些位置、类型和视觉身份重复出现的 slide-local 元素。

对于语义不确定的元素，可以把渲染截图和 `asset-registry.json` 一起交给视觉模型审阅。视觉模型应该负责标注和排除风险元素，而不是重新生成官方 logo、论文图、实验图表或事实性内容。

## 可选生成素材

如果当前 agent 具备生图能力，可以读取 `imagegen-briefs.json` 生成安全的非事实性素材：

- `style-cover-backdrop.png`；
- `style-section-divider.png`；
- `style-figure-backplate.png`。

选中的输出应保存到最终 HTML 项目的 `assets/generated/`，并记录 prompt 和工具来源。这些生成素材只能扩展提取出来的风格，不能替代 PPTX 中真实提取出的 logo、单位标识、论文图、数据图或事实性方法图。

默认情况下，生图 prompt 不会包含原始 slide title，避免把私人文本或事实内容泄露给生图模型。只有在明确需要主题化 motif 且文本可安全使用时，才使用 `--include-title-cues`。

## 作为 Agent Skill 安装

这个仓库遵循基于文件系统的 `SKILL.md` skill 结构。即使不在 agent 环境中使用，也可以直接从命令行运行 `scripts/` 里的 Python 工具。

### Codex

把仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.codex/skills/personal-style-ppt2html
```

之后可以这样调用：

```text
Use $personal-style-ppt2html to convert my reference PPTX into an academic HTML deck.
```

### Claude Code

把同一个仓库克隆到 Claude Code 的个人 skills 目录：

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.claude/skills/personal-style-ppt2html
```

如果希望作为某个项目的本地 skill 随仓库共享，可以放到项目目录下：

```bash
mkdir -p .claude/skills
git clone https://github.com/gsynb/personal-style-ppt2html.git .claude/skills/personal-style-ppt2html
```

之后可以直接用自然语言让 Claude Code 调用它；如果你的 Claude Code 环境暴露了 skill command，也可以直接按 skill 名称调用。

## 验证

运行：

```bash
python scripts/test_academic_html_tools.py
```

测试覆盖：

- PPTX 风格提取；
- CSS token 生成；
- PPTX 媒体资产提取；
- HTML 审计；
- image generation brief 生成；
- 带 slide-layout logo 和顶部横条的 PPTX 转 HTML；
- slide-master 保留和可复用视觉 registry 挖掘；
- master/layout 重复元素过滤和母版占位文本清理；
- 动效预设输出、图片优化属性和动画安全警告。

## 当前限制

- `.emf` 和 `.wmf` 不是浏览器原生支持的图片格式。如果本机没有 LibreOffice、Inkscape、ImageMagick 等转换工具，这些资源会显示为占位。
- 复杂 PowerPoint 几何形状会被近似处理。简单色块、横条、图片和文本框效果最好。
- 当前会跟随 slide-layout 和 slide-master 继承，处理常见图片、文本和形状，但复杂 PowerPoint 特效仍然是近似处理。
- 可复用视觉 registry 基于 PPTX 结构和重复几何信息。遇到可能是事实图、论文图或实验图的重复对象时，建议再用 LLM vision 审阅。
- 生成的 HTML 目标是忠实、可检查、可复用的预览和风格底座，不是像 PowerPoint 渲染器一样做到像素级完全一致。

## 隐私说明

默认 `.gitignore` 会排除源 PPTX 和生成输出。不要把私人汇报、未发表图片或敏感科研材料提交到仓库，除非你明确希望公开它们。
