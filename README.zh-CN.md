# personal-style-ppt2html

[English](README.md) | 中文

`personal-style-ppt2html` 是一个 Codex skill 和 PPTX 转 HTML 工具集，用来把用户已有 PowerPoint 中的个人风格、单位风格和学术汇报风格转成可复用的 HTML presentation。

它适合那些已经有自己常用 `.pptx` 模板、组会汇报风格、单位 logo 横条、论文图展示习惯的人。目标不是生成花哨网页模板，而是保留严肃、克制、可复用的个人学术表达系统。

## 这个仓库能做什么

- 把 `.pptx` 转成可翻页的 HTML slide preview。
- 保留 PPTX 中可复用的 slide-layout 元素，例如单位 logo、顶部横条、底部规则线、页眉页脚视觉标识。
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
- `ppt/theme/theme1.xml`：主题色；
- `ppt/media/*`：PPTX 内嵌图片。

这一步很关键，因为 logo 和顶部横条经常不在每一页的 `slideN.xml` 里，而是在 `slideLayout` 里。现在转换器会顺着 slide 到 layout 的关系，把这些可复用元素一起插入 HTML。

生成的 HTML 还可以带 `data-motion` 动效预设。动效只允许短促的透明度和轻微位移动画，并且包含 reduced-motion 和 print 降级规则，所以既适合学术汇报，也适合录屏展示。

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
│   └── institution-brand-rules.md
└── scripts/
    ├── audit_html_layout.py
    ├── build_theme_css.py
    ├── extract_pptx_style.py
    ├── make_asset_manifest.py
    ├── pptx_to_academic_html.py
    └── test_academic_html_tools.py
```

## 快速开始

使用 Python 3。核心 XML 解析路径不依赖第三方包。

```bash
python scripts/extract_pptx_style.py your-deck.pptx -o work/style-profile.json
python scripts/make_asset_manifest.py your-deck.pptx -o work/reference-assets
python scripts/build_theme_css.py work/style-profile.json -o work/theme.generated.css
python scripts/pptx_to_academic_html.py your-deck.pptx -o work/html-preview --profile work/style-profile.json --motion recording
python scripts/audit_html_layout.py work/html-preview/index.html
```

然后打开：

```text
work/html-preview/index.html
```

生成的 HTML 支持键盘翻页、reduced-motion 偏好，也支持打印或导出 PDF。

## 动效预设

使用 `--motion none` 可以得到最保守的学术输出。`--motion subtle` 适合现场浏览器汇报，`--motion recording` 适合录屏讲解，`--motion demo` 只建议用于更公开、更展示型的视频片段。

动效系统会避免循环装饰动画，保持 slide-layout 中的 logo 和横条稳定，并且在 reduced-motion 用户设置和打印输出中自动禁用动画。

## 作为 Codex Skill 安装

把仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.codex/skills/personal-style-ppt2html
```

之后可以这样调用：

```text
Use $personal-style-ppt2html to convert my reference PPTX into an academic HTML deck.
```

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
- 带 slide-layout logo 和顶部横条的 PPTX 转 HTML；
- 动效预设输出、图片优化属性和动画安全警告。

## 当前限制

- `.emf` 和 `.wmf` 不是浏览器原生支持的图片格式。如果本机没有 LibreOffice、Inkscape、ImageMagick 等转换工具，这些资源会显示为占位。
- 复杂 PowerPoint 几何形状会被近似处理。简单色块、横条、图片和文本框效果最好。
- 当前重点处理 slide 和 slide-layout 内容。如果某些模板把可复用视觉元素只放在 slide-master 层，还可以继续扩展完整 master 继承。
- 生成的 HTML 目标是忠实、可检查、可复用的预览和风格底座，不是像 PowerPoint 渲染器一样做到像素级完全一致。

## 隐私说明

默认 `.gitignore` 会排除源 PPTX 和生成输出。不要把私人汇报、未发表图片或敏感科研材料提交到仓库，除非你明确希望公开它们。
