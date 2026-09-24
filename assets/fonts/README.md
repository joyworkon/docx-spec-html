# 字体资源说明

## JINGDONGLangZhengTi1-Bold.woff2

Hero 大标题专用字体（京东朗正体），由生成器内嵌为 data-URI 随页面分发。

⚠️ 这个文件是**整字库**。它只能作为子集化的输入，绝不能原样嵌进成品页面：
`scripts/optimize_page_assets.py` 会按页面实际用字重新子集化，实测把 7.66 MB
的原始字形压到约 70 KB。历史上有成品页面直接嵌了未裁剪的字库，单文件因此涨到
55 MB，在第三方 HTML 编辑器里根本打不开。

## misans-src/（MiSans Normal / Bold）

规范页正文与强调文字使用的 MiSans 字体源文件（TTF，取自小米官方发布的 MiSans
字体包）。**只使用两个字重**：

| 文件 | 注册为 | 用途 |
| --- | --- | --- |
| `MiSans-Normal.ttf` | `font-family: "MiSans"; font-weight: 400` | 正文、列表内容、说明文字 |
| `MiSans-Bold.ttf` | `font-family: "MiSans"; font-weight: 700` | `{ 标题 }`、`INTRODUCTION`、红/灰方块标题、绿框文字、表头、按钮文字、外链文字 |

`MiSans-Heavy.ttf` 保留在目录里只为兼容旧页面，新页面不再嵌入它。

关键约定：

- 两个字重**各只嵌一份**，都挂在同一个 `MiSans` 家族上（按 400 / 700 区分）。
  不要再额外注册 `MiSans-Normal` / `MiSans-Bold` 这样的别名字体面——同一份字形
  嵌两次白白多出约 0.5 MB。
- 样式表不直接写字体名，只用 `var(--font-normal)` / `var(--font-bold)`（在
  `assets/styles.css` 的 `html` 上声明），所以改字体只需改这两个变量。
- 嵌入由两条路径完成：`scripts/batch_generate.py`（DOCX 草稿）和
  `scripts/optimize_page_assets.py`（**所有**页面，包括 PDF-only 的模型主导产出）。
  后者是必跑步骤——PDF-only 路径以前从不调用字体嵌入函数，页面里只有
  `font-family: "MiSans"` 的名字引用，接收方没装字体就静默回退成系统字体。
- 默认子集覆盖 `--font-coverage common`：页面实际用字 ∪ GB2312 一级字（3755 字）。
  这样审稿人在别的编辑器里**新打的中文也还是 MiSans**，而不是只有原文能正常显示。
  实测两个字重各约 0.48–0.50 MB。需要极限压缩时用 `--font-coverage page`（只裁到
  页面现有字符），需要完整字库时用 `--font-coverage full`。
- 需要 `pip install fonttools brotli`；缺 fontTools 或缺少源 TTF 时，生成器静默
  回退为按字体名引用（本机已装 MiSans 的用户不受影响，接收方会回退系统字体）。
- MiSans 字体的著作权与再分发条款以小米官方的 MiSans 许可声明为准；更新字体版本
  时替换本目录下对应 TTF 即可，文件名保持 `MiSans-{Normal,Bold}.ttf`。
