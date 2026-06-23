# 视觉设计规范与 HTML 生成协议（Design + Generation Specification）

> 本文档基于「家庭服务机器人 商品信息运营规范」页面提取，用于指导后续页面的视觉一致性，并作为 Word 转单文件 HTML 的执行规范。
> 最后更新：2026-06-01

---

## 零、生成使用原则

本规范用于生成同系列运营说明类 HTML 页面。生成时先复用页面骨架与组件系统，再替换行业、章节与业务内容，避免重新发明视觉风格。

### 0.1 页面定位

- 适用：品类运营规范、商品信息规范、页面结构说明、策略对比、整改示例、标准化流程说明。
- 不适用：营销落地页、商品详情页实稿、移动端自适应应用、需要真实交互的后台系统。
- 视觉关键词：固定画布、红色品牌头图、白色大卡片、灰色内容面板、结构化示例、轻装饰、无阴影。

### 0.2 生成优先级

1. **先保证结构一致**：`main.poster` → `section.hero` → 多个 `section.card`。
2. **再保证组件一致**：标题栏、灰色面板、红色列表、示例网格优先复用既有 class。
3. **最后才微调内容高度**：优先增加卡片高度和内容块，不随意缩小字号、压缩行高或改变画布宽度。

### 0.3 HTML 生成硬约束

- 文档语言使用 `<html lang="zh-CN">`。
- viewport 使用固定宽度：`<meta name="viewport" content="width=1280">`。
- 页面主容器必须是 `<main class="poster">`，宽度固定 `1280px`，不做流式布局。
- 每个一级内容区必须使用 `<section class="card">` 或 `<section class="card spec-card">`。
- Hero 只能出现一次，位于页面顶部；章节卡片按业务顺序向下排列。
- 避免在 HTML 中写大量 `style=""` 覆盖。参考文件里的少量内联样式属于人工微调，生成新页面时应优先通过 CSS class 解决。
- 装饰元素必须加 `aria-hidden="true"`，不要把纯视觉装饰暴露为正文内容。

### 0.4 输出封装规则

- 单页交付优先生成 **单文件 HTML**：CSS 写入 `<style>`，不再依赖外部 CSS 文件。
- 标题字体文件可用时必须用 `@font-face` 内嵌 `JINGDONGLangZhengTi1-Bold.ttf`，保证离线打开也能渲染正确标题字体。
- 如需保证离线可打开，图片使用 `data:` URI 内嵌；如图片总量过大，可改为 `assets/` 目录外链，但必须保持相对路径稳定。
- 批量页面可共用外部 CSS；但交付给非开发用户预览时，优先输出单文件 HTML。
- HTML 中不得残留生成说明、修订说明或“已删除某某”等非原文提示，除非用户明确要求展示。

### 0.5 内容组织规则

- 页面标题建议两行以内；如必须换行，用 `<br>` 主动控制断行。
- **固定断行规则：** 当 Hero 主标题包含品类后缀「商品信息运营规范」时，必须在「商品信息运营规范」前强制插入 `<br>`，使「商品信息运营规范」固定落到第二行，前缀品类词（如「汽车汽机油」）独占第一行。例如 `汽车汽机油 商品信息运营规范` → `<h1>汽车汽机油<br>商品信息运营规范</h1>`。
- 每张卡片只表达一个主题：综述、主图、标题、卖点、系列品、详情页等都应独立成卡。
- 规范文本放入 `.gray-panel.spec-text`；案例、对比、流程、表格化信息放入 `.gray-panel.example-block`。
- 一段说明超过 3 个要点时，用 `.red-list`；同一要点下的二级说明用 `.sub-dot`。
- 灰色面板内的内容小标题（如“产品主图（第一张）：”“短标题”“长标题”“示例如下：”）统一使用 `.label-line`，样式为红色方块 + 荧光条。
- `.label-line` 下一级条目统一使用 `.source-list`，整体比上级标题右缩进 `28px`；圆点降级为灰色方块，条目内重点词只加粗，不加荧光条。
- 示例区优先使用网格组件表达，不要堆长段落。
- 连续 5 个及以上“第 N 屏/各屏展示/详情页屏次”示例必须使用 `.detail-screen-grid`：默认一行两张；横向长图、关键大图或单图说明可用一行一张，不得四列挤压或全部竖向罗列。

### 0.6 Word 转 HTML 内容保真规则

- 以 Word 原有层级为准：标题、章节、子标题、列表、表格、图片按原顺序进入 HTML。
- 除用户指定删除的文字外，所有可见文字必须完整保留，包括标点、括号、大小写、数字、英文缩写和冒号后的内容。
- 为适配视觉组件可以拆分段落，但拆分后拼接出的可见文本必须仍能还原原文。
- 所有图片必须保留；如果 Word 中同一图片出现多次，HTML 中也应保留对应出现次数。
- 生成后做两项校验：
  - 文本校验：提取 Word 可见段落，删除用户指定删词后，逐段确认在 HTML 可见文本中存在。
  - 图片校验：统计 Word 图片引用次数和 HTML `<img>` 数量，确认一致。

### 0.7 禁止项

- 不使用 `box-shadow`、玻璃拟态、渐变卡片、悬浮投影或复杂纹理背景。
- 不使用默认列表样式；所有列表必须使用红色或灰色方形圆点。
- 不使用负字间距；中文、英文、数字默认 `letter-spacing: 0`。
- 不把卡片嵌套在卡片里；灰色面板内可放白色组件块。
- 不新增紫色、蓝紫、金色、棕橙等强主题色；除 Hero 红色渐变外，主体保持红、白、灰、黑。
- 不为了塞内容随意降低正文到 `16px` 以下；内容过多时增加卡片高度或拆分卡片。

---

## 一、色彩体系

### 1.1 主色（Brand Primary）

| 用途 | 色值 | 说明 |
|------|------|------|
| 品牌红 | `#FF2B22` | 核心品牌色，用于标题、强调标记、装饰线、列表圆点 |
| 品牌红深 | `rgba(255, 48, 38, 0.98)` | Hero 区渐变起始色 |
| 品牌红更深 | `rgba(255, 30, 22, 0.96)` | Hero 区渐变终止色 |

**品牌红使用场景：**
- 页面大标题（`h2`）文字色
- 章节编号（如 "01"）文字色
- 英文标签（`en-label`）文字色及装饰线
- 列表圆点背景色
- VS 对比标签（中高端 & 低端）
- 标题示例标签背景色
- 卖点编号背景色

### 1.2 背景色

| 用途 | 色值 | 说明 |
|------|------|------|
| 页面底色 | `#DCEDFF` | 淡蓝色，全局背景 |
| 卡片底色 | `#FFFFFF` | 白色，内容承载区 |
| 灰色面板 | `#F1F1F1` | 浅灰，内容分组背景 |
| 纯白面板 | `#FFFFFF` | `lead` 段落背景、VS 列/标题内容区 |

### 1.3 文字色

| 层级 | 色值 | 用途 |
|------|------|------|
| 主文字 | `#333333` | 正文、段落、列表文字 |
| 次要文字 | `#555555` | 帧描述、公式文字 |
| 深黑强调 | `#111111` | 特殊强调标题、VS 标题、卖点名称 |
| 白色文字 | `#FFFFFF` | Hero 区、标签、深色按钮上的文字 |

### 1.4 功能色

| 用途 | 色值 | 说明 |
|------|------|------|
| 超高端标签 | `#111111` | 深色背景标签 |
| 链接蓝 | `#0A72C4` | 标注引线色 |
| 高亮红条 | `rgba(255,43,34,0.2)` | 标题背后的半透明品牌红高亮条（替代旧的浅粉 #FFD4D4） |
| 副圆点灰 | `#D8D8D8` | 次级列表圆点背景色 |
| 占位符灰 | `#8A8A8A` | 占位符文字色、元数据色 |
| 价格红 | `#FF2B22` | 产品价格文字 |

### 1.5 渐变

| 名称 | 值 | 用途 |
|------|-----|------|
| Hero 主渐变 | `linear-gradient(102deg, rgba(255,48,38,0.98), rgba(255,30,22,0.96))` | Hero 区背景 |
| Hero 装饰光斑 | `radial-gradient(circle at 78% 22%, rgba(255,255,255,0.28) 0 2px, transparent 3px 100%)` | Hero 右上角光点 |
| 标题高亮条 | `linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%)` | 列表加粗标题底部半透明品牌红底，四角圆角 `3px`（与列表圆点一致），水平内边距 `4px` |
| 小标题高亮条 | `linear-gradient(transparent 62%, rgba(255,43,34,0.2) 62%)` | `h3` 标题底部半透明品牌红底，四角圆角 `3px`（与列表圆点一致），水平内边距 `4px` |
| 长标题示例高亮 | `linear-gradient(transparent 60%, rgba(255,43,34,0.2) 60%)` | 标题示例文字底部半透明品牌红底，四角圆角 `3px`（与列表圆点一致），水平内边距 `4px` |
| 轮廓线纹理 | `repeating-linear-gradient(0deg, rgba(255,255,255,0.45) 0 2px, transparent 2px 22px)` + `repeating-linear-gradient(90deg, ...)` | Hero 斜纹纹理 |
| VS 场景图 | `linear-gradient(120deg, #b9c8d3, #e0e8ee 55%, #aab8c4)` | 超高端场景化占位 |
| VS 参数图 | `linear-gradient(135deg, #f4f4f4, #d0d0d0)` | 中高端参数化占位 |

---

## 二、字体与排版

### 2.1 字体族

**Hero 主标题字体：**
```css
@font-face {
  font-family: "JINGDONGLangZhengTi1-Bold";
  src: url("JINGDONGLangZhengTi1-Bold.ttf") format("truetype");
  font-weight: 700;
  font-style: normal;
}

font-family: "JINGDONGLangZhengTi1-Bold", "jingdonglangzhengti1", "Microsoft YaHei", "PingFang SC", sans-serif;  /* h1 标题 */
font-weight: 700;
```

**Hero 区其他文字：**
```css
font-family: "Arial", "Microsoft YaHei", "PingFang SC", sans-serif;  /* 其他 */
```

**卡片区（非头部）：**
```css
/* 章节大标题 */
font-family: "MiSans-Bold", "MiSans", sans-serif; /* h2 */
font-weight: 700;

/* 章节编号 */
font-family: "MiSans-Heavy", "MiSans", sans-serif;    /* chapter */
font-weight: 900;

/* 其他标题 */
font-family: "MiSans-Semibold", "MiSans", sans-serif; /* h3, h4, 标签, 副标题等 */
font-weight: 600;

/* 正文内容 */
font-family: "MiSans-Normal", "MiSans", sans-serif;   /* .card 下所有正文 */
font-weight: 400;
```

- Hero 主标题使用 `JINGDONGLangZhengTi1-Bold`
- 卡片章节大标题 `h2` 固定使用 `MiSans-Bold`
- 卡片区正文与其他组件使用 MiSans 字体族，按层级分 Heavy/Semibold/Normal 三档

### 2.2 字号体系

| 层级 | 字号 | 字重 | 行高 | 用途 |
|------|------|------|------|------|
| Hero 标题 `h1` | `68px` | `700` / `JINGDONGLangZhengTi1-Bold` | `1.12` | 页面主标题 |
| 章节标题 `h2` | `47px` | `700` / `MiSans-Bold` | `1.0` | 卡片内大标题 |
| 示例标题 `h3` | `27px` | `600` | `1.0` | 灰色面板内小标题 |
| 引导段落 `lead` | `19px` | `600` | `1.55` | 综述段落 |
| 列表加粗 `b` | `19px` | `600` | `1.4` | 列表项标题 |
| 列表正文 `p` | `19px` | `normal` | `1.38` | 列表项描述 |
| 更新日期 | `18px` | `800` | — | Hero 区右下角日期，中文格式 |
| 英文标签 | `15px` | `600` | `1.0` | 区块英文副标题 |
| Hero 标记 | `14px` | `800` | `1.25` | Hero 右上角"OPERATION STANDARDS" |
| 章节编号 | `17px` | `900` | — | 如 "01"、"02" |
| 帧编号 | `14px` | `600` | — | 如 "第 1 帧"、"第 5–10 帧" |
| 帧标题 | `19px` | `600` | `1.2` | 帧展示卡片标题 |
| 帧描述 | `14px` | `400` | `1.4` | 帧展示卡片描述 |
| VS 列标题 | `22px` | `600` | `1.2` | VS 对比列标题 |
| VS 列描述 | `17px` | `400` | `1.45` | VS 对比列正文 |
| VS 图片标签 | `18px` | `600` | — | 占位图内文字 |
| VS 标签 | `16px` | `600` | — | 超高端/中高端标签 |
| VS 分隔符 | `36px` | `600` | — | "VS" 文字 |
| 标题示例标签 | `22px` | `600` | `1.2` | 短标题/长标题标签 |
| 标题公式 | `17px` | `600` | `1.4` | 公式行 |
| 标题示例文字 | `19px` | `600` | `1.45` | 示例标题文字 |
| 标题元数据 | `15px` | `400` | — | 字数/字符数说明 |
| 卖点名称 | `24px` | `600` | — | 基础卖点/产品性能 |
| 卖点列表项 | `18px` | `400` | `1.4` | 卖点详情 |
| 整改标签 | `22px` | `600` | — | 整改前/整改后 |
| 整改单元格 | `14px` | `600` | `1.25` | 产品规格文字 |
| 整改价格 | `16px` | `600` | — | 价格文字 |
| 整改备注 | `14px` | `400` | — | 说明文字 |
| 屏次编号 | `14px` | `600` | — | 如 "前置模块"、"第 1–2 屏" |
| 屏次标题 | `19px` | `600` | `1.25` | 屏次卡片标题 |
| 屏次描述 | `14px` | `400` | `1.45` | 屏次卡片描述 |

### 2.3 字重规范

| 名称 | 数值 | 使用场景 |
|------|------|----------|
| 标题 Bold | `700` | Hero 标题 `h1`，必须绑定 `JINGDONGLangZhengTi1-Bold` |
| 章节标题 Bold | `700` | 卡片章节大标题 `h2`，固定使用 `MiSans-Bold` |
| 极粗 | `900` | 章节编号（chapter） |
| 粗体 | `800` | 更新日期、Hero 标记 |
| 半粗 | `600` | 其他标题（h3、h4、标签、副标题、列表标题、示例标题等） |
| 常规 | `normal` / `400` | 正文段落 |

> **规则：** Hero 主标题使用 `JINGDONGLangZhengTi1-Bold` 字体文件；卡片章节大标题固定为 `MiSans-Bold`。其他字重分为 `900`（章节编号）/ `800`（Hero 特殊）/ `600`（标题与强调）/ `400`（正文）。不再使用旧的统一 `900` 规范。

### 2.4 字间距

- 所有文本 `letter-spacing: 0`，不做额外字间距调整
- 帧编号特殊：`letter-spacing: 0.08em`
- 屏次编号特殊：`letter-spacing: 0.06em`
- VS 分隔符特殊：`letter-spacing: 0`

---

## 三、间距与尺寸

### 3.1 页面级尺寸

| 属性 | 值 | 说明 |
|------|-----|------|
| 页面宽度 | `1280px` | 固定宽度居中 |
| 卡片宽度 | `1188px` | 内容区宽度 |
| 安全边距 | `(1280-1188)/2 = 46px` | 页面两侧留白 |
| viewport | `width=1280` | HTML meta 固定画布宽度 |
| 主容器溢出 | `overflow: hidden` | 防止 Hero 装饰超出画布 |

### 3.2 Hero 区

| 属性 | 值 |
|------|-----|
| 高度 | `520px` |
| 内边距 | `72px 44px 0` |
| 标题字号 | `68px` |
| 标题最大行数 | 2 行以内 |
| 标题下分隔线宽度 | `118px` |
| 标题下分隔线高度 | `2px` |
| 分隔线与标题间距 | `124px`（`margin-top`） |
| 更新日期位置 | `right: 50px, bottom: 116px` |

### 3.3 卡片间距

| 属性 | 值 | 说明 |
|------|-----|------|
| 卡片圆角 | `34px` | 大圆角，现代感 |
| 卡片内边距 | `68px 42px 58px` | 上/左右/下 |
| 卡片上偏移 | `-75px` | 与 Hero 区重叠 |
| 卡片间距 | `38px`（`margin-top`） | 多卡片之间 |
| 规范卡片上内边距 | `62px` | `spec-card` 专用 |
| 规范卡片下内边距 | `48px` | `spec-card` 专用 |

### 3.4 灰色面板

| 属性 | 值 | 说明 |
|------|-----|------|
| 圆角 | `16px` | 中等圆角 |
| 内边距（spec） | `30px 28px 24px` | 规范面板 |
| 内边距（example） | `24px 28px 28px` | 示例面板 |
| 规范面板与标题间距 | `30px` | `margin-top` |
| 示例面板与规范面板间距 | `50px` | `margin-top` |

### 3.5 内嵌白面板

| 属性 | 值 | 说明 |
|------|-----|------|
| `lead` 段落圆角 | `10px` | 小圆角 |
| `lead` 段落内边距 | `24px 34px` | 上下/左右 |
| `lead` 与列表间距 | `22px` | `margin-bottom` |

### 3.6 列表项间距

| 属性 | 值 | 说明 |
|------|-----|------|
| 列表项左侧缩进 | `25px` | `padding-left` |
| 列表项之间间距（intro） | `16px` | `margin-bottom` |
| 列表项之间间距（spec） | `18px` | `margin-bottom` |
| 列表标题与正文间距 | `3px` | `margin-top` of `p` |

### 3.7 子级缩进（sub-dot）

| 属性 | 值 |
|------|-----|
| 左侧缩进 | `28px` |
| 圆点大小 | `11px × 11px` |
| 圆点圆角 | `3px` |
| 圆点位置 | `left: 0, top: 10px` |

---

## 四、圆角体系

| 层级 | 圆角值 | 使用场景 |
|------|--------|----------|
| 超大圆角 | `34px` | 卡片 |
| 大圆角 | `16px` | 灰色面板 |
| 中圆角 | `10px` | 内嵌白面板、VS 列、帧卡片、标题内容区、屏次卡片 |
| 小圆角 | `8px` | VS 标签、卖点编号、整改标签、屏次占位图 |
| 极微圆角 | `3px` | 列表圆点、荧光粉高亮条（标题/小标题/标题示例）、帧占位图、整改单元格 |

**圆角递进规则：** 外层 > 内层，从 `34px` → `16px` → `10px` → `8px` → `3px` 逐级递减。

---

## 五、装饰元素

### 5.1 列表圆点

- **主圆点：** `11px × 11px`，圆角 `3px`，背景 `#FF2B22`，定位在 `left: 0, top: 9px`
- **副圆点（sub-dot）：** `11px × 11px`，圆角 `3px`，背景 `#D8D8D8`，定位在 `left: 0, top: 10px`

### 5.2 英文标签箭头（en-label span）

- 形态：统一使用“线身 + 弯钩箭头”的一体式 SVG 箭头（横线右侧/左侧带一个上扬的弧形箭头钩），不再使用旧的“横线 + 单斜线”两段式装饰。
- 实现：箭头作为 `.en-label span` 的 `background` 内嵌 SVG（`data:image/svg+xml`），`background: ... no-repeat center / contain`，不使用 `::before` 伪元素。
- 尺寸：宽 `67px`，高 `6px`（与 SVG 原始比例 `138.209 : 12.5 ≈ 11.06 : 1` 一致，等比缩放，不拉伸变形）
- 颜色：品牌红 `#FF2B22`（SVG `fill`）
- 方向：箭头钩朝**左**（SVG 内部 `transform: translate(138.209,0) scale(-1,1)` 水平翻转）
- 间距：`margin-left: auto`（右对齐）
- 右对齐：整个 `.en-label` 宽度 `192px`，`text-align: right`
- 布局：`display: flex; flex-direction: column; justify-content: space-between; height: 47px`
  - 英文文字顶部对齐中文标题顶部
  - 箭头底部对齐中文标题底部

### 5.3 Hero 分隔箭头（hero-rule）

- 形态：与 5.2 同一套 SVG 箭头，方向相反（与红色箭头互为水平翻转）。
- 实现：箭头作为 `.hero-rule` 的 `background` 内嵌 SVG，`background: ... no-repeat center / contain`，不使用 `::after` 伪元素。
- 尺寸：宽 `118px`，高 `11px`（与红色箭头保持同一 SVG 等比例，仅整体放大，不拉伸变形）
- 颜色：白色 `#FFFFFF`（SVG `fill`）
- 方向：箭头钩朝**右**（使用 SVG 原始朝向，不翻转）

### 5.4 Hero 标记胶囊（hero-mark）

- 尺寸：`155px × 62px`
- 边框：`1px solid rgba(255,255,255,0.95)`
- 上下边框透明：`border-top-color: transparent`, `border-bottom-color: transparent`（仅左右竖线）
- 圆角：`32px`（胶囊形）
- 定位：`top: 70px, right: 48px`

### 5.5 机器人装饰（Hero 背景）

- **圆环组**：3 个大小不同的圆环，模拟机器人主体轮廓
  - 圆环 1：`360px × 360px`，定位 `right: -60px, top: -80px`
  - 圆环 2：`220px × 220px`，定位 `right: 280px, top: 60px`
  - 圆环 3：`160px × 160px`，定位 `right: 100px, top: 280px`
  - 边框：`3px solid rgba(255,255,255,0.44)`
  - 内圈 `::before`：`inset: 26px`（圆环 1）/ `18px`（圆环 2）/ `14px`（圆环 3），边框 `2px solid rgba(255,255,255,0.3)`
  - 十字传感器 `::after`：`60px × 60px`，白色十字线，`rgba(255,255,255,0.45)`

### 5.6 扫地路径线条（Hero 背景）

- 高度：`2px`
- 颜色：`rgba(255,255,255,0.42)`
- 4 条路径线不同位置和微旋转，模拟机器人清扫路径
  - 路径 A：`left: 30px, right: 600px, top: 220px, rotate(-3deg)`
  - 路径 B：`left: 70px, right: 700px, top: 280px, rotate(2deg)`
  - 路径 C：`left: 120px, right: 350px, top: 360px, rotate(-2deg)`
  - 路径 D：`left: 40px, right: 720px, top: 420px, rotate(1deg)`

### 5.7 斜纹纹理（Hero ::before）

- 不透明度：`0.18`
- 斜光束：`linear-gradient(85deg, transparent 0 12%, rgba(255,255,255,0.9) 12.2%, transparent 12.6% 100%)`
- 水平线：`repeating-linear-gradient(0deg, rgba(255,255,255,0.45) 0 2px, transparent 2px 22px)`
- 垂直线：`repeating-linear-gradient(90deg, rgba(255,255,255,0.35) 0 2px, transparent 2px 28px)`
- 变形：`skewY(-4deg) scale(1.08)`

---

## 六、组件规范

### 6.1 卡片（Card）

```
宽度: 1188px
圆角: 34px
背景: #FFFFFF
内边距: 68px 42px 58px
上偏移: -75px（与 Hero 重叠，形成层叠效果）
z-index: 2
```

### 6.2 区块标题栏（Section Head）

```
布局: flex, align-items: flex-start, justify-content: space-between
最小高度: 86px
左侧: h2 标题（#FF2B22, 47px, MiSans-Bold, weight 700）
右侧: 英文标签（192px 宽, 右对齐）
  - 英文文字: 固定为 INTRODUCTION，#FF2B22, 15px, weight 600
  - 箭头: 一体式 SVG 弯钩箭头（线身 + 右侧弧形箭头钩），67px × 6px, fill #FF2B22, margin-left: auto，作为 background 内嵌，等比不拉伸
  - 箭头方向: 钩朝左（SVG 水平翻转），与 Hero 白色箭头互为反向
  - 高度: 47px（= h2 font-size），英文顶对齐、箭头底对齐
  - justify-content: space-between
```

**带章节编号变体（spec-head）：**
```
布局: flex, justify-content: flex-start, gap: 16px
章节编号: #FF2B22, 17px, weight 900, padding-top 0
英文标签: margin-left: auto（推至最右），文字固定为 INTRODUCTION
```

### 6.3 灰色面板（Gray Panel）

```
背景: #F1F1F1
圆角: 16px
内边距: 28px~30px 28px 24px
```

### 6.4 引导段落（Lead）

```
背景: #FFFFFF
圆角: 10px
内边距: 24px 34px
字号: 19px
字重: 600
行高: 1.55
下边距: 22px
```

### 6.5 红色列表（Red List）

```
无默认列表样式（list-style: none）
每项左侧缩进: 25px
圆点: 11×11px, 圆角3px, 背景#FF2B22, 定位left:0 top:9px
标题(b): 19px, weight 600, 底部半透明品牌红高亮条(linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%), border-radius 3px, padding 0 4px)
正文(p): 19px, 行高1.45, 颜色#333，作为标题下一级，继承父级缩进后再右移8px，左侧灰色方块
项间距: 16~18px
```

### 6.6 5 帧展示（Frames Grid）

```
布局: grid, 5列等分
列间距: 12px
帧卡片: 白底, 圆角10px, 内边距16px 14px 18px
帧编号: 14px, weight 600, 颜色#FF2B22, letter-spacing 0.08em
帧标题: 19px, weight 600, 颜色#111
帧描述: 14px, weight 400, 颜色#555
帧占位图: 高110px, 圆角6px, 背景#E8E8E8, 文字14px #8A8A8A
```

### 6.7 VS 对比展示（VS Grid）

```
布局: grid, 3列（1fr 80px 1fr）
列间距: 16px
VS 列: 白底, 圆角10px, 内边距22px, 最小高度280px
VS 标签: 自适应宽度, 圆角6px, 内边距6px 14px
  - "超高端": 背景#111, 颜色#FFF, 16px, weight 600
  - "中高端 & 低端": 背景#FF2B22, 颜色#FFF, 16px, weight 600
VS 分隔符: 居中, 36px, weight 600, 颜色#FF2B22
VS 占位图: 高140px, 圆角8px
  - 场景化: 渐变120deg, #b9c8d3 → #e0e8ee → #aab8c4
  - 参数化: 渐变135deg, #f4f4f4 → #d0d0d0
第二个 VS 区间距: margin-top 24px
```

### 6.8 标题示例展示（Title Demo）

```
布局: grid, 2列（200px 1fr）
列间距: 18px
标签列: 背景#FF2B22, 圆角10px, 居中, 22px, weight 600
  - 小字: 14px, weight 400, letter-spacing 0.06em, opacity 0.92
内容列: 白底, 圆角10px, 内边距20px 24px
公式行: 17px, weight 600, 颜色#555, 关键词用 em 标记为#FF2B22
示例文字: 19px, weight 600, 颜色#111, 底部半透明品牌红高亮(transparent 60%, rgba(255,43,34,0.2) 60%, border-radius 3px, padding 0 4px)
元数据: 15px, weight 400, 颜色#8A8A8A
两个标题示例间距: 18px
```

### 6.9 卖点优先级展示（Selling Grid）

```
布局: grid, 2列等分
列间距: 18px
卖点列: 白底, 圆角10px, 内边距24px 26px
卖点编号: 44px × 44px, 圆角8px, 背景#FF2B22, 白色文字, 22px
卖点名称: 24px, weight 600, 颜色#111
卖点列表: 无默认样式, 每项内边距10px 0, 底部边框1px #F1F1F1
  - 列表项: 18px, weight 400, 颜色#333, 左侧缩进22px
  - 列表圆点: 9px × 9px, 圆角3px, 背景#FF2B22, 定位left:0 top:18px
  - 加粗词: weight 600, 颜色#111
```

### 6.10 系列品整改对比（Bind Compare）

```
布局: grid, 2列等分
列间距: 16px
对比列: 白底, 圆角10px, 内边距18px 22px 24px
整改标签: 高44px, 圆角8px, 居中, 22px, weight 600
  - "整改前": 背景#F1F1F1, 颜色#555
  - "整改后": 背景#FF2B22, 颜色#FFF
产品行: grid 3列, 间距10px
产品单元格: 背景#F1F1F1, 圆角6px, 内边距14px 8px, 居中, 14px weight 600
  - 价格: 颜色#FF2B22, 16px, weight 600
  - 备注: 14px, weight 400, 颜色#8A8A8A, 居中
```

### 6.11 商详屏次展示（Screens Grid）

```
布局: grid, 4列等分
列间距: 14px
屏次卡片: 白底, 圆角10px, 内边距16px 16px 18px
屏次编号: 14px, weight 600, 颜色#FF2B22, letter-spacing 0.06em
屏次标题: 19px, weight 600, 颜色#111, 行高1.25
屏次描述: 14px, weight 400, 颜色#555, 行高1.45
屏次占位图: 高200px, 圆角6px, 虚线边框1px dashed #CFCFCF
  - 默认: 渐变180deg, #f7f7f7 → #e4e4e4
  - 场景变体(.scene): 渐变180deg, #e4dccc → #f5edd9, 边框#c9b890
  - 对比变体(.compare): 横线纹理, 背景#fff, 边框#CFCFCF
```

---

## 七、阴影体系

| 层级 | 值 | 使用场景 |
|------|-----|----------|
| 无阴影 | — | 卡片、面板、所有布局元素（依赖色块层级区分） |

> **规则：** 家庭服务机器人页面不使用 `box-shadow`，仅通过背景色层级和白/灰对比区分层次。

---

## 八、边框体系

| 类型 | 值 | 使用场景 |
|------|-----|----------|
| 分隔线 | `2px solid` | Hero 分隔线、装饰线 |
| 卖点列表分隔 | `1px solid #F1F1F1` | 卖点列表项底部分隔 |
| 占位虚线 | `1px dashed #CFCFCF` | 屏次占位图边框 |

---

## 九、z-index 层级

| 层级 | 值 | 使用场景 |
|------|-----|----------|
| 基底 | `auto` | 灰色面板、内容 |
| 内容层 | `1` | Hero 区文字、分隔线 |
| 卡片层 | `2` | 白色卡片（覆盖 Hero 底部） |
| 装饰层 | `auto`（不设 z-index） | 机器人圆环、路径线等纯装饰 |

---

## 十、响应式策略

```css
@media (max-width: 1280px) {
  .poster {
    transform-origin: top left;
    /* 建议配合 JS 做等比缩放：scale(viewportWidth / 1280) */
  }
}
```

> **当前策略：** 固定 1280px 宽度，小屏通过 `transform: scale()` 等比缩放，不做流式重排。

---

## 十一、完整 CSS 变量建议

为方便后续页面复用，建议将核心值提取为 CSS 自定义属性：

```css
:root {
  /* 品牌色 */
  --brand-red: #FF2B22;
  --brand-red-light: #FF5D35;
  --brand-red-glow: rgba(255, 43, 34, 0.1);

  /* 背景色 */
  --bg-page: #DCEDFF;
  --bg-card: #FFFFFF;
  --bg-panel: #F1F1F1;

  /* 文字色 */
  --text-primary: #333333;
  --text-secondary: #555555;
  --text-emphasis: #111111;
  --text-inverse: #FFFFFF;

  /* 功能色 */
  --color-highlight: rgba(255,43,34,0.2);
  --color-muted: #D8D8D8;
  --color-placeholder: #8A8A8A;
  --color-price: #FF2B22;

  /* 圆角 */
  --radius-xl: 34px;   /* 卡片 */
  --radius-lg: 16px;   /* 面板 */
  --radius-md: 10px;   /* 内嵌白面板、VS列、帧卡片 */
  --radius-sm: 8px;    /* VS标签、卖点编号 */
  --radius-xs: 3px;    /* 圆点、高亮条 */

  /* 间距 */
  --space-page: 46px;     /* 页面安全边距 */
  --space-card-px: 42px;  /* 卡片水平内边距 */
  --space-card-pt: 68px;  /* 卡片顶部内边距 */
  --space-card-pb: 58px;  /* 卡片底部内边距 */
  --space-panel: 28px;    /* 面板内边距 */
  --space-gap-lg: 34px;   /* 大间距 */
  --space-gap-md: 18px;   /* 中间距 */
  --space-gap-sm: 12px;   /* 小间距 */

  /* 字号 */
  --font-hero: 68px;
  --font-h2: 47px;
  --font-h3: 27px;
  --font-body: 19px;
  --font-label: 15px;
  --font-caption: 14px;

  /* 字重 */
  --weight-title: 700;     /* JINGDONGLangZhengTi1-Bold 标题 */
  --weight-black: 900;     /* 章节编号 */
  --weight-semibold: 600;  /* 其他标题 */
  --weight-bold: 800;      /* Hero 特殊 */
  --weight-regular: 400;   /* 正文 */

  /* 页面宽度 */
  --width-page: 1280px;
  --width-card: 1188px;
}
```

---

## 十二、页面结构模板

```
┌─────────────────────────────────────────────┐
│              页面底色 (#DCEDFF)               │
│  ┌───────────────────────────────────────┐  │
│  │  HERO 区 (520px)                       │  │
│  │  渐变红背景 + 斜纹纹理 + 机器人装饰     │  │
│  │  · 右上角标记胶囊                       │  │
│  │  · 左侧主标题 (68px)                    │  │
│  │  · 分隔线 + 箭头                        │  │
│  │  · 右下角更新日期                        │  │
│  └───────────────────────────────────────┘  │
│       ┌─────────────────────────────┐       │
│       │  卡片 1：整体规范综述          │       │
│       │  · 区块标题栏               │       │
│       │  · （内容区）               │       │
│       └─────────────────────────────┘       │
│       ┌─────────────────────────────┐       │
│       │  卡片 2：主图（01）           │       │
│       │  · 区块标题栏 (带编号)        │       │
│       │  · 灰色面板 - 规范文本        │       │
│       │  · 灰色面板 - 5 帧展示        │       │
│       │  · 灰色面板 - VS 对比展示      │       │
│       └─────────────────────────────┘       │
│       ┌─────────────────────────────┐       │
│       │  卡片 3：标题（02）           │       │
│       │  · 区块标题栏 (带编号)        │       │
│       │  · 灰色面板 - 规范文本        │       │
│       │  · 灰色面板 - 标题示例展示     │       │
│       └─────────────────────────────┘       │
│       ┌─────────────────────────────┐       │
│       │  卡片 4：卖点（03）           │       │
│       │  · 区块标题栏 (带编号)        │       │
│       │  · 灰色面板 - 引导段落        │       │
│       │  · 灰色面板 - 卖点优先级展示    │       │
│       └─────────────────────────────┘       │
│       ┌─────────────────────────────┐       │
│       │  卡片 5：系列品工具（04）      │       │
│       │  · 区块标题栏 (带编号)        │       │
│       │  · 灰色面板 - 规范文本        │       │
│       │  · 灰色面板 - 整改前后对比     │       │
│       └─────────────────────────────┘       │
│       ┌─────────────────────────────┐       │
│       │  卡片 6：商品详情（05）        │       │
│       │  · 区块标题栏 (带编号)        │       │
│       │  · 灰色面板 - 规范文本        │       │
│       │  · 灰色面板 - 屏次展示        │       │
│       └─────────────────────────────┘       │
│                                              │
└─────────────────────────────────────────────┘
```

---

## 十三、组件选择规则

生成新页面时，不要只按视觉相似选择组件，应按信息类型选择组件。

| 信息类型 | 首选组件 | 使用条件 |
|----------|----------|----------|
| 页面总标题 | `.hero h1` | 页面唯一主命题，最多 2 行 |
| 更新时间/版本 | `.updated` | 放 Hero 右下角，格式建议 `更新日期 YYYY年MM月` |
| 章节标题 | `.section-head` / `.spec-head` | 综述卡不用编号；标准章节用 `01`、`02` 编号 |
| 规则说明 | `.gray-panel.spec-text` + `.red-list` | 适合承载标准、要求、原则、注意事项 |
| 关键引导语 | `.lead` | 每个面板最多 1 条，用来概括本节核心判断 |
| 面板内小标题 | `.label-line` | 适合“产品主图（第一张）：”“短标题”“长标题”“示例如下：”等灰色面板内的二级标题 |
| 小标题下级条目 | `.source-list` | `.label-line` 下面的内容，整体右缩进 `28px`，统一使用灰色方块，不再使用红点；条目重点词只加粗 |
| 多步骤/多帧结构 | `.frames-grid` | 3-5 个节点最合适；超过 5 个应拆为两行或拆卡 |
| 双路径对比 | `.vs-grid` | 适合高端 vs 普通、整改前 vs 整改后、方案 A vs B |
| 标题/公式示例 | `.title-demo` | 左侧放类型，右侧放公式、示例和限制 |
| 带标题图片卡片 | `.image-frame` + `.source-list` + `.doc-image` | 卡片内标题顶部对齐；图片保持原比例，在标题下方剩余区域上下左右居中 |
| 优先级分类 | `.selling-grid` | 适合 2 类并列优先级；超过 2 类可改 3 列或拆分 |
| 整改前后 | `.bind-compare` | 用于明显的新旧状态对比 |
| 页面屏次/流程 | `.screens-grid` | 适合 4 个以内模块；以手机屏次或流程节点为单位 |
| 孤立提示语 | `.example-line` | 适合“示例如下：”“注意：”等只有提示、没有配套内容的原文行 |

### 13.1 章节卡片结构

标准章节卡片应按以下顺序组织：

```html
<section class="card spec-card">
  <div class="section-head spec-head">
    <span class="chapter">01</span>
    <h2>{ 章节标题 }</h2>
    <div class="en-label">
      <strong>INTRODUCTION</strong>
      <span></span>
    </div>
  </div>

  <div class="gray-panel spec-text">
    <!-- 规则说明 -->
  </div>

  <div class="gray-panel example-block">
    <!-- 示例组件，可按需要增删 -->
  </div>
</section>
```

### 13.2 最小 HTML 骨架

```html
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=1280">
  <title>页面标题</title>
  <style>
    /* 单文件交付：在这里内嵌本规范 CSS */
  </style>
</head>
<body>
<main class="poster">
  <section class="hero">
    <div class="robot-deco" aria-hidden="true">
      <div class="ring ring-one"></div>
      <div class="ring ring-two"></div>
      <div class="ring ring-three"></div>
      <div class="path-line path-a"></div>
      <div class="path-line path-b"></div>
      <div class="path-line path-c"></div>
      <div class="path-line path-d"></div>
    </div>
    <div class="hero-mark">OPERATION<br>STANDARDS</div>
    <h1>品类前缀词<br>商品信息运营规范</h1>
    <div class="hero-rule"></div>
    <p class="updated">更新日期 2026年06月</p>
  </section>

  <section class="card intro-card">
    <!-- 综述内容 -->
  </section>

  <section class="card spec-card">
    <!-- 标准章节 -->
  </section>
</main>
</body>
</html>
```

### 13.3 内容长度控制

- `.red-list li` 标题建议 4-14 个字，正文建议 1-2 行。
- `.sub-dot` 适合补充说明，不适合承载完整段落；超过 4 条时拆成新的列表项。
- `.frames-grid` 单卡描述建议不超过 44 个中文字符。
- `.vs-col p` 建议不超过 2 行，长解释放到上方规则面板。
- `.screen-desc` 建议不超过 3 行，否则屏次卡高度会显得拥挤。
- 有下级内容的“示例如下：”使用 `.label-line`，其下内容降级为 `.source-list` 或示例网格；孤立的“示例如下：”“注意：”不应强行做成空卡片，用 `.example-line` 独立保留。

### 13.4 面板内层级样式

```html
<span class="label-line"><span class="label-text">产品主图（第一张）：</span></span>
```

```css
.label-line {
  display: inline-flex;
  align-items: flex-start;
  position: relative;
  margin: 0 0 10px;
  padding-left: 25px;
  color: #111;
  font-size: 19px;
  line-height: 1.4;
  font-weight: 600;
}
.label-line::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #FF2B22;
}
.label-line .label-text {
  display: inline;
  background: linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.source-list {
  list-style: none;
  margin: 0 0 0 28px;
  padding: 0;
}
.source-list li {
  position: relative;
  padding-left: 25px;
  margin: 12px 0 0;
  font-size: 19px;
  line-height: 1.45;
}
.source-list li::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #D8D8D8;
}
.source-list b {
  display: inline;
  color: #111;
  font-weight: 600;
}
.red-list li > p:not(.sub-dot) {
  position: relative;
  margin: 8px 0 0 8px;
  padding-left: 25px;
  line-height: 1.45;
}
.red-list li > p:not(.sub-dot)::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #D8D8D8;
}
.example-line {
  margin: 18px 0 0;
  color: #FF2B22;
  font-size: 18px;
  line-height: 1.4;
  font-weight: 600;
}
```

### 13.5 带标题图片卡片样式

```css
.title-image-grid,
.before-after-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px;
  align-items: stretch;
}
.title-image-grid .image-frame,
.before-after-grid .image-frame {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  min-height: 380px;
}
.image-frame .source-list {
  margin-bottom: 12px;
}
.title-image-grid .image-frame .doc-image,
.before-after-grid .image-frame .doc-image {
  margin: auto;
  width: auto;
  max-width: 100%;
  max-height: 320px;
  object-fit: contain;
}
```

> **规则：** 图片卡片里的标题永远固定在顶部，和同级卡片标题水平对齐；图片不能拉伸变形，必须保持原比例并在标题下方剩余区域上下左右居中。

---

## 十四、新增页面检查清单

生成新页面时，逐一检查以下项目确保视觉一致性：

- [ ] 页面宽度 `1280px`，底色 `#DCEDFF`
- [ ] HTML viewport 为 `width=1280`
- [ ] 单文件交付时 CSS 已内嵌，图片路径或 data URI 可离线打开
- [ ] Hero 区使用品牌红渐变 + 斜纹纹理 + 机器人装饰
- [ ] Hero 标题不超过 2 行，更新日期格式为中文 `更新日期 YYYY年MM月`
- [ ] Hero 标题若含「商品信息运营规范」，已在其前强制 `<br>` 断行，使「商品信息运营规范」固定为第二行
- [ ] Hero 右上角必须保留括号式标记 `.hero-mark`，文字为 `OPERATION` / `STANDARDS` 两行，不得省略
- [ ] 卡片 `1188px` 宽，`34px` 圆角，`-75px` 上偏移
- [ ] 标题字号层级：`68px` → `47px` → `27px` → `19px`
- [ ] Hero 标题使用 `JINGDONGLangZhengTi1-Bold`，`font-weight: 700`
- [ ] 卡片章节大标题 `h2` 使用 `MiSans-Bold`，`font-weight: 700`
- [ ] 其他字重使用 `900`（章节编号） / `600`（其他标题） / `800`（Hero 特殊） / `400`（正文）四档
- [ ] 品牌红 `#FF2B22` 用于所有强调元素
- [ ] 圆角逐级递减：`34` → `16` → `10` → `8` → `3`
- [ ] 灰色面板 `#F1F1F1`，圆角 `16px`
- [ ] 列表使用红色/灰色圆点，非默认 `list-style`
- [ ] 标题底部半透明品牌红高亮条（`linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%)`，替代旧的浅粉 #FFD4D4），四角圆角 `3px`（与列表红色方块圆角一致），水平内边距 `4px`
- [ ] 灰色面板内二级标题使用 `.label-line`：红色方块 + 荧光条
- [ ] `.label-line` 必须包含内层 `.label-text`，荧光条只放在 `.label-text` 上，不覆盖左侧红色方块
- [ ] `.label-line` 的下一级内容使用 `.source-list`：整体右缩进 `28px`，灰色方块；重点词只加粗、不加荧光条
- [ ] `.red-list` 的标题正文 `p` 也作为下一级处理：继承父级缩进后再右移 `8px`，灰色方块，和 `.source-list` 视觉层级一致
- [ ] 灰色面板的直系内容必须是白色模块或网格模块，不允许裸 `.label-line`、裸列表或裸图片直接挂在 `.gray-panel` 下
- [ ] 灰色面板内所有相邻直系模块必须有统一间距：`.gray-panel > * + * { margin-top: 18px; }`，不得出现白底模块互相贴边或重合
- [ ] 带标题的图片卡片：标题顶部对齐；图片保持原比例，在剩余区域上下左右居中
- [ ] 5 个及以上各屏/详情页示例必须使用 `.detail-screen-grid` 两列为主，必要时 `.span-full` 单列，不得使用四列 `.screens-grid` 挤压
- [ ] 所有卡片章节标题右侧英文标签统一为 `INTRODUCTION`，不得按内容改成其他英文
- [ ] 综述卡片标题也必须有右侧英文标签和箭头，不能省略 `.en-label`
- [ ] 英文标签右对齐，箭头 `margin-left: auto`；英文标签箭头（红色，钩朝左）与 Hero 分隔箭头（白色，钩朝右）使用同一套一体式 SVG 弯钩箭头，互为水平翻转，等比缩放不拉伸
- [ ] `.en-label` 高度 `47px` 与 h2 对齐（英文顶对齐、线底对齐）
- [ ] Hero 标记胶囊仅左右竖线（上下边框透明）
- [ ] 无 `box-shadow` 用于布局元素
- [ ] 无大面积内联样式覆盖；新增样式优先写入 CSS class
- [ ] Word 转 HTML 时，可见文字逐段校验通过，图片出现次数校验通过
- [ ] 有下级内容的“示例如下：”用 `.label-line`；孤立提示语用 `.example-line`，不得渲染为空卡片
- [ ] 所有间距为 4 的倍数或设计中明确的数值

---

## 十五、Word 转 HTML 生成执行协议

本节是给 AI 执行用的操作协议。生成其他文档时，必须先执行本协议，再应用上方视觉规范。

### 15.1 输入与输出

- 输入：一个 `.docx` 文档，以及用户指定的局部处理要求（如删除标题中的某个词）。
- 输出：一个可离线打开的单文件 HTML。
- 输出文件建议命名：`<source-name>-output.html`。
- 如生成过程需要中间资源目录，仅作为内部构建使用；最终交付优先为单文件 HTML。

### 15.2 执行顺序

1. **解析 Word**
   - 提取所有可见段落文本。
   - 提取段落顺序、列表层级、表格单元格顺序、图片引用顺序。
   - 提取所有图片并记录出现次数；同一图片重复出现时，HTML 中也重复出现。

2. **处理用户要求**
   - 只执行用户明确要求的删除、替换或调整。
   - 例如“把标题的官方建议去掉”，只删除标题中的“官方建议”，不得删除其他正文。
   - 不得把“已删除某某”“按原文整理”等生成说明写进 HTML 页面。

3. **建立页面骨架**
   - Word 第一行或主标题进入 `.hero h1`。
   - 若标题包含「商品信息运营规范」，必须在其前插入 `<br>`，使「商品信息运营规范」固定为第二行，前缀品类词在第一行。
   - 标题超过 2 行时，优先主动断行为 2 行；不要缩小 Hero 字号。
   - 后续一级章节按顺序生成 `.card.spec-card`。
   - 概述类内容生成 `.card.intro-card`。

4. **映射内容组件**
   - 按第十六章映射表将 Word 内容转成对应 HTML 组件。
   - 不确定层级时，优先保留原文顺序和完整文本，再选择最保守组件。
   - 内容过多时拆分卡片，不压缩字体和行高。

5. **封装资源**
   - CSS 写入 `<style>`。
   - `JINGDONGLangZhengTi1-Bold.ttf` 可用时用 `@font-face` 内嵌；不可用时保留字体栈。
   - 图片优先转为 `data:` URI；如文件过大可使用相对 `assets/` 路径，但必须保证同目录可打开。

6. **验收**
   - 按第十八章验收协议执行。
   - 验收失败时先修正 HTML，再交付。

### 15.3 禁止行为

- 不得总结、删减、改写原文来换取版面简洁。
- 不得把普通正文改成不存在于原文的新标题。
- 不得新增解释性文案填补空白。
- 不得使用默认 Word 蓝色标题样式、默认列表样式或表格默认边框。
- 不得把图片拉伸变形。
- 不得把没有下级内容的提示语做成空白卡片。

---

## 十六、Word 层级到 HTML 组件映射表

| Word 内容形态 | HTML 组件 | 生成规则 |
|---------------|-----------|----------|
| 文档主标题 | `.hero h1` | 删除用户指定删词后进入 Hero；最多 2 行；若含「商品信息运营规范」则在其前强制 `<br>`，使其固定为第二行 |
| 更新日期/版本 | `.updated` | 没有日期时可沿用当前月份或省略，但不要编造业务日期 |
| 概述段落 | `.lead` | 放在 `.gray-panel.spec-text` 顶部；只用于该卡片核心总结 |
| 概述中的并列规则 | `.red-list` | 一级条目用红方块；标题 `b` 带荧光条 |
| `.red-list` 标题下正文 | `.red-list li > p` | 作为下一级；继承父级缩进后再右移 `8px`；灰方块 |
| 一级章节 | `.card.spec-card` + `.section-head.spec-head` | 章节编号 `01/02/03` 顺序生成；右侧英文标签固定为 `INTRODUCTION` |
| 灰面板内小标题 | `.label-line > .label-text` | 红方块 + 文字荧光条；荧光条只覆盖文字，不覆盖红方块 |
| 小标题下条目 | `.source-list` | 整体右缩进 `28px`；灰方块；重点词只加粗 |
| 孤立提示语 | `.example-line` | 只有“示例如下：”“注意：”且没有下级内容时使用 |
| 有下级内容的“示例如下：” | `.label-line` | 下方接示例网格、图片卡片或 `.source-list` |
| 双列图片或整改前后 | `.title-image-grid` / `.before-after-grid` | 卡片标题顶部对齐；图片保持比例居中 |
| 图片后紧跟说明文字 | `.caption-image-card` | 图片在上，说明文字在下；说明属于图片的下一级，不得提升为 `.label-line` 或图片卡片顶部标题 |
| 左右对比说明 | `.vs-grid` / `.image-compare` | 适合高端 vs 中低端、方案 A vs B |
| 2 类优先级 | `.selling-grid` | 两列并排；每列内部使用 `.selling-list` |
| 3-4 个流程/屏次 | `.screens-grid` / `.detail-block-grid` | 每个节点用 `.screen-card`；只适合数量少、内容短的流程节点 |
| 5 个及以上各屏/详情页展示示例 | `.detail-screen-grid` + `.image-frame.detail-screen-frame` | 默认两列排布；横向长图或重点图可加 `.span-full` 单列全宽；每张图必须有标题、`.image-holder` 和等比居中图片 |
| 表格中图片 + 中间说明 + 图片 | `.image-compare` | 还原左右图片和中间说明，不强行改成普通表格 |
| Word 三列表格：分类/内容要求/示例 | `.word-table-spec` | 必须保留表头和逐行对应关系；每行按“分类名 / 内容要求文本 / 对应图片”三列排版，不得拆成双列卡片或打乱行列对应 |
| 多张连续小图 | `.image-strip` | 保持原图顺序；不裁切 |

### 16.1 标题层级规则

- 页面只有一个 `.hero h1`。
- 卡片章节标题只用 `.section-head h2`，固定 `MiSans-Bold`。
- 卡片章节标题文本必须写成 `{ 标题 }`，左大括号后保留一个空格，右大括号前保留一个空格；不要写成 `{标题}`、`{标题 }`、`{ 标题}`、`{{ 标题 }}` 或任何双大括号形式。
- 每个卡片标题栏都必须使用 `.section-head.spec-head`，并包含右侧 `.en-label`，文字统一固定为 `INTRODUCTION`，包括“整体规范综述”这类综述卡片。
- 灰面板内小标题只用 `.label-line`。
- `.label-line` 下方不能直接平铺正文，必须使用 `.source-list` 或示例组件。
- 灰色面板 `.gray-panel` 的直系内容必须是白色模块或网格模块，例如 `.text-block`、`.image-frame`、`.caption-image-card`、`.word-table-spec` 外壳或 `.detail-screen-grid`；不得把裸 `.label-line`、`.source-list`、`.example-line` 或 `<img>` 直接塞进灰面板。
- `.source-list b` 只加粗，不加荧光条。
- 如果 Word 中图片后紧跟说明文字，且该文字是图片说明或示例公式，说明文字必须放在对应图片下方，作为图片的下一级内容。
- 图片说明文字不得使用红方块；优先使用灰色方块或普通小字，避免被误判成上一级小标题。

#### 16.1.1 章节识别与标题规范化（生成器约定）

这些规则同时约束 `batch_generate.py` 的自动识别和模型重建：

- **主标题前缀剥离**：文档主标题若带 `主标题：`、`标题：`、`文档标题：` 等前缀词，进入 `.hero h1` 前必须剥掉前缀，只保留真正标题（如 `主标题：安全锤商品信息运营规范` → `安全锤` / `商品信息运营规范`）。
- **优先信任 Word 标题样式**：当文档用了真正的 `Heading` 样式标章节时，只用「标题样式 + 显式章节序号（`（一）`/`一、`/`第N章`/`【第N屏】`）」判定一级章节；其余段落一律按章节内内容处理。
- **冒号结尾段落不是章节**：任何以 `：`/`:` 结尾的段落都是 `.label-line` 小标题或 lead-in，绝不提升为章节卡（如 `主图首张：`、`主图优化核心总结：` 必须留在所属章节内）。
- **前缀/编号启发式仅在无标题样式时启用**：`主图`、`卖点` 等前缀词或 `1. 2.` 编号行只有在文档完全没有标题样式时才作为章节兜底；且前缀词需短（≤12 字）且不含冒号，避免把句子误判成章节。
- **章节标题清洗**：显示在 `{ 标题 }` 里的文本要去掉前导章节序号（`（一）`/`一、`/`第N章`）和结尾冒号；章节序号改由左侧 `01`–`05` 徽标承载。
- **去重复标题**：若某段落去掉冒号后与所在卡片标题完全相同（如综述卡下的 `整体规范综述：`），视为标题复述，丢弃不渲染。
- **校验放行规范化改写**：`validate_output.py` 对上述改写（前缀剥离、序号入徽标、`{ }` 括号、结尾冒号）做归一化匹配，不再误报「文本缺失」。

#### 16.1.2 综述卡与内容分组组件（生成器约定）

以下约定同时约束 `batch_generate.py` 自动输出和模型重建，对应 `.poster.auto-doc` 作用域的组件 CSS：

- **综述首段独立白底**：综述（intro）卡的第一段总述句单独用 `.lead`（白底圆角），与下方要点区分开。
- **`【…】` 要点转红色列表**：以 `【主图】`、`【标题】` 等方括号开头的并列要点统一进 `.red-list`，方括号词作为 `<b>`（红色方块 + 荧光条），其后描述紧随其后。
- **顶层「前缀：内容」转红色列表**：章节里**顶层**（不在某标签分组下）形如 `总结：…`、`字数范围：…`、`卖点建议顺序：…`、`视频基础要求：…` 的行（冒号在中间、冒号后有内容）须进 `.red-list`，冒号前缀作为 `<b>`（红色方块 + 荧光条）。不要把它们当普通段落 `.plain-block`。以冒号**结尾**的纯标签（如 `结构：`）仍走 `.label-line`。
- **标签 + 示例图合并同一白底，且示例图下级缩进**：当 `1.首张主图模块化布局图：`、`示例：`、`第N屏：…` 等标签后紧跟示例图时，标签（`.label-line` 红方块 + 荧光条）、灰色方块说明 `.caption-line`（文案固定「示例图：」，**只用灰方块、不加红方块和荧光条**）与图片须包在**同一个** `.text-block` 白底里；`.caption-line` 与图片作为标签的下一级须右缩进 `28px`（`.indent`，与 `.source-list` 对齐）。
- **优化前/优化后对比**：`2.优化前后图…` 这类标题与其下「优化前 / 优化后」内容须同处一个白底；用 `.ba-compare` 两列，`优化前`表头底色为灰 `.ba-before`，`优化后`为品牌红 `.ba-after`。
- **三列规范表**：`主图 / 内容要求 / 示例` 三列表用 `.spec-table`（`grid-template-columns: 1fr 2fr 3fr`）：`主图`列最窄、`内容要求`约其两倍、`示例`列约占一半宽；数据单元格内容垂直居中，`示例`列图片 `width:100%` 保持等宽，并与标题同处一个白底。
- **转化率绿色强调条**：`XX率+X%` 指标用 `.metric-emphasis`（**白底 + 绿色描边 2px `#47b250`、圆角 14px**；标签部分 `.metric-text` 用**黑色** `#111`；`+X%`（**含正号**，拆分点在正号前）用绿色 `#47b250` 的 `.metric-value`、放大约两倍，其后紧跟绿色**上扬箭头** `.metric-arrow`（内嵌 SVG，`fill:#47b250`，高约 34px）；内容 `justify-content:center` 水平居中、`align-items:center` 垂直居中）。两种情形：①**独立成行**（如 `搜索点击率+2.5%`）时作为灰面板直系模块、左右撑满灰面板；②**嵌在更长标签内**（如 `2.优化前后图 商详转化率+2%`）时，把指标从标签中**拆出来**，放进该标签所在的白底模块、置于 `.label-line` 标题**下方**，宽度撑满白底内容区（即与下方 `.ba-compare` 的「优化前+优化后」两列同宽）。以后所有这类「标签内嵌指标」一律按 ② 处理。
- **图文详情半幅图**：`图文详情` 卡内示例图用 `.half-image`（`max-width:600px`，约页面一半宽）并居中，避免长图占满整宽。
- **同一标签下多张并排示例图半宽等宽**：当一个标签（如 `短标题` 的「示例：」）下有 **≥2 张** 并排参考图时，holder 加 `.sample-image`（`doc-image width:50%`），每张图取容器一半宽且**等宽**（不论原图比例）；**单张**示例图保持原大图、不缩半。`图文详情` 卡的连续屏次图仍用 `.half-image`。
- **更新日期默认取源文件修改日期、精确到日**：`.updated` 默认用源 `.docx` 的最后修改日期（`st_mtime`），格式 `更新日期 YYYY年M月D日`（月/日不补零，如 `更新日期 2026年6月22日`）；显式传 `--updated` 支持 `YYYY.MM` 或 `YYYY.MM.DD`，给到日则输出到日。
- **卡片正文整体放大约 1.5×（生成器默认比例）**：为匹配目标视觉比例，卡片正文统一放大约 1.5 倍——`.lead`/`.red-list`/`.label-line`/`.source-list`/`.plain-block p`≈28px、`.example-line`≈27px、`.caption-line`≈26px、表格 `.doc-table`/`.spec-cell`≈24px、`.ba-head`≈27px/`.ba-text`≈23px；同时列表圆点≈16px、二级缩进≈38–42px、模块间距≈27px 一并等比放大。hero 区主标题 `h1`（68→102px）、右上 `OPERATION STANDARDS`（14→21px，连同其括号胶囊框 `.hero-mark` 155×62→232×93px、圆角 32→48px、边框 1→2px 一并放大，文字仍居中以保持与括号的间距）、右下更新日期（18→27px）同样 ×1.5，`.hero` 增高至 600px 以容纳放大后的标题、分隔线与日期。**唯卡片标题栏 `.section-head`（`h2` 47px / `.chapter` 17px / `.en-label`＝INTRODUCTION 15px）与转化率绿框 `.metric-emphasis`（保持原始 24/40px 文字＋34px 箭头）不参与放大。** 实现：在 `GENERIC_CSS` 末尾用更高特异性 `.poster.auto-doc .类名` 覆盖 §2.2／§17 基础字号。

### 16.2 图片处理规则

- 图片必须保留原比例。
- 图片放入卡片时，使用 `object-fit: contain`。
- 带标题图片卡片中，标题固定在顶部，图片在标题下方剩余空间居中。
- 所有内容图片必须放在 `.image-holder` 内，不得让 `<img>` 直接成为灰面板或白色模块的散落子节点。
- 灰色面板里的白底内容块、图片块、表格块、网格块都必须作为直系模块参与统一垂直间距；新增模块不得只靠局部相邻选择器补距，统一由 `.gray-panel > * + * { margin-top: 18px; }` 防止贴边或重合。
- 同一图片在 Word 中重复出现，HTML 中也重复出现，不合并。
- 如果图片本身是长图，不裁切；必要时限制 `max-height` 并让图片等比缩小。

### 16.3 各屏展示示例规则

- 当原文出现“第1屏、第2屏……”“各屏展示示例”“详情页屏次”等连续屏次内容，且数量达到 5 个及以上时，必须使用 `.detail-screen-grid`。
- `.detail-screen-grid` 默认两列：一行两个示例，列间距和行间距保持一致，不能使用四列 `.screens-grid` 挤压图片。
- 每个示例用 `.image-frame.detail-screen-frame`，标题用 `.label-line > .label-text` 放在图片上方，图片放在 `.image-holder` 中等比居中。
- 横向长图、信息密集图或需要单独强调的示例可加 `.span-full` 单列全宽；不要把所有示例无判断地竖向堆成一串。
- 如果某个屏次只有文字说明没有图片，放入同一网格中的白色说明卡，仍需保持和图片卡同级，不得混到上一级列表里。

### 16.4 Word 表格处理规则

- Word 表格如果承担“字段对应关系”，必须先保留表格语义，再做视觉美化。
- 判断标准：出现明确表头，如“主图 / 内容要求 / 示例”“模块 / 规范 / 示例图”“优化前 / 说明 / 优化后”时，按行列对应关系生成。
- 三列表格优先使用 `.word-table-spec`：第一列放分类名，第二列放内容要求，第三列放对应图片。
- 三列表格列宽规则：整张表必须由父级 `.word-table-spec` 统一控制三列轨道；第一列固定；第二列按全表“内容要求”中最长条目统一定宽；第三列使用第二列缩窄后释放出的剩余空间。
- 三列表格对齐规则：每个数据单元格里的内容都必须在单元格内上下居中；第二列的多条列表整体作为一组纵向居中，不要贴顶部。
- 不得把这类表格自动拆成普通图片卡片、双列网格或自由卡片，否则会破坏原文层级。
- 只有当 Word 表格本身只是图片排版容器、没有字段对应关系时，才可转为图片网格组件。

---

## 十七、完整 CSS 模板

生成单文件 HTML 时，将以下 CSS 放入 `<style>`。其中 `@font-face` 的 `src` 可替换为字体文件 data URI；若无法内嵌字体，则保留字体栈。

```css
@font-face {
  font-family: "JINGDONGLangZhengTi1-Bold";
  src: url("JINGDONGLangZhengTi1-Bold.ttf") format("truetype");
  font-weight: 700;
  font-style: normal;
  font-display: block;
}

* { box-sizing: border-box; }

html, body {
  margin: 0;
  min-height: 100%;
  background: #737373;
  color: #333;
  font-family: "Arial", "Microsoft YaHei", "PingFang SC", sans-serif;
}

.poster {
  width: 1280px;
  margin: 0 auto;
  background: #dcedff;
  overflow: hidden;
}

.hero {
  position: relative;
  height: 520px;
  padding: 72px 44px 0;
  background:
    radial-gradient(circle at 78% 22%, rgba(255,255,255,0.28) 0 2px, transparent 3px 100%),
    linear-gradient(102deg, rgba(255,48,38,0.98), rgba(255,30,22,0.96));
  color: #fff;
  overflow: hidden;
}

.hero::before {
  content: "";
  position: absolute;
  inset: 0;
  opacity: 0.18;
  background-image:
    linear-gradient(85deg, transparent 0 12%, rgba(255,255,255,0.9) 12.2%, transparent 12.6% 100%),
    repeating-linear-gradient(0deg, rgba(255,255,255,0.45) 0 2px, transparent 2px 22px),
    repeating-linear-gradient(90deg, rgba(255,255,255,0.35) 0 2px, transparent 2px 28px);
  transform: skewY(-4deg) scale(1.08);
}

.robot-deco {
  position: absolute;
  inset: 0;
  opacity: 0.52;
  pointer-events: none;
}

.ring {
  position: absolute;
  border: 3px solid rgba(255,255,255,0.44);
  border-radius: 50%;
}
.ring::before {
  content: "";
  position: absolute;
  inset: 26px;
  border: 2px solid rgba(255,255,255,0.3);
  border-radius: 50%;
}
.ring::after {
  content: "";
  position: absolute;
  inset: 50% auto auto 50%;
  width: 60px;
  height: 60px;
  margin: -30px 0 0 -30px;
  background:
    linear-gradient(rgba(255,255,255,0.45) 0 2px, transparent 2px 100%) 50% 50% / 60px 60px no-repeat,
    linear-gradient(90deg, rgba(255,255,255,0.45) 0 2px, transparent 2px 100%) 50% 50% / 60px 60px no-repeat;
}
.ring-one { right: -60px; top: -80px; width: 360px; height: 360px; }
.ring-two { right: 280px; top: 60px; width: 220px; height: 220px; }
.ring-two::before { inset: 18px; }
.ring-three { right: 100px; top: 280px; width: 160px; height: 160px; }
.ring-three::before { inset: 14px; }

.path-line {
  position: absolute;
  height: 2px;
  background: rgba(255,255,255,0.42);
}
.path-a { left: 30px; right: 600px; top: 220px; transform: rotate(-3deg); }
.path-b { left: 70px; right: 700px; top: 280px; transform: rotate(2deg); }
.path-c { left: 120px; right: 350px; top: 360px; transform: rotate(-2deg); }
.path-d { left: 40px; right: 720px; top: 420px; transform: rotate(1deg); }

.hero h1 {
  position: relative;
  z-index: 1;
  margin: 0;
  font-family: "JINGDONGLangZhengTi1-Bold", "jingdonglangzhengti1", "Microsoft YaHei", "PingFang SC", sans-serif;
  font-size: 68px;
  line-height: 1.12;
  letter-spacing: 0;
  font-weight: 700;
}

.hero-rule {
  position: relative;
  z-index: 1;
  width: 118px;
  height: 11px;
  margin-top: 124px;
  background: url("data:image/svg+xml,%3Csvg%20xmlns=%27http://www.w3.org/2000/svg%27%20viewBox=%270%200%20138.209%2012.5%27%20preserveAspectRatio=%27none%27%3E%3Cpath%20d=%27M130.44582%209.5L0%209.5L0%2012.5L138%2012.5L138.20905%209.5146379Q137.62546%209.4325085%20136.69609%209.1682091Q134.81464%208.6331568%20133.25896%207.7437816Q128.52896%205.0396996%20128.52898%20-2.420493e-06L125.52898%202.420493e-06Q125.52896%203.7219667%20127.51295%206.5661678Q128.67047%208.2256107%20130.44582%209.5Z%27%20fill=%27%23ffffff%27/%3E%3C/svg%3E") no-repeat center / contain;
}
.updated {
  position: absolute;
  right: 50px;
  bottom: 116px;
  z-index: 1;
  margin: 0;
  font-size: 18px;
  font-weight: 800;
}
.hero-mark {
  position: absolute;
  top: 70px;
  right: 48px;
  z-index: 1;
  width: 155px;
  height: 62px;
  border: 1px solid rgba(255,255,255,0.95);
  border-top-color: transparent;
  border-bottom-color: transparent;
  border-radius: 32px;
  display: grid;
  place-items: center;
  font-size: 14px;
  line-height: 1.25;
  font-weight: 800;
}

.card {
  width: 1188px;
  margin: -75px auto 0;
  padding: 68px 42px 58px;
  border-radius: 34px;
  background: #fff;
  position: relative;
  z-index: 2;
  font-family: "MiSans-Normal", "MiSans", "Microsoft YaHei", "PingFang SC", sans-serif;
}
.card:last-child { margin-bottom: 46px; }
.intro-card { min-height: 0; }
.spec-card {
  margin-top: 38px;
  padding-top: 62px;
  padding-bottom: 48px;
}

.section-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  min-height: 86px;
}
.section-head h2 {
  margin: 0;
  color: #ff2b22;
  font-family: "MiSans-Bold", "MiSans", "Microsoft YaHei", "PingFang SC", sans-serif;
  font-size: 47px;
  line-height: 1;
  font-weight: 700;
  letter-spacing: 0;
}
.spec-head {
  position: relative;
  justify-content: flex-start;
  gap: 16px;
}
.spec-head .en-label { margin-left: auto; }
.chapter {
  color: #ff2b22;
  font-family: "MiSans-Heavy", "MiSans", sans-serif;
  font-size: 17px;
  font-weight: 900;
  line-height: 1;
}
.en-label {
  width: 192px;
  height: 47px;
  color: #ff2b22;
  text-align: right;
  font-size: 15px;
  line-height: 1;
  font-weight: 600;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}
.en-label span {
  display: block;
  width: 67px;
  height: 6px;
  margin-left: auto;
  background: url("data:image/svg+xml,%3Csvg%20xmlns=%27http://www.w3.org/2000/svg%27%20viewBox=%270%200%20138.209%2012.5%27%20preserveAspectRatio=%27none%27%3E%3Cg%20transform=%27translate(138.209,0)%20scale(-1,1)%27%3E%3Cpath%20d=%27M130.44582%209.5L0%209.5L0%2012.5L138%2012.5L138.20905%209.5146379Q137.62546%209.4325085%20136.69609%209.1682091Q134.81464%208.6331568%20133.25896%207.7437816Q128.52896%205.0396996%20128.52898%20-2.420493e-06L125.52898%202.420493e-06Q125.52896%203.7219667%20127.51295%206.5661678Q128.67047%208.2256107%20130.44582%209.5Z%27%20fill=%27%23ff2b22%27/%3E%3C/g%3E%3C/svg%3E") no-repeat center / contain;
}

.gray-panel {
  background: #f1f1f1;
  border-radius: 16px;
}
.spec-text { margin-top: 30px; padding: 30px 28px 24px; }
.example-block { margin-top: 50px; padding: 24px 28px 28px; }
.gray-panel > * + * { margin-top: 18px; }

.lead {
  margin: 0 0 22px;
  padding: 24px 34px;
  border-radius: 10px;
  background: #fff;
  font-size: 19px;
  line-height: 1.55;
  font-weight: 600;
}

.red-list,
.source-list,
.selling-list {
  list-style: none;
  margin: 0;
  padding: 0;
}
.red-list li {
  position: relative;
  padding-left: 25px;
  margin: 0 0 16px;
}
.spec-text .red-list li { margin-bottom: 18px; }
.red-list li::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #ff2b22;
}
.red-list b {
  display: inline;
  font-size: 19px;
  line-height: 1.4;
  font-weight: 600;
  background: linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.red-list p {
  margin: 3px 0 0;
  color: #333;
  font-size: 19px;
  line-height: 1.38;
}
.red-list li > p:not(.sub-dot) {
  position: relative;
  margin: 8px 0 0 8px;
  padding-left: 25px;
  line-height: 1.45;
}
.red-list li > p:not(.sub-dot)::before,
.source-list li::before,
.sub-dot::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #d8d8d8;
}
.sub-dot {
  position: relative;
  padding-left: 28px;
}
.sub-dot::before { top: 10px; }

.text-block {
  background: #fff;
  border-radius: 10px;
  padding: 22px 26px;
  margin-top: 18px;
}
.text-block:first-child { margin-top: 0; }
.label-line {
  display: inline-flex;
  align-items: flex-start;
  position: relative;
  margin: 0 0 10px;
  padding-left: 25px;
  color: #111;
  font-size: 19px;
  line-height: 1.4;
  font-weight: 600;
}
.label-line::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #ff2b22;
}
.label-line .label-text {
  display: inline;
  background: linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.source-list {
  margin: 0 0 0 28px;
}
.source-list li {
  position: relative;
  padding-left: 25px;
  margin: 12px 0 0;
  color: #333;
  font-size: 19px;
  line-height: 1.45;
}
.source-list li:first-child { margin-top: 0; }
.source-list b {
  color: #111;
  font-weight: 600;
  margin-right: 3px;
}

.example-block h3 {
  display: inline-block;
  margin: 0 0 22px;
  color: #333;
  font-size: 27px;
  line-height: 1;
  font-weight: 600;
  background: linear-gradient(transparent 62%, rgba(255,43,34,0.2) 62%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}

.doc-image {
  display: block;
  max-width: 100%;
  height: auto;
  border-radius: 8px;
  background: #f1f1f1;
}
.image-frame {
  background: #fff;
  border-radius: 10px;
  padding: 18px;
}
.image-frame .source-list { margin-bottom: 12px; }
.image-frame .doc-image {
  margin: 0 auto;
  max-height: 420px;
  object-fit: contain;
}

.title-image-grid,
.before-after-grid,
.selling-grid,
.bind-compare {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px;
  margin-top: 18px;
  align-items: stretch;
}
.title-image-grid .image-frame,
.before-after-grid .image-frame {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  min-height: 380px;
}
.title-image-grid .image-frame .doc-image,
.before-after-grid .image-frame .doc-image {
  margin: auto;
  width: auto;
  max-width: 100%;
  max-height: 320px;
  object-fit: contain;
}
.caption-image-card {
  display: flex;
  flex-direction: column;
  min-height: 320px;
  background: #fff;
  border-radius: 10px;
  padding: 18px;
}
.caption-image-card .image-holder {
  flex: 1;
  min-height: 220px;
  display: grid;
  place-items: center;
}
.caption-image-card .doc-image {
  width: auto;
  max-width: 100%;
  max-height: 300px;
  object-fit: contain;
}
.image-caption-line {
  position: relative;
  margin-top: 14px;
  padding-left: 25px;
  color: #111;
  font-size: 18px;
  line-height: 1.38;
  font-weight: 600;
}
.image-caption-line::before {
  content: "";
  position: absolute;
  left: 0;
  top: 9px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #d8d8d8;
}

.compare-head,
.image-compare {
  display: grid;
  grid-template-columns: 1fr 280px 1fr;
  gap: 14px;
  align-items: stretch;
  margin-top: 18px;
}
.compare-head div {
  display: grid;
  place-items: center;
  min-height: 46px;
  border-radius: 8px;
  color: #fff;
  font-size: 20px;
  font-weight: 600;
}
.compare-head .left { background: #111; }
.compare-head .mid,
.compare-head .right { background: #ff2b22; }
.compare-text {
  background: #fff;
  border-radius: 10px;
  padding: 18px 18px 16px;
  display: flex;
  flex-direction: column;
  justify-content: center;
}
.compare-text h4 {
  margin: 0 0 10px;
  color: #111;
  font-size: 21px;
  line-height: 1.25;
  font-weight: 600;
}
.compare-text p {
  margin: 7px 0 0;
  color: #333;
  font-size: 16px;
  line-height: 1.45;
}
.image-strip {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 8px;
  align-items: center;
}
.image-strip img {
  width: 100%;
  border-radius: 6px;
  background: #f1f1f1;
}

.selling-col,
.bind-col,
.screen-card {
  background: #fff;
  border-radius: 10px;
}
.selling-col { padding: 24px 26px 26px; }
.selling-head {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 18px;
}
.selling-num {
  width: 44px;
  height: 44px;
  display: grid;
  place-items: center;
  background: #ff2b22;
  color: #fff;
  border-radius: 8px;
  font-size: 22px;
  font-weight: 600;
}
.selling-name {
  font-size: 24px;
  font-weight: 600;
  color: #111;
}
.selling-list li {
  position: relative;
  padding: 10px 0 10px 22px;
  border-bottom: 1px solid #f1f1f1;
  font-size: 18px;
  color: #333;
  line-height: 1.4;
}
.selling-list li:last-child { border-bottom: none; }
.selling-list li::before {
  content: "";
  position: absolute;
  left: 0;
  top: 18px;
  width: 9px;
  height: 9px;
  background: #ff2b22;
  border-radius: 3px;
}
.selling-list b {
  font-weight: 600;
  color: #111;
}

.detail-block-grid,
.screens-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 14px;
  margin-top: 18px;
}
.screens-grid { grid-template-columns: repeat(4, 1fr); }
.detail-screen-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 18px;
  margin-top: 18px;
  align-items: stretch;
}
.detail-screen-grid .span-full,
.detail-screen-frame.span-full {
  grid-column: 1 / -1;
}
.detail-screen-frame {
  display: flex;
  flex-direction: column;
  min-height: 520px;
}
.detail-screen-frame .image-holder {
  flex: 1;
  min-height: 360px;
  display: grid;
  place-items: center;
  overflow: hidden;
}
.detail-screen-frame .doc-image {
  width: auto;
  max-width: 100%;
  max-height: 640px;
  object-fit: contain;
}
.screen-card {
  padding: 16px 16px 18px;
  display: flex;
  flex-direction: column;
}
.screen-no {
  font-size: 14px;
  font-weight: 600;
  color: #ff2b22;
  letter-spacing: 0.06em;
  margin-bottom: 6px;
}
.screen-title {
  font-size: 19px;
  font-weight: 600;
  color: #111;
  line-height: 1.25;
  margin-bottom: 10px;
}
.screen-desc {
  flex: 1;
  font-size: 14px;
  color: #555;
  line-height: 1.45;
  margin-bottom: 12px;
}
.screen-mock {
  height: 200px;
  border-radius: 6px;
  background: linear-gradient(180deg, #f7f7f7, #e4e4e4);
  border: 1px dashed #cfcfcf;
  display: grid;
  place-items: center;
  color: #8a8a8a;
  font-size: 13px;
}

.example-line {
  margin: 18px 0 0;
  color: #ff2b22;
  font-size: 18px;
  line-height: 1.4;
  font-weight: 600;
}

@media (max-width: 1280px) {
  .poster { transform-origin: top left; }
}
```

### 17.1 补充组件 CSS

以下组件不是每篇文档都必须使用，但属于参考 HTML/CSS 中已经验证过的通用组件。生成新 HTML 时，只要文档内容出现对应结构，就把这段接在第十七章主 CSS 后面使用。

```css
/* ===== VS 对比展示 ===== */
.vs-grid {
  display: grid;
  grid-template-columns: 1fr 80px 1fr;
  gap: 16px;
  align-items: stretch;
  margin-top: 16px;
}
.vs-col {
  background: #fff;
  border-radius: 10px;
  padding: 22px 22px 26px;
  min-height: 280px;
  display: flex;
  flex-direction: column;
}
.vs-tag {
  align-self: flex-start;
  padding: 6px 14px;
  border-radius: 6px;
  font-size: 16px;
  line-height: 1;
  font-weight: 600;
  color: #fff;
  margin-bottom: 16px;
}
.vs-tag.high { background: #111; }
.vs-tag.mid { background: #ff2b22; }
.vs-col h4 {
  margin: 0 0 10px;
  font-size: 22px;
  line-height: 1.2;
  font-weight: 600;
  color: #111;
}
.vs-col p {
  margin: 0 0 12px;
  font-size: 17px;
  line-height: 1.45;
  color: #333;
}
.vs-image {
  margin-top: auto;
  height: 140px;
  border-radius: 8px;
  display: grid;
  place-items: center;
  color: #fff;
  font-size: 18px;
  font-weight: 600;
}
.vs-image.scene {
  background: linear-gradient(120deg, #b9c8d3, #e0e8ee 55%, #aab8c4);
  color: #333;
}
.vs-image.param {
  background: linear-gradient(135deg, #f4f4f4, #d0d0d0);
  color: #555;
}
.vs-divider {
  display: grid;
  place-items: center;
  color: #ff2b22;
  font-size: 36px;
  line-height: 1;
  font-weight: 600;
  letter-spacing: 0;
}

/* ===== 5 帧展示 ===== */
.frames-grid {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 12px;
  margin-top: 18px;
}
.frame {
  background: #fff;
  border-radius: 10px;
  padding: 16px 14px 18px;
  display: flex;
  flex-direction: column;
}
.frame-no {
  font-size: 14px;
  line-height: 1.2;
  font-weight: 600;
  color: #ff2b22;
  margin-bottom: 6px;
  letter-spacing: 0.08em;
}
.frame-title {
  font-size: 19px;
  line-height: 1.2;
  font-weight: 600;
  color: #111;
  margin-bottom: 10px;
}
.frame-desc {
  font-size: 14px;
  line-height: 1.4;
  color: #555;
  flex: 1;
}
.frame-img {
  margin-top: 12px;
  height: 110px;
  border-radius: 6px;
  background: #e8e8e8;
  display: grid;
  place-items: center;
  color: #8a8a8a;
  font-size: 14px;
}

/* ===== 标题示例展示 ===== */
.title-demo {
  display: grid;
  grid-template-columns: 200px 1fr;
  gap: 18px;
  align-items: stretch;
  margin-top: 18px;
}
.title-demo + .title-demo { margin-top: 18px; }
.title-label {
  background: #ff2b22;
  border-radius: 10px;
  display: grid;
  place-items: center;
  color: #fff;
  font-size: 22px;
  line-height: 1.2;
  font-weight: 600;
  text-align: center;
  padding: 18px 8px;
}
.title-label small {
  display: block;
  margin-top: 6px;
  font-size: 14px;
  font-weight: 400;
  letter-spacing: 0.06em;
  opacity: 0.92;
}
.title-content {
  background: #fff;
  border-radius: 10px;
  padding: 20px 24px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 8px;
}
.title-formula {
  font-size: 17px;
  line-height: 1.4;
  font-weight: 600;
  color: #555;
}
.title-formula em {
  font-style: normal;
  color: #ff2b22;
}
.title-example {
  display: inline;
  align-self: flex-start;
  font-size: 19px;
  line-height: 1.45;
  font-weight: 600;
  color: #111;
  background: linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.title-meta {
  font-size: 15px;
  line-height: 1.4;
  color: #8a8a8a;
}

/* ===== 系列品工具：整改前后 ===== */
.bind-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-bottom: 12px;
}
.bind-cell {
  background: #f1f1f1;
  border-radius: 6px;
  padding: 14px 8px;
  text-align: center;
  font-size: 14px;
  line-height: 1.25;
  font-weight: 600;
  color: #333;
}
.bind-cell .price {
  display: block;
  margin-top: 6px;
  color: #ff2b22;
  font-size: 16px;
}
.bind-label {
  height: 44px;
  border-radius: 8px;
  display: grid;
  place-items: center;
  font-size: 22px;
  line-height: 1;
  font-weight: 600;
  margin-bottom: 16px;
}
.bind-label.before { background: #f1f1f1; color: #555; }
.bind-label.after { background: #ff2b22; color: #fff; }
.bind-note {
  font-size: 14px;
  line-height: 1.4;
  color: #8a8a8a;
  margin-top: 8px;
  text-align: center;
}

/* ===== Word 三列表格：分类 / 内容要求 / 示例 ===== */
.word-table-spec {
  margin-top: 18px;
  display: grid;
  grid-template-columns: 126px fit-content(410px) minmax(0, 1fr);
  gap: 12px;
  align-items: stretch;
}
.word-table-head,
.word-table-row {
  display: contents;
}
.word-table-head span {
  display: grid;
  place-items: center;
  min-height: 38px;
  border-radius: 8px;
  background: #ff2b22;
  color: #fff;
  font-size: 17px;
  line-height: 1;
  font-weight: 600;
}
.word-table-cell {
  background: #f7f7f7;
  border-radius: 10px;
  padding: 16px;
  min-height: 260px;
}
.word-table-main {
  display: grid;
  place-items: center;
}
.word-table-main span {
  display: inline-block;
  color: #111;
  font-size: 21px;
  line-height: 1.1;
  font-weight: 600;
  background: linear-gradient(transparent 58%, rgba(255,43,34,0.2) 58%);
  border-radius: 3px;
  padding: 0 4px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.word-table-req .source-list {
  margin-left: 0;
}
.word-table-req {
  display: grid;
  align-items: center;
}
.word-table-req .source-list li {
  font-size: 17px;
  line-height: 1.38;
  white-space: nowrap;
}
.word-table-example {
  display: grid;
  place-items: center;
}
.word-table-example .image-holder {
  min-height: 230px;
  width: 100%;
  display: grid;
  place-items: center;
}
.word-table-example .doc-image {
  width: auto;
  max-width: 100%;
  max-height: 420px;
  object-fit: contain;
}

/* ===== 商详结构：屏次展示变体 ===== */
.screen-mock.scene {
  background: linear-gradient(180deg, #e4dccc, #f5edd9);
  border-color: #c9b890;
  color: #6b5d3a;
}
.screen-mock.compare {
  background:
    repeating-linear-gradient(0deg, transparent 0 20px, #e8e8e8 20px 21px),
    #fff;
  border-color: #cfcfcf;
}
```

---

## 十八、验收与交付协议

生成 HTML 后，必须完成以下验收。

### 18.1 内容保真验收

- 提取 Word 可见段落文本。
- 删除用户明确要求删除的词后，逐段检查是否出现在 HTML 可见文本中。
- 统计 Word 图片引用次数，检查 HTML `<img>` 数量是否一致。
- 检查 HTML 中是否残留用户要求删除的词。
- 检查 HTML 中是否出现生成说明、解释说明或非原文补充文案。

### 18.2 视觉验收

- 用浏览器打开 HTML。
- 截图检查至少包含：
  - Hero 首屏。
  - 概述卡片。
  - 至少一个标准章节卡片。
  - 至少一个图片示例区。
- 检查标题字体：
  - Hero `h1` 为 `JINGDONGLangZhengTi1-Bold`。
  - 卡片 `h2` 为 `MiSans-Bold`。
  - 卡片 `h2` 文本格式为 `{ 标题 }`，大括号内左右各一个空格。
- 检查层级：
  - `.red-list b` 是红方块 + 荧光条。
  - `.red-list p` 是灰方块 + 下一级缩进。
  - `.label-line` 是红方块 + `.label-text` 荧光条。
  - `.source-list` 是灰方块 + 右缩进。
- 检查图片：
  - 不拉伸、不裁切。
  - 带标题图片卡片标题顶部对齐。
  - 图片在剩余空间居中。
  - 图片后紧跟说明文字时，说明必须在对应图片下方，不得放到图片上方或提升为同级标题。
- 检查 Word 表格：
  - 有明确表头的表格必须保留表头。
  - 每行分类、内容要求、示例图片必须横向对应。
  - 内容要求列宽应按全表最长内容统一定宽；所有行宽度一致，示例列吃剩余宽度。
  - 每个数据单元格内容上下居中，内容要求列表整体也必须上下居中。
  - 不得把三列表格拆成双列卡片后破坏对应关系。
- 检查没有空卡片、重叠文字、明显过深缩进、图片溢出。

### 18.3 交付要求

- 最终只交付 HTML 文件，除非用户要求保留构建脚本或资源目录。
- 如果使用相对图片路径，必须同时交付资源目录。
- 如果使用 data URI，HTML 文件应可单独打开。
- 最终回复中说明：
  - 输出文件路径。
  - 是否通过文字保真校验。
  - 是否通过图片数量校验。
  - 是否内嵌标题字体。

### 18.4 失败处理

- 文字缺失：先补文字，不调整视觉。
- 图片缺失：先补图片，不调整视觉。
- 层级错乱：按第十六章映射表重新归类。
- 缩进不一致：优先检查 `.red-list p`、`.label-line`、`.source-list` 三者关系。
- 图片变形：优先检查 `.doc-image` 是否为 `object-fit: contain` 且未设置固定宽高。
