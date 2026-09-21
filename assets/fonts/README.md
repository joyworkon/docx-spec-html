# 字体资源说明

## JINGDONGLangZhengTi1-Bold.woff2

Hero 大标题专用字体（京东朗正体），已由生成器内嵌为 data-URI，随页面分发。

## misans-src/（MiSans Normal / Bold / Heavy）

规范页正文、卡片标题使用的 MiSans 字体源文件（TTF，取自小米官方发布的 MiSans
字体包）。`scripts/batch_generate.py` 在生成页面时会用 fontTools 按页面实际用字
对这三个字重做子集化（woff2），再以 data-URI `@font-face` 嵌入最终 HTML——每个
页面只需嵌入约几十 KB，不含整字库，因此接收方无需安装 MiSans 即可得到与设计一致
的显示效果。

要求：

- 需要 `pip install fonttools brotli`；缺 fontTools 或缺少任一源 TTF 时，生成器
  静默回退为按字体名引用（本机已装 MiSans 的用户不受影响）。
- MiSans 字体的著作权与再分发条款以小米官方的 MiSans 许可声明为准；如需更新
  字体版本，替换本目录下对应 TTF 即可，文件名保持 `MiSans-{Normal,Bold,Heavy}.ttf`。
