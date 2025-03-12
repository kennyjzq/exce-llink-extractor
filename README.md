# Excel文件链接导出工具

一个简单易用的网页应用，用于从Excel文件中提取所有超链接，并支持导出为TXT或Excel格式。

## 功能特点

- 支持拖放或选择上传Excel文件（.xlsx和.xls格式）
- 全面提取Excel文件中的所有类型超链接，包括：
  - 单元格直接超链接
  - 单元格链接属性（cell.l）
  - HYPERLINK函数创建的超链接
  - 单元格HTML内容中的超链接
  - 单元格中的URL文本
- 按照链接在Excel中的顺序提取，保持原始顺序
- 在网页上以表格形式显示提取的链接，包含链接文本、URL、所在工作表和单元格位置
- 支持导出链接为TXT文本文件
- 支持导出链接为Excel文件（带有可点击的超链接）
- 苹果风格的现代化UI设计
- 完全在浏览器中运行，无需服务器支持
- 响应式设计，适配各种设备屏幕

## 技术栈

- 纯前端实现：HTML5 + CSS3 + JavaScript
- 使用SheetJS库解析Excel文件
- 无需后端服务器，可直接在浏览器中运行

## 使用方法

1. 直接在浏览器中打开`index.html`文件
2. 拖放Excel文件到指定区域或点击选择文件
3. 等待文件处理完成，查看提取的链接
4. 可选择导出为TXT或Excel格式

## 部署说明

本应用是纯前端实现，可以通过以下方式部署：

1. **本地使用**：直接在浏览器中打开index.html文件
2. **虚拟主机部署**：将所有文件上传到虚拟主机的根目录或子目录
3. **静态网站托管**：可部署到GitHub Pages、Netlify等静态网站托管服务

## 文件结构

```
excel-link-extractor/
├── index.html          # 主HTML文件
├── css/
│   └── styles.css      # 样式文件
├── js/
│   └── main.js         # JavaScript主文件
└── img/
    └── upload-icon.svg # 上传图标
```

## 超链接提取方法

本工具使用多种方法提取Excel文件中的超链接：

1. **直接超链接**：提取Excel文件中直接设置的超链接（worksheet['!hyperlinks']）
2. **单元格链接属性**：检查单元格是否有链接属性（cell.l.Target）
3. **HYPERLINK函数**：解析单元格中使用HYPERLINK函数创建的超链接
4. **HTML内容**：分析单元格中可能包含的HTML内容中的超链接
5. **URL文本**：识别单元格中的URL格式文本（以http://、https://或www.开头）

所有链接按照在Excel中的原始顺序提取和显示。

## 浏览器兼容性

- Chrome 60+
- Firefox 60+
- Safari 11+
- Edge 16+

## 注意事项

- 所有处理都在浏览器中完成，文件不会上传到任何服务器
- 大文件处理可能需要较长时间，请耐心等待
- 如果Excel文件中没有超链接，将显示相应提示
- 控制台中会输出详细的链接提取过程，便于调试

## 许可证

MIT 