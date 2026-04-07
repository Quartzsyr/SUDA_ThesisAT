# SUDA_ThesisAT

一个面向本科论文场景的原生桌面工具，用于导入 Word 论文、补全封面字段、实时预览 PDF，并导出标准化的 PDF / DOCX。
**[Latex模板见这里](https://github.com/Quartzsyr/Soochow-University-Thesis-Overleaf-LaTeX-Template)**
@Tianhaoo
## 项目简介

`SUDA_ThesisAT` 主要解决论文整理过程中的几个高频问题：

- 从 `.docx` 导入论文正文
- 快速补全封面字段
- 实时查看 PDF 预览效果
- 在本地离线环境下一键导出 PDF 和 DOCX

项目当前采用 PyQt6 原生桌面界面，强调稳定性、预览优先和高效表单录入，不依赖 Web 前端。

## 功能特性

- 导入 `.docx` 论文源文件
- 自动提取正文并转换为可编辑内容
- 支持多学校论文模板切换
- 封面字段快速录入与补全
- 实时 PDF 预览
- 支持缩放、适配宽度、页码反馈
- 支持导出 PDF 和 DOCX
- 导出前基础字段检查
- 支持打开导出目录
- 纯本地运行，无需联网

## 技术栈

- PyQt6
- PyMuPDF
- python-docx
- PyInstaller

## 项目结构

```text
app.py                    # 主界面入口
paperwrite/docx_loader.py # Word 导入与正文提取
paperwrite/formatting.py  # 论文格式化与导出逻辑
paperwrite/preview.py     # PDF 预览组件
assets/                   # 模板、图标、界面样式资源
ThesisFlow.spec           # PyInstaller 打包配置
installer.iss             # Inno Setup 安装包脚本
```

## 使用方式

### 直接运行

```bash
python app.py
```

### 基本流程

1. 选择论文模板
2. 导入 `.docx` 源文件
3. 补全封面字段
4. 观察右侧 PDF 实时预览
5. 选择输出位置
6. 导出 PDF / DOCX

## 打包

使用 PyInstaller：

```bash
pyinstaller ThesisFlow.spec
```

生成安装包时可配合 Inno Setup：

```bash
installer.iss
```

## 当前设计方向

项目目前重点优化以下方向：

- 学术工作台风格的原生桌面体验
- 预览优先的大工作区布局
- 更简洁的顶部导航和状态反馈
- 更高效的封面字段录入流程
- 更稳定的本地导出链路

## 适用场景

- 本科毕业论文封面整理
- 学校论文模板适配
- 导出前的快速检查与预览
- 需要离线处理 Word / PDF 的桌面工作流

## 后续规划

- 自动保存草稿
- 最近项目列表
- 更多学校模板支持
- 更完整的导出前检查
- 模板预设管理
- 更细致的预览交互

## License

MIT
