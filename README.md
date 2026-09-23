# MarkPivot · 文档规范化

> 本地运行的文档规范化工具：把已经写好的 Word，按目标模板或规则识别、统一格式并批量输出可交付文件。

产品主叙事是 **Document → Normalize → Deliver**。Word / Excel / Markdown 转换能力仍保留为内部引擎；Markdown 不再作为面向用户的主概念。

## 当前状态

2026-09-23 完成 **M0–M11**：入口统一为 `MarkPivot.html`，主流程为独立任务「上传文件 → 选择规范 → 自动识别/微调 → 批量导出」。规范库可新建、编辑、复制、导入导出 JSON、从 Word 创建；任务内微调不会自动覆盖库里的规范。工作台中栏可切换原文/当前效果；完成调整后进入导出页。解压包需保持 `MarkPivot.html`、`vendor/`、`templates/` 同层；带着本地运行库即可离线启动。V4.1 转换引擎仍复用，不上传文档，不需要账号或后端。

## 立即使用

1. 保持 `MarkPivot.html`、`vendor/`、`templates/` 在同一目录。直接双击可以使用基础能力；推荐通过本地 HTTP 服务打开，以获得完整的文件与资源行为。
2. 在项目根目录运行：

   ```powershell
   python -m http.server 8765 --bind 127.0.0.1
   ```

   如果 Windows 中 `python` 不可用，可改用 `py -m http.server 8765 --bind 127.0.0.1`。
3. 访问 <http://127.0.0.1:8765/MarkPivot.html>。

带着 `vendor/` 首次打开不必访问 CDN。顶部状态显示「本地运行库可用（可离线）」时，9 个转换库都来自本机。只有缺少 vendor 且本机没有缓存时，才会尝试 jsDelivr / unpkg。`templates/reference-template.docx` 可作为参考规范样例上传。

## 核心工作方式

```text
Word / Excel / CSV / Markdown
              ↓ 导入
       Markdown 语义中枢
              ↓ 预览、编辑、套用规范
       Word / Excel / Markdown 导出
```

所有转换都经过 Markdown。平台不提供 Excel 直接转 Word 或 Word 直接转 Excel 的旁路入口；转换无法完全保留的内容进入“转换报告”，不会静默丢失。

## 支持范围

| 类型 | 支持能力 |
| --- | --- |
| Markdown | 导入 `.md/.txt`，源码/预览双态编辑，导出 Markdown 与 MD 包 |
| Word | 导入 `.docx`，Word 纸面预览可编辑回写 Markdown，导出 `.docx` |
| Excel | 导入 `.xlsx/.xls/.csv`，多 Sheet 网格编辑，导出 `.xlsx` 或 `.csv` |
| 批量 | 多选 Markdown、Word、Excel/CSV，按目标格式旁路转换并打包 ZIP |
| 格式 | 预设、自定义样式 JSON、参考 Word 样式提取、页眉页脚/页码等 Word 导出设置 |
| 语义 | 标题、列表、任务列表、表格、图片、脚注、链接、Callout、Frontmatter、标签、内部链接和嵌入降级 |

## 重要边界

- `.doc` 二进制格式不支持，请先另存为 `.docx`。
- 不做云端存储、账号、协作、PPT/PDF，也不把文件上传服务器。
- Excel 公式目前按导入时的计算结果进入 Markdown；Markdown 本身不保存 Excel 的合并、列宽、行高等全部视觉元数据。
- Word 页眉页脚、分节、目录域、浮动布局、复杂 OOXML 和 OMML 公式的完整可逆回写仍有边界，降级会写入报告。
- WPS 与 Word 的排版引擎不同。宽表已加入固定列宽、紧凑规则和显式单元格边距；用户已确认导出后出现的前导空隙属于 WPS 渲染差异，不是 Markdown 多了空格。不要通过在 Markdown 表格中添加空格来修复这类渲染差异。

## 文档入口

- [使用说明](docs/使用说明.md)：给实际使用者的操作流程、格式映射和故障排查。
- [文档导航](docs/文档导航.md)：按受众说明每份文档的职责，避免把历史方案当成现行规则。
- [项目定位](docs/项目定位.md)：产品目标、用户、边界和成功标准。
- [开发文档](docs/开发文档.md)：技术结构、实现范围、验证命令和已知限制。
- [项目总则](docs/项目总则.md)：开发和维护时必须遵守的五条铁律。
- [设计规范](docs/设计规范.md)：界面 Token 和默认 Word 结果视图规范。
- [交接表](docs/交接表.md)：当前开发状态、验收基线和下一批任务。
- [更新日志](docs/更新日志.md)：按批次记录功能、验证数字和有意变更。
- [历史归档](docs/archive/README.md)：旧规划稿和原型，仅作追溯，不作为当前实现依据。

## 开发验证

开发测试依赖只放在 `devDependencies`，产品运行时仍是单文件 HTML，运行库副本在 `vendor/`。重新生成夹具并运行全量回归：

```powershell
npm.cmd run qa:all
```

刷新锁定版本运行库：

```powershell
npm.cmd run vendor
```

报告写入 `tools/output/v4-report.json`，测试截图和导出产物写入 `tools/output/playwright/v4/` 等版本目录。
