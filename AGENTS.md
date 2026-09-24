# MarkPivot · Agent 指南

## 项目定位

产品入口是根目录的 `MarkPivot.html`。这是一个本地优先的原生 JS 单文件工具。面向用户的主流程是文档规范化（上传 Word → 选择规范 → 微调 → 导出）；Markdown 仍是内部语义层，不要做 Excel↔Word 直转入口。AI 仅用于用户主动发起的文章格式匹配，不生成或改写正文；调用时文章会发送到 DeepSeek，API Key 默认仅保存在当前浏览器会话。

转向依据：根目录 `产品定位.md`、`开发迭代.md`、`样式规范.md`。旧互转工作台文档在 `docs/`，冲突时以根目录三份新文档和当前代码为准。

## 启动与验证

- 手工启动：`python -m http.server 8765 --bind 127.0.0.1`
- 全量回归：`npm.cmd run qa:all`
- 夹具生成：`npm.cmd run fixtures`
- 刷新本地运行库：`npm.cmd run vendor`
- QA 报告：`tools/output/v4-report.json`

## 目录约定

- `docs/`：现役用户、产品、开发、设计、交接和日志文档。
- `docs/archive/`：旧规划稿和原型，只用于历史追溯。
- `vendor/`：锁定版本运行库 UMD，与 HTML 一起分发。
- `templates/`：解压即用的参考 Word 样例。
- `tools/fixtures/`：可重复测试输入。
- `tools/output/`：自动生成的报告、截图和导出产物，不是产品源码。

## 维护纪律

- 先读根目录 `产品定位.md`、`开发迭代.md`、`样式规范.md`，以及 `docs/项目总则.md`、`docs/交接表.md`。
- 保持 Markdown 为内部内容真源；格式预设、预览和导出不得改写 Markdown 语义。
- 不新增 Excel↔Word 直转入口；有损转换必须进入报告。
- 修改转换逻辑或主流程后运行 `npm.cmd run qa:all`。QA 使用 `?qa=` 进入旧工作台回归；产品默认进入新任务流程。
- 每次开发完成同步 `docs/开发文档.md`、`docs/交接表.md` 和 `docs/更新日志.md`；用户可见操作变化同步 `docs/使用说明.md`。
- 归档使用移动和重链，不删除仍有追溯价值的材料；清理候选先汇报。

## 当前状态

M0 + M1 + M2 + M3 + M4 + M5 + M6 + M7 + M8 + M9 + M10 + M11 + M12 + M13 + M14 + M15 已落地。M14 完成工作台暖白橙视觉精修、AI 窗口避让与窄屏适配；M15 增加按 Markdown 标题生成的左侧文档大纲与预览定位；AI 首期仅支持 DeepSeek；规范库后续（版本 / 标签 / 搜索 / 最近使用）尚未做。
