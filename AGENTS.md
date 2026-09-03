# 文档互转工作台 · Agent 指南

## 项目定位

产品入口是根目录的 `文档互转工作台.html`。这是一个本地优先的原生 JS 单文件工具，Markdown 是唯一语义层，Word 和 Excel 是结果视图。

## 启动与验证

- 手工启动：`python -m http.server 8765 --bind 127.0.0.1`
- 全量回归：`npm.cmd run qa:all`
- 夹具生成：`npm.cmd run fixtures`
- QA 报告：`tools/output/v4-report.json`

## 目录约定

- `docs/`：现役用户、产品、开发、设计、交接和日志文档。
- `docs/archive/`：旧规划稿和原型，只用于历史追溯。
- `tools/fixtures/`：可重复测试输入。
- `tools/output/`：自动生成的报告、截图和导出产物，不是产品源码。

## 维护纪律

- 先读 `docs/项目总则.md`、`docs/交接表.md` 和 `docs/设计规范.md`，再改代码。
- 保持 Markdown 为内容真源；格式预设、预览和导出不得改写 Markdown 语义。
- 不新增 Excel↔Word 直转入口；有损转换必须进入报告。
- 修改转换逻辑后运行静态检查、真实文件导入/导出/回导和 `npm.cmd run qa:all`。
- 每次开发完成同步 `docs/开发文档.md`、`docs/交接表.md` 和 `docs/更新日志.md`；用户可见操作变化同步 `docs/使用说明.md`。
- 归档使用移动和重链，不删除仍有追溯价值的材料；清理候选先汇报。

## 当前状态

V4.1 已完成，自动化基线为 `45 PASS / 0 FAIL / 0 SKIP`。下一批是 M3.4：大文件分片、Worker 解析、进度和取消操作。
