import { createServer } from 'node:http';
import { inflateRawSync } from 'node:zlib';
import { existsSync } from 'node:fs';
import { copyFile, mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import ExcelJS from 'exceljs';
import { VENDOR_LIBRARIES } from './vendor-sync.mjs';

const toolsDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)));
const projectRoot = path.resolve(toolsDir, '..');
const appFile = path.join(projectRoot, 'MarkPivot.html');
const fixtureDir = path.join(toolsDir, 'fixtures');
const baselineDir = path.join(toolsDir, 'baseline');
const artifactDir = path.join(toolsDir, 'output', 'playwright', 'v4');
const reportPath = path.join(toolsDir, 'output', 'v4-report.json');
const fixturePaths = {
  md: path.join(fixtureDir, 'sample.md'),
  cover: path.join(fixtureDir, 'cover.md'),
  v4Table: path.join(fixtureDir, 'v4-table.md'),
  v4WideTable: path.join(fixtureDir, 'v4-wide-table.md'),
  docx: path.join(fixtureDir, 'sample.docx'),
  referenceDocx: path.join(fixtureDir, 'reference-template.docx'),
  manualHeading: path.join(fixtureDir, 'manual-heading.docx'),
  largeDocx: path.join(fixtureDir, 'large-doc.docx'),
  xlsx: path.join(fixtureDir, 'sample.xlsx'),
};
const screenshotPaths = {
  markdown: path.join(baselineDir, 'md-preview.png'),
  word: path.join(baselineDir, 'word-preview.png'),
  excel: path.join(baselineDir, 'excel-grid.png'),
};
const v2MobileScreenshotPath = path.join(artifactDir, 'v2-mobile-scroll.png');
const v2StyleScreenshotPath = path.join(artifactDir, 'v2-style-drawer.png');
const f3CoverDesktopScreenshotPath = path.join(artifactDir, 'f3-cover-desktop.png');
const f3CoverMobileScreenshotPath = path.join(artifactDir, 'f3-cover-mobile.png');
const v4TableDesktopScreenshotPath = path.join(artifactDir, 'v4-table-desktop.png');
const v4TableMobileScreenshotPath = path.join(artifactDir, 'v4-table-mobile.png');
const v4WideTableScreenshotPath = path.join(artifactDir, 'v4-wide-table-after-fix.png');
const m1Dir = path.join(toolsDir, 'output', 'playwright', 'm1');
const m1HomeScreenshotPath = path.join(m1Dir, 'home.png');
const m1ChoiceScreenshotPath = path.join(m1Dir, 'choice.png');
const m1WorkspaceScreenshotPath = path.join(m1Dir, 'workspace.png');
const m2Dir = path.join(toolsDir, 'output', 'playwright', 'm2');
const m2WorkspaceScreenshotPath = path.join(m2Dir, 'workspace.png');
const m4Dir = path.join(toolsDir, 'output', 'playwright', 'm4');
const m4ExtractScreenshotPath = path.join(m4Dir, 'extract-summary.png');
const m4WorkspaceScreenshotPath = path.join(m4Dir, 'workspace.png');
const m5Dir = path.join(toolsDir, 'output', 'playwright', 'm5');
const m5WorkspaceScreenshotPath = path.join(m5Dir, 'workspace.png');
const m5ReviewScreenshotPath = path.join(m5Dir, 'review.png');
const m6Dir = path.join(toolsDir, 'output', 'playwright', 'm6');
const m6WorkspaceScreenshotPath = path.join(m6Dir, 'workspace.png');
const m7Dir = path.join(toolsDir, 'output', 'playwright', 'm7');
const m7WorkspaceScreenshotPath = path.join(m7Dir, 'workspace.png');
const m8Dir = path.join(toolsDir, 'output', 'playwright', 'm8');
const m8ExportScreenshotPath = path.join(m8Dir, 'export.png');
const m9Dir = path.join(toolsDir, 'output', 'playwright', 'm9');
const m9WorkspaceScreenshotPath = path.join(m9Dir, 'workspace.png');
const m9JobScreenshotPath = path.join(m9Dir, 'job.png');
const m10Dir = path.join(toolsDir, 'output', 'playwright', 'm10');
const m10HomeScreenshotPath = path.join(m10Dir, 'home.png');
const m11Dir = path.join(toolsDir, 'output', 'playwright', 'm11');
const m11LibraryScreenshotPath = path.join(m11Dir, 'library.png');
const m11EditorScreenshotPath = path.join(m11Dir, 'editor.png');
const vendorDir = path.join(projectRoot, 'vendor');
const productTemplatePath = path.join(projectRoot, 'templates', 'reference-template.docx');
const exportedPaths = {
  docx: path.join(artifactDir, 'exported-sample.docx'),
  v1CompatibilityDocx: path.join(artifactDir, 'v1-compatibility.docx'),
  v2EmphasisDocx: path.join(artifactDir, 'v2-emphasis.docx'),
  f3CoverDocx: path.join(artifactDir, 'f3-cover.docx'),
  v4TableDocx: path.join(artifactDir, 'v4-table.docx'),
  v4WideTableDocx: path.join(artifactDir, 'v4-wide-table.docx'),
  frontmatterDocx: path.join(artifactDir, 'frontmatter-sample.docx'),
  xlsx: path.join(artifactDir, 'exported-sample.xlsx'),
  batchMdZip: path.join(artifactDir, 'batch-md.zip'),
  batchWordZip: path.join(artifactDir, 'batch-word.zip'),
  batchExcelZip: path.join(artifactDir, 'batch-excel.zip'),
  batchCsvZip: path.join(artifactDir, 'batch-csv.zip'),
  taskZip: path.join(artifactDir, 'task-export.zip'),
  taskZip20: path.join(artifactDir, 'task-export-20.zip'),
  taskCancelZip: path.join(artifactDir, 'task-export-cancel.zip'),
};
const invalidDocxPath = path.join(artifactDir, 'invalid.docx');

const checks = [];
const screenshotCreated = new Set();
const browserEvents = {
  consoleErrors: [],
  consoleWarnings: [],
  pageErrors: [],
  httpErrors: [],
};

function pass(name, detail) {
  checks.push({ name, status: 'PASS', detail: detail || '通过' });
  return detail || '通过';
}

function fail(name, error) {
  const detail = error instanceof Error ? error.message : String(error);
  checks.push({ name, status: 'FAIL', detail });
  return detail;
}

function skip(name, detail) {
  checks.push({ name, status: 'SKIP', detail: detail || '跳过' });
}

async function check(name, action) {
  try {
    const detail = await action();
    return pass(name, detail);
  } catch (error) {
    return fail(name, error);
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function count(selector, page) {
  return page.locator(selector).count();
}

function textContent(selector, page) {
  return page.locator(selector).textContent();
}

function formatBytes(bytes) {
  return `${bytes} B / ${(bytes / 1024).toFixed(1)} KB`;
}

function mimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.md': 'text/markdown; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }[ext] || 'application/octet-stream';
}

function makeStaticServer() {
  return createServer(async (request, response) => {
    try {
      const rawPath = decodeURIComponent((request.url || '/').split('?')[0]);
      if (rawPath === '/favicon.ico') {
        response.writeHead(204);
        response.end();
        return;
      }
      const relativePath = rawPath === '/' ? 'MarkPivot.html' : rawPath.replace(/^[/\\]+/, '');
      const targetPath = rawPath === '/images/sample.png'
        ? path.join(fixtureDir, 'images', 'sample.png')
        : path.resolve(projectRoot, relativePath);
      const relativeToRoot = path.relative(projectRoot, targetPath);
      if (relativeToRoot.startsWith('..') || path.isAbsolute(relativeToRoot)) {
        response.writeHead(403);
        response.end('Forbidden');
        return;
      }
      const body = await readFile(targetPath);
      response.writeHead(200, {
        'Content-Type': mimeType(targetPath),
        'Cache-Control': 'no-store',
      });
      response.end(body);
    } catch (error) {
      response.writeHead(error?.code === 'ENOENT' ? 404 : 500, { 'Content-Type': 'text/plain; charset=utf-8' });
      response.end(error?.code === 'ENOENT' ? 'Not found' : String(error));
    }
  });
}

async function listen(server) {
  const requestedPort = Number(process.env.QA_PORT || 8766);
  return new Promise((resolve, reject) => {
    const onError = (error) => {
      server.off('listening', onListening);
      if (error.code === 'EADDRINUSE' && requestedPort !== 0) {
        server.once('error', reject);
        server.listen(0, '127.0.0.1');
        return;
      }
      reject(error);
    };
    const onListening = () => {
      server.off('error', onError);
      const address = server.address();
      resolve(address.port);
    };
    server.once('error', onError);
    server.once('listening', onListening);
    server.listen(requestedPort, '127.0.0.1');
  });
}

function closeServer(server) {
  return new Promise((resolve) => {
    if (!server.listening) {
      resolve();
      return;
    }
    server.close(() => resolve());
  });
}

class SimpleZip {
  constructor(buffer) {
    this.buffer = buffer;
    this.entries = new Map();
    this.readCentralDirectory();
  }

  readCentralDirectory() {
    const signature = 0x06054b50;
    let eocd = -1;
    for (let offset = this.buffer.length - 22; offset >= Math.max(0, this.buffer.length - 22 - 0xffff); offset -= 1) {
      if (this.buffer.readUInt32LE(offset) === signature) {
        eocd = offset;
        break;
      }
    }
    if (eocd < 0) throw new Error('导出文件不是有效 ZIP/DOCX 包：缺少 EOCD。');
    const entryCount = this.buffer.readUInt16LE(eocd + 10);
    const centralSize = this.buffer.readUInt32LE(eocd + 12);
    const centralOffset = this.buffer.readUInt32LE(eocd + 16);
    let offset = centralOffset;
    const centralEnd = centralOffset + centralSize;
    for (let index = 0; index < entryCount && offset < centralEnd; index += 1) {
      if (this.buffer.readUInt32LE(offset) !== 0x02014b50) {
        throw new Error(`ZIP 中央目录损坏，偏移 ${offset}。`);
      }
      const method = this.buffer.readUInt16LE(offset + 10);
      const compressedSize = this.buffer.readUInt32LE(offset + 20);
      const uncompressedSize = this.buffer.readUInt32LE(offset + 24);
      const nameLength = this.buffer.readUInt16LE(offset + 28);
      const extraLength = this.buffer.readUInt16LE(offset + 30);
      const commentLength = this.buffer.readUInt16LE(offset + 32);
      const localOffset = this.buffer.readUInt32LE(offset + 42);
      const name = this.buffer.subarray(offset + 46, offset + 46 + nameLength).toString('utf8');
      this.entries.set(name, { method, compressedSize, uncompressedSize, localOffset });
      offset += 46 + nameLength + extraLength + commentLength;
    }
  }

  has(name) {
    return this.entries.has(name);
  }

  read(name) {
    const entry = this.entries.get(name);
    if (!entry) throw new Error(`ZIP 中缺少 ${name}。`);
    const localOffset = entry.localOffset;
    if (this.buffer.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`ZIP 本地文件头损坏：${name}。`);
    }
    const nameLength = this.buffer.readUInt16LE(localOffset + 26);
    const extraLength = this.buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + nameLength + extraLength;
    const compressed = this.buffer.subarray(dataStart, dataStart + entry.compressedSize);
    if (entry.method === 0) return compressed;
    if (entry.method === 8) return inflateRawSync(compressed);
    throw new Error(`不支持的 ZIP 压缩方法 ${entry.method}：${name}。`);
  }

  readText(name) {
    return this.read(name).toString('utf8');
  }
}

function zipSummary(buffer) {
  const zip = new SimpleZip(buffer);
  return {
    zip,
    names: [...zip.entries.keys()],
    documentXml: zip.has('word/document.xml') ? zip.readText('word/document.xml') : '',
    numberingXml: zip.has('word/numbering.xml') ? zip.readText('word/numbering.xml') : '',
  };
}

function xmlParagraphs(xml) {
  return [...String(xml || '').matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
}

function waitForDownload(page, selector, outputPath, timeout = 20_000) {
  return Promise.all([
    page.waitForEvent('download', { timeout }),
    page.locator(selector).click(),
  ]).then(async ([download]) => {
    await download.saveAs(outputPath);
    return outputPath;
  });
}

async function clickWithDialog(page, selector, accept, expectedText) {
  let dialogMessage = '';
  const dialogPromise = page.waitForEvent('dialog', { timeout: 5_000 }).then(async (dialog) => {
    dialogMessage = dialog.message();
    if (accept) await dialog.accept();
    else await dialog.dismiss();
    if (expectedText && !dialogMessage.includes(expectedText)) {
      throw new Error(`对话框文案不包含“${expectedText}”：${dialogMessage}`);
    }
    return dialogMessage;
  });
  await Promise.all([page.locator(selector).click(), dialogPromise]);
  return dialogMessage;
}

async function installDependencyReplay(page) {
  const cache = new Map();
  let primaryDown = false;
  let externalOffline = false;
  let vendorDown = false;
  await page.route('**/*', async (route) => {
    const url = route.request().url();
    if (url.includes('fonts.googleapis.com')) {
      await route.fulfill({ status: 200, contentType: 'text/css', body: '' });
      return;
    }
    if (url.includes('fonts.gstatic.com')) {
      await route.fulfill({ status: 200, contentType: 'font/woff2', body: Buffer.alloc(0) });
      return;
    }
    const isVendor = /\/vendor\/[^/?#]+\.js(?:$|\?)/.test(url);
    if (isVendor) {
      if (vendorDown) {
        await route.fulfill({ status: 200, contentType: 'application/javascript; charset=utf-8', body: '' });
        return;
      }
      await route.continue();
      return;
    }
    const isPrimary = url.includes('cdn.jsdelivr.net');
    const isDependency = isPrimary || url.includes('unpkg.com');
    if (!isDependency) {
      await route.continue();
      return;
    }
    if (externalOffline || (primaryDown && isPrimary)) {
      await route.fulfill({ status: 200, contentType: 'application/javascript; charset=utf-8', body: '' });
      return;
    }
    try {
      let body = cache.get(url);
      let contentType = 'application/javascript; charset=utf-8';
      if (!body) {
        const response = await fetch(url);
        body = Buffer.from(await response.arrayBuffer());
        contentType = response.headers.get('content-type') || contentType;
        cache.set(url, { body, contentType });
      } else {
        contentType = body.contentType;
        body = body.body;
      }
      await route.fulfill({ status: 200, contentType, body });
    } catch (error) {
      await route.fulfill({ status: 200, contentType: 'application/javascript; charset=utf-8', body: '' });
    }
  });
  return {
    setPrimaryDown(value) {
      primaryDown = Boolean(value);
    },
    setOffline(value) {
      externalOffline = Boolean(value);
    },
    setVendorDown(value) {
      vendorDown = Boolean(value);
    },
  };
}

async function waitForApp(page, { navigate = true } = {}) {
  if (navigate) await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
  await page.waitForFunction(
    () => {
      const node = document.querySelector('#libraryStatus');
      return node && !node.textContent.includes('正在检查');
    },
    null,
    { timeout: 30_000 },
  );
  await page.waitForFunction(
    () => !!window.JSZip && !!window.docx && !!window.XLSX && !!window.markdownit,
    null,
    { timeout: 30_000 },
  );
  await page.waitForFunction(
    () => document.querySelector('#libraryStatus')?.textContent.includes('依赖就绪'),
    null,
    { timeout: 5_000 },
  );
  await page.waitForTimeout(100);
  return page.locator('#libraryStatus').textContent();
}

async function startTaskToExport(page, files) {
  const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
  await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
  await waitForApp(page, { navigate: false });
  await page.locator('#homeInput').setInputFiles(files);
  const expected = Array.isArray(files) ? files.length : 1;
  await page.waitForFunction((count) => window.__WB_M1__?.getTask()?.fileCount >= count, expected, { timeout: 8_000 });
  await page.locator('#homeNextBtn').click();
  await page.locator('#choiceBlankBtn').click();
  await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
  await page.waitForFunction(() => /已载入|需确认|识别失败|导入失败/.test(document.querySelector('#fileRailList')?.textContent || ''), null, { timeout: 20_000 });
  await page.locator('#taskExportBtn').click();
  await page.locator('#exportView').waitFor({ state: 'visible', timeout: 8_000 });
  assert((await page.evaluate(() => window.__WB_M8__?.snapshot()?.view)) === 'export', '未进入导出页。');
  assert((await page.locator('#batchDrawer').getAttribute('class') || '').includes('open') === false, '任务导出打开了 F1 批量抽屉。');
}

async function startTaskToWorkspace(page, files) {
  const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
  await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
  await waitForApp(page, { navigate: false });
  await page.locator('#homeInput').setInputFiles(files);
  const expected = Array.isArray(files) ? files.length : 1;
  await page.waitForFunction((count) => window.__WB_M1__?.getTask()?.fileCount >= count, expected, { timeout: 8_000 });
  await page.locator('#homeNextBtn').click();
  await page.locator('#choiceBlankBtn').click();
  await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
  await page.waitForFunction(() => /已载入|需确认|识别失败|导入失败/.test(document.querySelector('#fileRailList')?.textContent || ''), null, { timeout: 90_000 });
}

async function waitForDependencySettled(page) {
  await page.waitForFunction(
    () => {
      const state = window.__WB_DEPENDENCY_STATE__;
      const status = document.querySelector('#libraryStatus')?.textContent || '';
      return state && state.loading === false && !status.includes('依赖加载中');
    },
    null,
    { timeout: 30_000 },
  );
  return page.locator('#libraryStatus').textContent();
}

async function checkStaticContracts() {
  const [htmlStat, packageJson, html] = await Promise.all([
    stat(appFile),
    readFile(path.join(projectRoot, 'package.json'), 'utf8').then(JSON.parse),
    readFile(appFile, 'utf8'),
  ]);
  await check('单文件产品入口存在', () => {
    assert(htmlStat.size > 100_000, `HTML 体积异常：${formatBytes(htmlStat.size)}。`);
    assert(!/\b(?:import|require)\s*\(/.test(html), '产品 HTML 发现运行时 import/require。');
    return `MarkPivot.html · ${formatBytes(htmlStat.size)}`;
  });
  await check('运行库版本锁定与离线缓存契约', () => {
    assert(html.includes('window.__WB_DEPENDENCY_LOCK__'), '缺少运行库版本锁定表。');
    assert(html.includes("local:'vendor/") && html.includes('vendor/jszip-3.10.1.min.js'), '缺少本地 vendor 路径。');
    assert(html.includes('cdn.jsdelivr.net') && html.includes('unpkg.com'), '缺少 jsDelivr 主源或 unpkg 备源。');
    assert(html.includes('document-workbench-dependency-cache-v1:') && html.includes('localStorage.setItem'), '缺少 localStorage 运行库缓存。');
    assert(html.includes("version:'3.10.1'") && html.includes("version:'8.5.0'") && html.includes("version:'0.18.5'"), '核心运行库版本未锁定。');
    return 'vendor 优先 · jsDelivr/unpkg 后备 · localStorage 缓存 · 版本锁定';
  });
  await check('工具依赖只进入 devDependencies', () => {
    const dependencies = Object.keys(packageJson.dependencies || {});
    const devDependencies = Object.keys(packageJson.devDependencies || {});
    const allowed = new Set(['playwright', 'docx', 'exceljs']);
    assert(dependencies.length === 0, `dependencies 不应有运行时包：${dependencies.join(', ')}`);
    assert(devDependencies.every((name) => allowed.has(name)), `发现未登记工具依赖：${devDependencies.filter((name) => !allowed.has(name)).join(', ')}`);
    assert(allowed.size === devDependencies.length, `devDependencies 预期为 playwright/docx/exceljs，实际为 ${devDependencies.join(', ')}`);
    return `dependencies={} · devDependencies={${devDependencies.join(', ')}}`;
  });
  await check('真实 QA 夹具完整', async () => {
    const entries = await Promise.all(Object.entries(fixturePaths).map(async ([key, filePath]) => {
      assert(existsSync(filePath), `${key} 夹具不存在：${filePath}，请先运行 npm run fixtures。`);
      const fileStat = await stat(filePath);
      assert(fileStat.size > 0, `${key} 夹具为空。`);
      return `${key}=${formatBytes(fileStat.size)}`;
    }));
    assert(existsSync(path.join(fixtureDir, 'images', 'sample.png')), '图片夹具不存在。');
    return entries.join(' · ');
  });
  await check('V1 静态契约', () => {
    assert(html.includes('.app{height:100dvh;min-height:0;overflow:hidden}'), '缺少桌面工作台锁高契约。');
    assert(html.includes('.excel-stage[hidden]{display:none!important}'), '缺少 Excel 隐藏视图契约。');
    assert(html.includes('--body-first-line-indent') && html.includes('styleBodyFirstLineIndent'), '缺少正文首行缩进样式契约。');
    assert(html.includes('ascii:\'SimSun\'') && html.includes('hAnsi:\'SimSun\'') && html.includes('eastAsia:\'SimSun\''), '代码块没有三通道宋体契约。');
    assert(html.includes('firstLineChars=\"200\"') && html.includes('w:wordWrap w:val=\"0\"'), '缺少 DOCX 首行缩进或不折行契约。');
    return '桌面锁高 · 移动端释放 · 2 字符缩进 · SimSun 三通道 · WPS 不折行';
  });
  await check('F2 文件关联静态契约', () => {
    assert(html.includes('showOpenFilePicker'), '缺少 File System Access API 打开入口。');
    assert(html.includes('createWritable') && html.includes('getFile'), '缺少关联文件读写 API。');
    assert(html.includes('refreshMdBtn') && html.includes('saveMdBtn') && html.includes('autoSyncToggle'), '缺少刷新、保存或自动同步控件。');
    assert(html.includes('lastModified') && html.includes('本地有未保存修改') && html.includes('磁盘上被修改'), '缺少双向冲突保护契约。');
    return 'showOpenFilePicker · getFile/createWritable · lastModified 冲突保护 · 2 秒自动同步';
  });
  await check('V2 Word 提取与强调静态契约', () => {
    assert(html.includes('themeColor') && html.includes('themeShade') && html.includes('themeTint'), '缺少 Word 主题色引用或 shade/tint 解析契约。');
    assert(html.includes('OOXML 默认黑') && html.includes('style-source-badge') && html.includes('style-clear-field'), '缺少默认黑、来源徽标或表头清除控件契约。');
    assert(html.includes('STYLE_CONTROL_OPTIONS') && html.includes('__custom__') && html.includes('style-color-picker'), '缺少字体/字号/颜色控件下拉与自定义路径。');
    assert(html.includes('emphasisMark') && html.includes('text-emphasis') && html.includes("type:'dot'"), '缺少三端着重号输出契约。');
    return 'OOXML 默认黑 · themeColor shade/tint · 字段来源 · 下拉/取色器 · emphasisMark dot';
  });
  await check('F3 封面页静态契约', () => {
    assert(html.includes('md-cover') && html.includes('data-auto-cover-break') && html.includes('coverDetected'), '缺少封面节点、自动分页标记或导出上下文。');
    assert(html.includes('AlignmentType.CENTER') && html.includes('new D.PageBreak'), '缺少封面 Word 居中或强制分页导出契约。');
    assert(html.includes('F3') || html.includes('COVER ENDS'), '缺少封面视觉标识。');
    return '正文开头识别 · 居中封面 · 自动分页 · Word 居中与 PageBreak';
  });
  await check('V4 表格布局与缩进静态契约', () => {
    assert(html.includes('tableLayoutMetrics') && html.includes('document.createElement(\'colgroup\')'), '缺少预览表格列宽模型或 colgroup 注入。');
    assert(html.includes('columnWidths') && html.includes('TableLayoutType.FIXED') && html.includes('widthType'), '缺少 DOCX gridCol/tcW/fixed 列宽契约。');
    assert(html.includes('WBBodyIndent') && html.includes('firstLineChars="200"') && !html.includes('paragraphProps.indent={firstLine:ctx.bodyIndentTwips}'), '正文缩进仍由 firstLine 双写。');
    assert(html.includes('left:420,hanging:420'), '列表编号缩进没有校准到 V4 量级。');
    return 'colgroup · gridCol/tcW · fixed · firstLineChars 单写 · 列表 left=420/hanging=420';
  });
  await check('M0 入口与 M1 UI Foundation 契约', () => {
    assert(path.basename(appFile) === 'MarkPivot.html', `产品入口不是 MarkPivot.html：${appFile}`);
    assert(html.includes('--color-brand:#F97316'), '缺少品牌橙 Token。');
    assert(html.includes('--color-bg:#F7F6F3') && html.includes('--radius-card:10px'), '缺少暖白画布或卡片圆角 Token。');
    assert(html.includes('id="homeView"') && html.includes('下一步：选择规范'), '缺少首页上传主流程。');
    assert(html.includes('id="choiceView"') && html.includes('上传模板') && html.includes('从空白开始'), '缺少规范选择三入口。');
    assert(html.includes('id="fileRail"') && html.includes('id="specRail"') && html.includes('data-spec-scope'), '缺少工作台三栏或作用域开关。');
    assert(html.includes('data-app-view="home"') && html.includes('data-app-view="specs"') && !/<button[^>]*data-app-view="workspace"/.test(html), '一级导航仍把工作台当作入口。');
    return 'MarkPivot.html · 暖白橙 Token · 首页/选规范/三栏工作台 · 一级页不含工作台';
  });
  await check('M2 Task 与 M3 Specification 契约', () => {
    assert(html.includes('function createTask()') && html.includes('function endTask()') && html.includes('fileOverrides'), '缺少 Task 生命周期。');
    assert(html.includes('markpivot-spec-library-v1') && html.includes('id="specSaveLibraryBtn"') && html.includes('离开当前任务'), '缺少规范库隔离或离开确认。');
    assert(html.includes('data-spec-path="page.paper"') && html.includes('data-spec-path="heading1.bold"') && html.includes('data-spec-path="header.text"') && html.includes('data-spec-path="footer.pageNumber"') && html.includes('data-spec-path="table.keepTogether"'), 'Specification 字段未绑定到规范栏。');
    assert(html.includes('window.__WB_M2__') && html.includes('window.__WB_M3__') && html.includes('roundtripOk'), '缺少 M2/M3 探针。');
    assert(!html.includes('data-spec-style='), '规范栏仍直接绑定散落 CSS 变量。');
    return 'Task 生命周期 · 规范库隔离 · Specification JSON/UI 绑定 · 探针';
  });
  await check('M4 Spec Extractor 契约', () => {
    assert(html.includes('extractSpecificationFromZip') && html.includes('extractSpecificationFromFile(file)'), '缺少 DOCX → Specification 提取器。');
    assert(html.includes('id="extractSummary"') && html.includes('id="extractExtractedList"') && html.includes('id="extractMissingList"') && html.includes('进入工作台继续修改'), '缺少提取摘要 UI。');
    assert(html.includes('window.__WB_M4__') && html.includes('id="specExtractBanner"'), '缺少 M4 探针或规范栏提取横幅。');
    assert(!/templateInput[\s\S]{0,1200}specFromCss\(\)/.test(html), '模板上传仍经 CSS 往返猜测未提取项。');
    return 'DOCX 直出 Specification · 已提取/未提取摘要 · 禁止 CSS 往返';
  });
  await check('M5 Semantic Mapper 契约', () => {
    assert(html.includes('mapSemanticsFromZip') && html.includes('mapSemanticsFromFile(file)'), '缺少 DOCX → 语义块识别器。');
    assert(html.includes('id="semanticMap"') && html.includes('semanticType') && html.includes('需确认'), '缺少语义列表或异常文件状态。');
    assert(html.includes('window.__WB_M5__') && html.includes('处待确认') && html.includes('处低置信度，未自动改写'), '缺少 M5 探针或低置信度提示。');
    assert(html.includes("source==='heuristic'&&conf<0.85") && html.includes('Math.min(0.84'), '启发式标题未封顶或低置信度仍可能静默改写。');
    return '确定性规则+启发式 · 用户可改 · 低置信度不静默改写';
  });
  await check('M6 工作台预览契约', () => {
    assert(html.includes('id="previewToolbar"') && html.includes('data-preview-mode="source"') && html.includes('data-preview-mode="effect"'), '缺少原文 / 当前效果切换。');
    assert(html.includes('id="previewZoomIn"') && html.includes('id="previewPageLabel"') && html.includes('第 1 / 1 页'), '缺少缩放或页码。');
    assert(html.includes('window.__WB_M6__') && html.includes('function paperCssSize') && html.includes('rememberTaskMarkdown'), '缺少 M6 探针、纸面尺寸或切文件 Markdown 记忆。');
    return '原文/当前效果 · 缩放页码 · 切文件保留 MD';
  });
  await check('M7 Rule Engine 契约', () => {
    assert(html.includes('function compileRules') && html.includes('function renderRules') && html.includes('function paintRuleStyles') && html.includes('function applyRuleEngine'), '缺少 Rule Engine 编译或涂绘。');
    assert(html.includes('window.__WB_M7__') && html.includes('function promoteRuleNodes') && html.includes('decorateExportHolder'), '缺少 M7 探针或导出提升。');
    assert(html.includes('is-rule-heading1') && html.includes("previewMode==='source'"), '缺少标题涂绘或原文模式清除。');
    return 'compile/render/paint · 导出提升 · 原文不涂绘';
  });
  await check('M8 批量导出契约', () => {
    assert(html.includes('id="exportView"') && html.includes('id="exportAllBtn"') && html.includes('导出全部文件'), '缺少任务导出页。');
    assert(html.includes('function runTaskExport') && html.includes('function buildTaskWordBlob') && html.includes('importWordEngine'), '缺少任务导出沙箱或同一套导出函数。');
    assert(html.includes('window.__WB_M8__') && html.includes('specification.json') && html.includes('处理记录.txt'), '缺少 M8 探针或 ZIP 规范/记录。');
    assert(!/addBatchFiles\(task\.files/.test(html) && html.includes("setAppView('export')"), '任务导出仍复用 F1 批量抽屉。');
    return '导出页 · 单多档同一套 · 失败隔离 · 不走 F1 抽屉';
  });
  await check('M9 性能与大文件契约', () => {
    assert(html.includes('function jobPlan') && html.includes('function forEachBodyChunk') && html.includes('JOB_WORKER_SOURCE') && html.includes('function prepareAssetBlob'), '缺少分片调度或图片 Worker。');
    assert(html.includes('id="jobBanner"') && html.includes('id="jobCancelBtn"') && html.includes("previewMarkdown") && html.includes('大图已延后预览'), '缺少进度条、取消或大图延后预览。');
    assert(html.includes('window.__WB_M9__') && html.includes("tier:'50mb'") && html.includes("tier:'20mb'") && html.includes("tier:'5mb'") && html.includes("tier:'100p'"), '缺少 M9 探针或档位契约。');
    assert(html.includes('entry.blob=null') && html.includes('JOB_CANCELLED') && html.includes('这个文件导入失败，可继续处理其他文件'), '缺少内存释放、取消或任务隔离。');
    return 'Worker ping · 分片 yield · 进度取消 · 大图延后 · 失败隔离';
  });
  await check('M10 真正离线交付契约', async () => {
    assert(html.includes('window.__WB_M10__') && html.includes('function loadVendor') && html.includes('fromVendor'), '缺少 M10 探针或本地 vendor 加载。');
    assert(html.includes('本地运行库可用（可离线）') && html.includes('loadScriptTag'), '缺少离线状态或 file:// script 回退。');
    const vendorFiles = await Promise.all(VENDOR_LIBRARIES.map(async (item) => {
      const filePath = path.join(vendorDir, item.file);
      assert(existsSync(filePath), `缺少 vendor/${item.file}，请先运行 npm run vendor。`);
      const fileStat = await stat(filePath);
      assert(fileStat.size >= item.minBytes, `vendor/${item.file} 过小：${fileStat.size}`);
      assert(html.includes(item.file), `HTML 未锁定 ${item.file}`);
      return `${item.id}=${formatBytes(fileStat.size)}`;
    }));
    assert(existsSync(path.join(vendorDir, 'manifest.json')), '缺少 vendor/manifest.json。');
    assert(existsSync(productTemplatePath), '缺少 templates/reference-template.docx。');
    const templateStat = await stat(productTemplatePath);
    assert(templateStat.size > 0, '发行模板为空。');
    return `${vendorFiles.join(' · ')} · template=${formatBytes(templateStat.size)}`;
  });
  await check('M11 规范库契约', () => {
    assert(html.includes('id="libraryNewBtn"') && html.includes('id="libraryFromWordBtn"') && html.includes('id="libraryImportBtn"') && html.includes('id="libraryEditor"'), '缺少规范库新建/Word/JSON/编辑入口。');
    assert(html.includes('window.__WB_M11__') && html.includes('function putLibrarySpec') && html.includes('function importLibraryJson') && html.includes('function duplicateLibrarySpec'), '缺少规范库写入、导入或复制。');
    assert(html.includes('修改只写入规范库，不会改当前任务') && html.includes('id="specSaveLibraryBtn"'), '缺少任务隔离文案或保存为新规范。');
    assert(html.includes('data-lib-action') || html.includes('dataset.libAction'), '缺少规范卡片操作。');
    return '新建/编辑/复制/重命名/删除 · JSON · 从 Word · 不覆盖任务';
  });
}

async function runBrowserChecks(page) {
  const dependencyReplay = await installDependencyReplay(page);
  await check('浏览器页面与依赖就绪', async () => {
    await waitForApp(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForApp(page, { navigate: false });
    const libs = await page.evaluate(() => ({
      JSZip: !!window.JSZip,
      docx: !!window.docx,
      XLSX: !!window.XLSX,
      markdownit: !!window.markdownit,
    }));
    assert(Object.values(libs).every(Boolean), `页面全局依赖不完整：${JSON.stringify(libs)}`);
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const sources = Object.values(dependency?.libraries || {}).map((item) => item.source);
    assert(dependency?.locked === true && sources.length === 9, `依赖状态不完整：${JSON.stringify(dependency)}`);
    assert(sources.every((source) => source === 'vendor'), `首次未命中本地运行库：${sources.join('/')}`);
    assert(dependency.fromVendor === true && dependency.offlineCapable === true, '本地运行库未标记可离线。');
    assert(text.includes('本地运行库可用'), `首次离线提示不正确：${text}`);
    return `${text} · 全局依赖=${Object.keys(libs).join('/')} · 源=vendor`;
  });
  if (checks.at(-1)?.status !== 'PASS') {
    skip('浏览器业务流程', '页面依赖未就绪，跳过真实文件操作。');
    return;
  }

  await check('M10 无 CDN 首次启动仍可用', async () => {
    await mkdir(m10Dir, { recursive: true });
    await page.evaluate(() => localStorage.clear());
    dependencyReplay.setOffline(true);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForApp(page, { navigate: false });
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const sources = Object.values(dependency?.libraries || {}).map((item) => item.source);
    assert(sources.length === 9 && sources.every((source) => source === 'vendor'), `内网首次未走 vendor：${sources.join('/')}`);
    assert(text.includes('本地运行库可用'), `内网状态提示不正确：${text}`);
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.screenshot({ path: m10HomeScreenshotPath, fullPage: false });
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    dependencyReplay.setOffline(false);
    return `无 CDN · vendor 9/9 · ${text}`;
  });

  await check('本地 vendor 缺失时回退 jsDelivr', async () => {
    await page.evaluate(() => localStorage.clear());
    dependencyReplay.setVendorDown(true);
    dependencyReplay.setPrimaryDown(false);
    dependencyReplay.setOffline(false);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForApp(page, { navigate: false });
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const sources = Object.values(dependency?.libraries || {}).map((item) => item.source);
    assert(sources.length === 9 && sources.every((source) => source === 'jsdelivr'), `vendor 缺失后未回退 jsDelivr：${sources.join('/')}`);
    return `${text} · 源=jsDelivr`;
  });

  await check('CDN 备源故障转移与本地缓存', async () => {
    await page.evaluate(() => localStorage.clear());
    dependencyReplay.setVendorDown(true);
    dependencyReplay.setPrimaryDown(true);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForApp(page, { navigate: false });
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const sources = Object.values(dependency?.libraries || {}).map((item) => item.source);
    assert(sources.length === 9 && sources.every((source) => source === 'unpkg'), `备源未全部接管：${sources.join('/')}`);
    assert(dependency.storedCount === 9, `备源成功后缓存数量错误：${dependency.storedCount}`);
    return `${text} · 备源=unpkg · 已缓存=${dependency.storedCount}/9`;
  });

  await check('外部网络不可用时本地缓存启动', async () => {
    dependencyReplay.setVendorDown(true);
    dependencyReplay.setOffline(true);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForApp(page, { navigate: false });
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const sources = Object.values(dependency?.libraries || {}).map((item) => item.source);
    assert(dependency?.fromCache === true && sources.length === 9 && sources.every((source) => source === 'cache'), `离线缓存未完整加载：${JSON.stringify(dependency)}`);
    assert(text.includes('本机缓存可用'), `离线状态提示不正确：${text}`);
    dependencyReplay.setOffline(false);
    dependencyReplay.setVendorDown(false);
    dependencyReplay.setPrimaryDown(false);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `离线重载成功 · 缓存=9/9 · 版本已锁定`;
  });

  await check('无网且无 vendor/缓存时降级不白屏', async () => {
    await page.evaluate(() => localStorage.clear());
    dependencyReplay.setVendorDown(true);
    dependencyReplay.setOffline(true);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    const text = await waitForDependencySettled(page);
    const dependency = await page.evaluate(() => window.__WB_DEPENDENCY_STATE__);
    const buttons = await page.evaluate(() => ['exportWordBtn', 'exportExcelBtn', 'exportCsvBtn', 'exportZipBtn'].map((id) => document.getElementById(id)?.disabled));
    assert(dependency?.missing?.length === 9, `无网缺库数量错误：${JSON.stringify(dependency?.missing)}`);
    assert(buttons.every(Boolean), `无网降级按钮未全部禁用：${JSON.stringify(buttons)}`);
    assert(text.includes('部分能力不可用'), `无网降级提示缺失：${text}`);
    dependencyReplay.setOffline(false);
    dependencyReplay.setVendorDown(false);
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `无缓存降级成功 · 9 个运行库缺失可见 · 导出按钮已禁用`;
  });

  await check('Markdown 黄金样本结构', async () => {
    const source = await readFile(fixturePaths.md, 'utf8');
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(source);
    await page.waitForTimeout(350);
    await page.locator('#previewTab').click();
    await page.waitForTimeout(150);
    const headingCounts = {};
    for (let level = 1; level <= 6; level += 1) headingCounts[`h${level}`] = await count(`#markdownPreview h${level}`, page);
    assert(Object.values(headingCounts).every((value) => value === 1), `H1-H6 数量不符合黄金样本：${JSON.stringify(headingCounts)}`);
    const tableCount = await count('#markdownPreview table', page);
    const calloutCount = await count('#markdownPreview .md-callout', page);
    const footnoteCount = await count('#markdownPreview .footnote-ref', page);
    const frontmatterCount = await count('#markdownPreview .md-frontmatter', page);
    const tagCount = await count('#markdownPreview .md-tag', page);
    const internalLinkCount = await count('#markdownPreview .md-internal-link', page);
    const embedCount = await count('#markdownPreview .md-embed', page);
    const missingEmbedCount = await count('#markdownPreview .md-embed-missing', page);
    const semanticSummary = await textContent('#semanticSummaryText', page);
    assert(tableCount >= 1, 'Markdown 预览缺少表格。');
    assert(calloutCount >= 3, `Markdown 预览 Callout 不足 3 个：${calloutCount}。`);
    assert(footnoteCount >= 1, 'Markdown 预览缺少脚注引用。');
    assert(frontmatterCount === 1, `Markdown 预览缺少 Frontmatter 语义带：${frontmatterCount}。`);
    assert(tagCount >= 2, `Markdown 预览标签不足：${tagCount}。`);
    assert(internalLinkCount === 2, `Markdown 预览内部链接数量错误：${internalLinkCount}。`);
    assert(embedCount === 2 && missingEmbedCount === 1, `Markdown 嵌入降级数量错误：embed=${embedCount} · missing=${missingEmbedCount}。`);
    assert(semanticSummary.includes('Frontmatter 已识别') && semanticSummary.includes('4 个标签') && semanticSummary.includes('2 个内部链接') && semanticSummary.includes('2 个嵌入'), `语义摘要不完整：${semanticSummary}`);
    const headingAnchor = await page.locator('#markdownPreview .md-internal-link').first().getAttribute('href');
    assert(headingAnchor && await count(`#markdownPreview ${headingAnchor}`, page) === 1, `标题内部链接没有对应锚点：${headingAnchor}`);
    await page.locator('#markdownPreview').screenshot({ path: screenshotPaths.markdown });
    screenshotCreated.add(screenshotPaths.markdown);
    return `H1-H6=${Object.values(headingCounts).join('/')} · table=${tableCount} · callout=${calloutCount} · footnote=${footnoteCount} · frontmatter=${frontmatterCount} · tags=${tagCount} · links=${internalLinkCount} · embeds=${embedCount}`;
  });

  await check('Markdown 语义回写', async () => {
    await page.locator('#wordPaper').dispatchEvent('input');
    await page.waitForTimeout(350);
    const roundTrip = await page.locator('#mdEditor').inputValue();
    assert(roundTrip.trim().startsWith('---'), `Frontmatter 回写未保留：${roundTrip.slice(0, 80)}`);
    assert(roundTrip.includes('#回归') && roundTrip.includes('#文档'), '标签回写未保留。');
    assert(roundTrip.includes('[[#H2 结构|H2 结构]]') && roundTrip.includes('[[项目说明|项目说明]]'), '内部链接回写未保留。');
    assert(roundTrip.includes('![[images/missing.png]]') && roundTrip.includes('![[相关笔记#结论|相关笔记]]'), '嵌入回写未保留。');
    return 'Frontmatter、标签、内部链接和两类嵌入均回写为原始 Markdown 语法。';
  });

  await check('Frontmatter 不泄漏到 Word 正文', async () => {
    await page.locator('[data-mode="word"]').click();
    await page.waitForTimeout(150);
    assert(await count('#wordPaper .md-frontmatter', page) === 1, 'Word 预览没有保留可识别的 Frontmatter 节点。');
    await waitForDownload(page, '#exportWordBtn', exportedPaths.frontmatterDocx);
    const summary = zipSummary(await readFile(exportedPaths.frontmatterDocx));
    assert(!summary.documentXml.includes('FRONTMATTER') && !summary.documentXml.includes('title: M3.1'), 'Frontmatter 元数据被写入 Word 正文。');
    return 'Frontmatter 保留在 Markdown 语义层，Word 正文不写入界面元数据。';
  });

  await check('Markdown → Word 预览', async () => {
    await page.locator('[data-mode="word"]').click();
    await page.waitForTimeout(150);
    const headingCount = await count('#wordPaper h1, #wordPaper h2, #wordPaper h3, #wordPaper h4, #wordPaper h5, #wordPaper h6', page);
    const tableCount = await count('#wordPaper table', page);
    const calloutCount = await count('#wordPaper .md-callout', page);
    const excelStage = await page.locator('#excelStage').evaluate((node) => ({ hidden: node.hidden, display: getComputedStyle(node).display, height: node.getBoundingClientRect().height }));
    assert(headingCount >= 6, `Word 预览标题不足：${headingCount}。`);
    assert(tableCount >= 1, 'Word 预览缺少表格。');
    assert(calloutCount >= 3, `Word 预览 Callout 不足：${calloutCount}。`);
    assert(excelStage.hidden && excelStage.height === 0, 'Word 模式泄漏了 Excel 结果区：' + JSON.stringify(excelStage));
    await page.locator('#wordPaper').screenshot({ path: screenshotPaths.word });
    screenshotCreated.add(screenshotPaths.word);
    return `headings=${headingCount} · tables=${tableCount} · callouts=${calloutCount} · Excel hidden=${excelStage.hidden}/0px`;
  });

  await check('V1 样式抽屉与预览缩进同步', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = [
      '# V1 缩进回归',
      '',
      '\u3000\u3000这是一段带手打空白的正文。',
      '',
      '## 标题不缩进',
      '',
      '- 列表不缩进',
      '',
      '| 表头 | 值 |',
      '| --- | --- |',
      '| 表格单元格 | 不缩进 |',
      '',
      '> 引用不缩进',
    ].join('\n');
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(source);
    await page.waitForTimeout(350);
    await page.locator('#previewTab').click();
    await page.waitForTimeout(150);
    await page.locator('#styleBtn').click();
    assert(await page.locator('#styleBodyFirstLineIndent').isChecked(), '正文首行缩进默认没有开启。');
    const on = await page.evaluate(() => {
      const body = document.querySelector('#markdownPreview > p');
      const heading = document.querySelector('#markdownPreview h2');
      const list = document.querySelector('#markdownPreview li');
      const cell = document.querySelector('#markdownPreview td');
      return {
        bodyIndent: body ? getComputedStyle(body).textIndent : '',
        bodyText: body ? body.textContent : '',
        headingIndent: heading ? getComputedStyle(heading).textIndent : '',
        listIndent: list ? getComputedStyle(list).textIndent : '',
        cellIndent: cell ? getComputedStyle(cell).textIndent : '',
      };
    });
    assert(on.bodyIndent !== '0px' && on.bodyIndent !== '', '预览正文没有显示首行缩进：' + on.bodyIndent);
    assert(!on.bodyText.startsWith('\u3000'), '开启首行缩进后正文仍保留行首全角空白。');
    assert(on.headingIndent === '0px' && on.listIndent === '0px' && on.cellIndent === '0px', '非正文节点被错误缩进：' + JSON.stringify(on));
    await page.locator('#styleBodyFirstLineIndent').uncheck();
    await page.waitForTimeout(150);
    const off = await page.evaluate(() => {
      const body = document.querySelector('#markdownPreview > p');
      return { indent: body ? getComputedStyle(body).textIndent : '', text: body ? body.textContent : '' };
    });
    const sourceAfterOff = await page.locator('#mdEditor').inputValue();
    await page.locator('#styleBodyFirstLineIndent').check();
    await page.waitForTimeout(150);
    assert(off.indent === '0px' && sourceAfterOff.includes('\u3000\u3000'), '关闭首行缩进后行为不一致：' + JSON.stringify({ ...off, sourceHasU3000: sourceAfterOff.includes('\u3000\u3000') }));
    await page.locator('#styleDrawer [data-close="styleDrawer"]').click();
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').waitFor({ state: 'visible' });
    await page.locator('#mdEditor').fill(original);
    await page.waitForTimeout(350);
    return '默认 text-indent=' + on.bodyIndent + ' · 关闭=' + off.indent + ' · 标题/列表/表格=0 · 行首空白剥离=1';
  });

  await check('V1 桌面滚动主权与移动端释放', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const longSource = '# 长文档滚动哨兵\n\n' + Array.from({ length: 220 }, (_, index) => '第 ' + (index + 1) + ' 段：用于验证桌面面板内部滚动，不应把整个工作台撑高。').join('\n\n');
    await page.setViewportSize({ width: 1440, height: 1100 });
    await page.locator('[data-mode="word"]').click();
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(longSource);
    await page.waitForTimeout(350);
    await page.locator('#previewTab').click();
    await page.waitForTimeout(150);
    const desktop = await page.evaluate(() => {
      const root = document.documentElement;
      const preview = document.querySelector('#markdownPreview');
      const word = document.querySelector('#wordStage');
      const app = document.querySelector('.app');
      return {
        pageHeight: root.scrollHeight,
        clientHeight: root.clientHeight,
        appHeight: app ? Math.round(app.getBoundingClientRect().height) : 0,
        previewClient: preview ? preview.clientHeight : 0,
        previewScroll: preview ? preview.scrollHeight : 0,
        wordClient: word ? word.clientHeight : 0,
        wordScroll: word ? word.scrollHeight : 0,
      };
    });
    assert(desktop.pageHeight === desktop.clientHeight, '桌面页面仍可整页滚动：' + JSON.stringify(desktop));
    assert(desktop.previewScroll > desktop.previewClient, 'Markdown 预览没有内部滚动：' + JSON.stringify(desktop));
    assert(desktop.wordScroll > desktop.wordClient, 'Word 纸张没有内部滚动：' + JSON.stringify(desktop));
    const tableSource = '| 列 A | 列 B |\n| --- | --- |\n' + Array.from({ length: 100 }, (_, index) => '| ' + (index + 1) + ' | grid scroll |').join('\n');
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(tableSource);
    await page.locator('[data-mode="excel"]').click();
    await page.waitForTimeout(350);
    const grid = await page.evaluate(() => {
      const root = document.documentElement;
      const node = document.querySelector('#gridWrap');
      return {
        pageHeight: root.scrollHeight,
        clientHeight: root.clientHeight,
        client: node ? node.clientHeight : 0,
        scroll: node ? node.scrollHeight : 0,
      };
    });
    assert(grid.pageHeight === grid.clientHeight, 'Excel 网格把桌面页面撑高：' + JSON.stringify(grid));
    assert(grid.scroll > grid.client, 'Excel grid-wrap 没有内部滚动：' + JSON.stringify(grid));
    await page.setViewportSize({ width: 800, height: 900 });
    await page.locator('[data-mode="word"]').click();
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(longSource);
    await page.waitForTimeout(350);
    const mobile = await page.evaluate(() => ({
      pageHeight: document.documentElement.scrollHeight,
      clientHeight: document.documentElement.clientHeight,
      appHeight: Math.round(document.querySelector('.app').getBoundingClientRect().height),
    }));
    assert(mobile.pageHeight > mobile.clientHeight, '移动端没有释放为整页滚动：' + JSON.stringify(mobile));
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.screenshot({ path: v2MobileScreenshotPath, fullPage: false });
    await page.setViewportSize({ width: 1440, height: 1100 });
    await page.locator('#mdEditor').fill(original);
    await page.waitForTimeout(350);
    await page.locator('#sourceTab').click();
    return '桌面 page=' + desktop.pageHeight + '/' + desktop.clientHeight + ' · MD=' + desktop.previewScroll + '>' + desktop.previewClient + ' · Word=' + desktop.wordScroll + '>' + desktop.wordClient + ' · grid=' + grid.scroll + '>' + grid.client + ' · mobile page=' + mobile.pageHeight + '>' + mobile.clientHeight;
  });

  await check('V1 DOCX WPS 兼容与首行缩进 OOXML', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const fence = String.fromCharCode(96).repeat(3);
    const source = [
      '# 标题不缩进',
      '',
      '\u3000\u3000正文段落不应保留行首空白。',
      '',
      '> 引用段落不缩进',
      '',
      '| 表头 | 值 |',
      '| --- | --- |',
      '| 表格单元格 | 不缩进 |',
      '',
      fence + 'text',
      '┌' + '─'.repeat(120) + '┐',
      '│ 中文 WPS 字符画与 ASCII 长行对齐：' + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.repeat(4) + ' │',
      '└' + '─'.repeat(120) + '┘',
      fence,
    ].join('\n');
    await page.setViewportSize({ width: 1440, height: 1100 });
    await page.locator('[data-mode="word"]').click();
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').waitFor({ state: 'visible' });
    await page.locator('#mdEditor').fill(source);
    await page.waitForTimeout(350);
    await page.locator('#previewTab').click();
    await page.waitForTimeout(150);
    const output = await waitForDownload(page, '#exportWordBtn', exportedPaths.v1CompatibilityDocx);
    const summary = zipSummary(await readFile(output));
    const paragraphs = xmlParagraphs(summary.documentXml);
    const codeParagraph = paragraphs.find((paragraph) => paragraph.includes('w:pStyle w:val="WBCode"'));
    const bodyParagraph = paragraphs.find((paragraph) => paragraph.includes('正文段落不应保留行首空白'));
    const headingParagraph = paragraphs.find((paragraph) => paragraph.includes('标题不缩进'));
    const quoteParagraph = paragraphs.find((paragraph) => paragraph.includes('引用段落不缩进'));
    const cellParagraph = paragraphs.find((paragraph) => paragraph.includes('表格单元格'));
    assert(codeParagraph, '没有找到标记为 WBCode 的代码段落。');
    assert(bodyParagraph, '没有找到正文段落。');
    assert(headingParagraph && quoteParagraph && cellParagraph, '标题、引用或表格段落缺失。');
    assert(codeParagraph.includes('w:ascii="SimSun"') && codeParagraph.includes('w:hAnsi="SimSun"') && codeParagraph.includes('w:eastAsia="SimSun"'), '代码块 rFonts 三通道不是 SimSun。');
    assert(codeParagraph.includes('w:wordWrap w:val="0"'), '代码块没有写入 WPS 不折行标记。');
    const codeSizes = [...codeParagraph.matchAll(/w:sz w:val="(\d+)"/g)].map((match) => Number(match[1]));
    assert(codeSizes.length > 0 && Math.min(...codeSizes) < 18, '超宽代码行字号没有缩小：' + codeSizes.join('/'));
    assert(bodyParagraph.includes('w:firstLineChars="200"') && !bodyParagraph.includes('w:firstLine="'), '正文首行缩进不是单一 firstLineChars 属性。');
    assert(!bodyParagraph.includes('\u3000'), '导出正文仍保留 U+3000 行首空白。');
    assert(!headingParagraph.includes('w:firstLine') && !quoteParagraph.includes('w:firstLine') && !cellParagraph.includes('w:firstLine') && !codeParagraph.includes('w:firstLine'), '标题/引用/表格/代码块错误携带首行缩进。');
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').waitFor({ state: 'visible' });
    await page.locator('#mdEditor').fill(original);
    await page.waitForTimeout(350);
    return 'DOCX=' + formatBytes((await stat(output)).size) + ' · code rFonts=3 · min sz=' + Math.min(...codeSizes) + ' half-points · wordWrap=0 · firstLineChars=200 · firstLine=0 · body U+3000=0';
  });

  await check('V2 Word 模板提取与样式控件', async () => {
    await page.locator('#styleBtn').click();
    await page.locator('#styleReferenceInput').setInputFiles(fixturePaths.referenceDocx);
    await page.waitForFunction(() => document.querySelector('#styleStatus')?.textContent.includes('参考 Word'), null, { timeout: 20_000 });
    const extracted = await page.evaluate(() => {
      const sources = {};
      document.querySelectorAll('[data-style-source]').forEach((node) => {
        sources[node.dataset.styleSource] = node.textContent.trim();
      });
      const root = getComputedStyle(document.documentElement);
      const profile = window.__WB_V2__?.getLastReferenceProfile();
      return {
        body: root.getPropertyValue('--body-color').trim().toLowerCase(),
        title: root.getPropertyValue('--title-color').trim().toLowerCase(),
        h1: root.getPropertyValue('--h1-color').trim().toLowerCase(),
        h2: root.getPropertyValue('--h2-color').trim().toLowerCase(),
        h3: root.getPropertyValue('--h3-color').trim().toLowerCase(),
        table: root.getPropertyValue('--table-head').trim().toLowerCase(),
        sources,
        profile,
        sourceBadgeCount: document.querySelectorAll('#styleDrawer [data-style-source]').length,
        fontSelects: document.querySelectorAll('#styleDrawer select[data-style-select="--body-font"]').length,
        sizeSelects: document.querySelectorAll('#styleDrawer select[data-style-select]').length,
        customInputs: document.querySelectorAll('#styleDrawer input[data-style]').length,
        colorPickers: document.querySelectorAll('#styleDrawer input[data-style-picker]').length,
        clearButton: !!document.querySelector('#styleDrawer [data-style-clear="--table-head"]'),
        reportText: document.querySelector('#reportBody')?.textContent || '',
      };
    });
    assert(extracted.body === '#000000', 'Normal 无 w:color 未提取为默认黑：' + JSON.stringify(extracted));
    assert(extracted.title === '#1b2e4e', 'Title themeShade 换算不正确：' + extracted.title);
    assert(extracted.h1 === '#f1975a', 'Heading 1 themeTint 换算不正确：' + extracted.h1);
    assert(extracted.h2 === '#000000' && extracted.h3 === '#7030a0', 'Heading 2/3 颜色提取不正确：' + JSON.stringify(extracted));
    assert(extracted.table === '#698ed0', '表格首行 themeFill 换算不正确：' + extracted.table);
    assert(extracted.sources['--body-color'] === 'OOXML 默认黑', '正文来源徽标错误：' + extracted.sources['--body-color']);
    assert(extracted.sources['--title-color'].includes('提取自 Title'), 'Title 来源徽标错误：' + extracted.sources['--title-color']);
    assert(extracted.sources['--h2-color'] === 'OOXML 默认黑', 'Heading 2 默认黑来源错误：' + extracted.sources['--h2-color']);
    assert(extracted.sources['--table-head'].includes('表格首行底纹'), '表头来源徽标错误：' + extracted.sources['--table-head']);
    assert(extracted.profile?.summary.includes('主题色换算'), '提取报告没有记录主题色换算。');
    assert(extracted.sourceBadgeCount >= 24, '样式字段来源徽标数量不足：' + extracted.sourceBadgeCount);
    assert(extracted.fontSelects === 1 && extracted.sizeSelects >= 6, '字体/字号下拉数量不足：' + JSON.stringify(extracted));
    assert(extracted.customInputs >= 23 && extracted.colorPickers >= 8, '自定义输入或颜色取色器数量不足：' + JSON.stringify(extracted));
    assert(extracted.clearButton && extracted.reportText.includes('样式字段来源') && extracted.reportText.includes('OOXML 默认黑'), '表头清除或逐字段报告缺失。');
    await page.locator('#styleDrawer').screenshot({ path: v2StyleScreenshotPath });
    await page.locator('[data-style-clear="--table-head"]').click();
    const cleared = await page.evaluate(() => ({
      value: getComputedStyle(document.documentElement).getPropertyValue('--table-head').trim().toLowerCase(),
      source: document.querySelector('[data-style-source="--table-head"]')?.textContent.trim(),
    }));
    assert(cleared.value === '#d9e2f3' && cleared.source === '沿用默认', '表头清除没有恢复默认值：' + JSON.stringify(cleared));
    await page.locator('#styleDrawer [data-close="styleDrawer"]').click();
    return 'body=' + extracted.body + ' · Title=' + extracted.title + ' · H1=' + extracted.h1 + ' · H2=' + extracted.h2 + ' · table=' + extracted.table + ' · badges=' + extracted.sourceBadgeCount + ' · controls=' + extracted.fontSelects + '/' + extracted.sizeSelects + '/' + extracted.colorPickers + ' · clear=1';
  });

  await check('V2 中文强调着重号三端与 DOCX OOXML', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = '# V2 强调号\n\n*中文强调* 与 ***粗强调***。';
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(source);
    await page.waitForTimeout(350);
    await page.locator('#previewTab').click();
    await page.waitForTimeout(150);
    const visual = await page.evaluate(() => {
      const read = (node) => {
        if (!node) return null;
        const style = getComputedStyle(node);
        return {
          fontStyle: style.fontStyle,
          emphasis: String(style.textEmphasis || style.webkitTextEmphasis || '').toLowerCase(),
          position: String(style.textEmphasisPosition || style.webkitTextEmphasisPosition || '').toLowerCase(),
        };
      };
      return {
        left: read(document.querySelector('#markdownPreview em')),
        right: read(document.querySelector('#wordPaper em')),
        leftCount: document.querySelectorAll('#markdownPreview em').length,
        rightCount: document.querySelectorAll('#wordPaper em').length,
      };
    });
    assert(visual.leftCount >= 2 && visual.rightCount >= 2, '预览强调节点数量不足：' + JSON.stringify(visual));
    assert(visual.left?.fontStyle === 'normal' && visual.right?.fontStyle === 'normal', '强调号仍被渲染为斜体：' + JSON.stringify(visual));
    assert((visual.left?.emphasis || '').includes('dot') && (visual.right?.emphasis || '').includes('dot'), '浏览器未应用 filled dot 着重号：' + JSON.stringify(visual));
    assert((visual.left?.position || '').includes('under') && (visual.right?.position || '').includes('under'), '着重号位置不是 under：' + JSON.stringify(visual));
    const output = await waitForDownload(page, '#exportWordBtn', exportedPaths.v2EmphasisDocx);
    const summary = zipSummary(await readFile(output));
    const paragraph = xmlParagraphs(summary.documentXml).find((item) => item.includes('中文强调'));
    assert(paragraph && /w:em w:val="dot"/.test(paragraph), '导出 DOCX 缺少 w:em val="dot"：' + (paragraph || ''));
    assert(!/<w:i(?:\s|>)/.test(paragraph) && !/<w:iCs(?:\s|>)/.test(paragraph), '强调号段落残留 italics OOXML。');
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(original);
    await page.waitForTimeout(350);
    return 'Markdown em=' + visual.leftCount + ' · Word em=' + visual.rightCount + ' · font-style=normal · emphasis=filled dot/under · DOCX w:em=1 · italics=0';
  });

  await check('F3 封面预览与仅开头识别', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = await readFile(fixturePaths.cover, 'utf8');
    try {
      await page.setViewportSize({ width: 1440, height: 1100 });
      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill(source);
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const metrics = await page.evaluate((value) => {
        const read = (node) => {
          if (!node) return null;
          const style = getComputedStyle(node);
          return {
            textAlign: style.textAlign,
            borderLeftStyle: style.borderLeftStyle,
            backgroundColor: style.backgroundColor,
          };
        };
        const outside = (selector, root) => [...root.querySelectorAll(selector)].filter((node) => !node.closest('.md-cover')).length;
        const left = document.querySelector('#markdownPreview');
        const right = document.querySelector('#wordPaper');
        return {
          analysis: window.__WB_F3__?.analyze(value),
          leftCover: left?.querySelectorAll(':scope > .md-cover').length || 0,
          rightCover: right?.querySelectorAll(':scope > .md-cover').length || 0,
          leftBreak: left?.querySelectorAll(':scope > .md-cover-pagebreak[data-auto-cover-break="true"]').length || 0,
          rightBreak: right?.querySelectorAll(':scope > .md-cover-pagebreak[data-auto-cover-break="true"]').length || 0,
          leftTitle: read(left?.querySelector('.md-cover .md-title')),
          rightTitle: read(right?.querySelector('.md-cover .md-title')),
          leftDesc: read(left?.querySelector('.md-cover .md-desc')),
          rightDesc: read(right?.querySelector('.md-cover .md-desc')),
          middleTitles: outside('.md-title', left),
          source: document.querySelector('#mdEditor')?.value,
        };
      }, source);
      assert(metrics.analysis?.hasCover && metrics.analysis?.hasDesc, '封面夹具未被识别为带说明的封面：' + JSON.stringify(metrics));
      assert(metrics.leftCover === 1 && metrics.rightCover === 1, '左右预览封面节点数量错误：' + JSON.stringify(metrics));
      assert(metrics.leftBreak === 1 && metrics.rightBreak === 1, '封面后自动分页缺失：' + JSON.stringify(metrics));
      assert(metrics.leftTitle?.textAlign === 'center' && metrics.rightTitle?.textAlign === 'center', '封面标题未居中：' + JSON.stringify(metrics));
      assert(metrics.leftDesc?.textAlign === 'center' && metrics.rightDesc?.textAlign === 'center' && metrics.leftDesc?.borderLeftStyle === 'none' && metrics.rightDesc?.borderLeftStyle === 'none', '封面说明仍保留普通 desc 样式：' + JSON.stringify(metrics));
      assert(metrics.middleTitles === 1 && metrics.source === source, '文档中部 title 或源 Markdown 被错误处理：' + JSON.stringify(metrics));
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.screenshot({ path: f3CoverDesktopScreenshotPath, fullPage: false });

      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill('::title 仅标题封面\n\n# 正文');
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const titleOnly = await page.evaluate(() => ({
        cover: document.querySelectorAll('#markdownPreview > .md-cover').length,
        desc: document.querySelectorAll('#markdownPreview > .md-cover .md-desc').length,
        break: document.querySelectorAll('#markdownPreview > .md-cover-pagebreak[data-auto-cover-break="true"]').length,
      }));
      assert(titleOnly.cover === 1 && titleOnly.desc === 0 && titleOnly.break === 1, '仅 title 未形成无说明封面：' + JSON.stringify(titleOnly));

      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill('# 一级标题\n\n::title 中部标题\n\n正文');
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const middle = await page.evaluate(() => ({
        cover: document.querySelectorAll('#markdownPreview > .md-cover').length,
        title: [...document.querySelectorAll('#markdownPreview > .md-title')].length,
        break: document.querySelectorAll('#markdownPreview > .md-cover-pagebreak[data-auto-cover-break="true"]').length,
      }));
      assert(middle.cover === 0 && middle.title === 1 && middle.break === 0, '文档中部 title 错误触发封面：' + JSON.stringify(middle));

      await page.setViewportSize({ width: 390, height: 844 });
      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill(source);
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.screenshot({ path: f3CoverMobileScreenshotPath, fullPage: false });
      return '左/右封面=1/1 · 自动分页=1/1 · 标题居中 · desc 去边框底纹 · 仅开头识别 · 仅 title 兼容 · 移动端截图已落盘';
    } finally {
      await page.setViewportSize({ width: 1440, height: 1100 }).catch(() => {});
      await page.locator('#sourceTab').click().catch(() => {});
      await page.locator('#mdEditor').fill(original).catch(() => {});
      await page.waitForTimeout(350).catch(() => {});
    }
  });

  await check('F3 封面 Word 导出与首行缩进豁免', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = await readFile(fixturePaths.cover, 'utf8');
    try {
      await page.setViewportSize({ width: 1440, height: 1100 });
      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill(source);
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const output = await waitForDownload(page, '#exportWordBtn', exportedPaths.f3CoverDocx);
      const summary = zipSummary(await readFile(output));
      const paragraphs = xmlParagraphs(summary.documentXml);
      const titleIndex = paragraphs.findIndex((paragraph) => paragraph.includes('F3 封面页样本'));
      const descIndex = paragraphs.findIndex((paragraph) => paragraph.includes('这是封面说明'));
      const breakIndex = paragraphs.findIndex((paragraph) => /<w:br\b[^>]*(?:w:type|type)="page"/.test(paragraph));
      const bodyIndex = paragraphs.findIndex((paragraph) => paragraph.includes('正文首段用于确认'));
      assert(titleIndex >= 0 && descIndex > titleIndex, 'DOCX 封面标题或说明段落缺失。');
      assert(breakIndex > descIndex, 'DOCX 封面后没有强制分页。');
      assert(/<w:jc\b[^>]*w:val="center"/.test(paragraphs[titleIndex]) && /<w:jc\b[^>]*w:val="center"/.test(paragraphs[descIndex]), 'DOCX 封面段落没有水平居中。');
      assert(/<w:spacing\b[^>]*w:before="1800"/.test(paragraphs[titleIndex]), 'DOCX 封面没有写入上部三分之一位置的 spacing before。');
      assert(!paragraphs.slice(titleIndex, breakIndex + 1).some((paragraph) => /w:firstLine(?:Chars)?=/.test(paragraph)), 'DOCX 封面区错误携带正文首行缩进。');
      assert(bodyIndex > breakIndex && /w:firstLineChars="200"/.test(paragraphs[bodyIndex]), '封面后的正文没有恢复 2 字符首行缩进。');
      const reportText = await textContent('#reportBody', page);
      assert(reportText.includes('封面页格式已应用'), '封面导出报告缺少应用记录。');
      return `title/desc=center · before=1800 · pageBreak=1 · cover firstLine=0 · body firstLineChars=200 · DOCX=${formatBytes((await stat(output)).size)}`;
    } finally {
      await page.locator('#sourceTab').click().catch(() => {});
      await page.locator('#mdEditor').fill(original).catch(() => {});
      await page.waitForTimeout(350).catch(() => {});
    }
  });

  await check('V4 表格列宽预览、DOCX 导出与缩进', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = await readFile(fixturePaths.v4Table, 'utf8');
    try {
      await page.setViewportSize({ width: 1440, height: 1100 });
      await page.locator('[data-mode="word"]').click();
      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill(source);
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const preview = await page.evaluate(() => {
        const read = (selector) => {
          const table = document.querySelector(selector);
          if (!table) return null;
          const colgroup = [...table.children].find((node) => node.tagName.toLowerCase() === 'colgroup');
          const firstRowCells = [...table.querySelectorAll('tbody tr:first-child td')];
          const range = firstRowCells[0] ? document.createRange() : null;
          if (range) range.selectNodeContents(firstRowCells[0]);
          return {
            layout: table.dataset.wbTableLayout || '',
            widths: (table.dataset.wbTableColumnWidths || '').split(',').filter(Boolean).map(Number),
            colCount: colgroup ? colgroup.children.length : 0,
            cellWidths: firstRowCells.map((cell) => Math.round(cell.getBoundingClientRect().width)),
            longTextLines: range ? range.getClientRects().length : 0,
          };
        };
        return { markdown: read('#markdownPreview table'), word: read('#wordPaper table') };
      });
      assert(preview.markdown?.layout === 'fixed' && preview.word?.layout === 'fixed', `预览表格未进入 fixed 布局：${JSON.stringify(preview)}`);
      assert(preview.markdown.colCount === 3 && preview.word.colCount === 3, `预览 colgroup 列数错误：${JSON.stringify(preview)}`);
      assert(preview.markdown.widths.length === 3 && preview.word.widths.length === 3, `预览列宽数据缺失：${JSON.stringify(preview)}`);
      assert(JSON.stringify(preview.markdown.widths) === JSON.stringify(preview.word.widths), `左右预览列宽不一致：${JSON.stringify(preview)}`);
      assert(preview.markdown.cellWidths[0] > preview.markdown.cellWidths[1] && preview.markdown.cellWidths[0] >= 80, `长中文列没有获得有效宽度：${JSON.stringify(preview)}`);
      assert(preview.markdown.longTextLines <= 3 && preview.word.longTextLines <= 3, `长中文列仍接近逐字竖排：${JSON.stringify(preview)}`);
      await page.waitForTimeout(2300);
      await page.locator('#markdownPreview').screenshot({ path: v4TableDesktopScreenshotPath });
      screenshotCreated.add(v4TableDesktopScreenshotPath);
      await page.setViewportSize({ width: 800, height: 900 });
      await page.locator('#wordPaper').screenshot({ path: v4TableMobileScreenshotPath });
      screenshotCreated.add(v4TableMobileScreenshotPath);
      await page.setViewportSize({ width: 1440, height: 1100 });

      const output = await waitForDownload(page, '#exportWordBtn', exportedPaths.v4TableDocx);
      const summary = zipSummary(await readFile(output));
      const tableXml = (summary.documentXml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>/) || [])[0] || '';
      const firstRowXml = (tableXml.match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [])[0] || '';
      const gridWidths = [...tableXml.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"/g)].map((match) => Number(match[1]));
      const firstRowWidths = [...firstRowXml.matchAll(/<w:tcW\b[^>]*w:w="(\d+)"/g)].map((match) => Number(match[1]));
      const bodyParagraph = xmlParagraphs(summary.documentXml).find((paragraph) => paragraph.includes('正文首段用于确认'));
      const listParagraph = xmlParagraphs(summary.documentXml).find((paragraph) => paragraph.includes('一级列表不应继承'));
      const numbering = summary.numberingXml;
      const docxRatios = gridWidths.map((value) => value / gridWidths.reduce((sum, item) => sum + item, 0) * 100);
      const ratioDiff = Math.max(...docxRatios.map((value, index) => Math.abs(value - preview.markdown.widths[index])));
      assert(/<w:tblLayout\b[^>]*w:type="fixed"/.test(tableXml), 'DOCX 表格没有 fixed 布局。');
      assert(/<w:tblW\b[^>]*w:type="pct"[^>]*w:w="100%"/.test(tableXml), 'DOCX 表格没有 100% 总宽。');
      assert(gridWidths.length === 3 && gridWidths.every((value) => value > 0), `DOCX tblGrid 列宽错误：${gridWidths.join('/')}`);
      assert(firstRowWidths.length === 3 && firstRowWidths.every((value) => value > 0), `DOCX 首行 tcW 缺失：${firstRowWidths.join('/')}`);
      assert(ratioDiff < 0.2, `预览与 DOCX 列宽比例偏差过大：${ratioDiff}`);
      assert(bodyParagraph && bodyParagraph.includes('w:firstLineChars="200"') && !bodyParagraph.includes('w:firstLine="'), 'DOCX 正文首行缩进仍存在 firstLine 双写。');
      assert(listParagraph && !listParagraph.includes('w:firstLine'), 'DOCX 列表错误携带正文首行缩进。');
      assert(/<w:lvl\b[^>]*w:ilvl="0"[\s\S]*?<w:ind\b[^>]*w:left="420"[^>]*w:hanging="420"/.test(numbering), '编号 level0 缩进未校准为 left=420/hanging=420。');
      const reportText = await textContent('#reportBody', page);
      assert(reportText.includes('表格列宽已应用'), `转换报告缺少表格列宽记录：${reportText}`);
      return `preview=${preview.markdown.widths.join('/')}% · gridCol=${gridWidths.join('/')} twips · tcW=${firstRowWidths.join('/')} · fixed=1 · firstLine=0 · list=420/420 · DOCX=${formatBytes((await stat(output)).size)}`;
    } finally {
      await page.setViewportSize({ width: 1440, height: 1100 }).catch(() => {});
      await page.locator('#sourceTab').click().catch(() => {});
      await page.locator('#mdEditor').fill(original).catch(() => {});
      await page.waitForTimeout(350).catch(() => {});
    }
  });

  await check('V4.1 九列表头不应逐字竖排', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const source = await readFile(fixturePaths.v4WideTable, 'utf8');
    try {
      await page.setViewportSize({ width: 1440, height: 1100 });
      await page.locator('[data-mode="word"]').click();
      await page.locator('#sourceTab').click();
      await page.locator('#mdEditor').fill(source);
      await page.waitForTimeout(350);
      await page.locator('#previewTab').click();
      await page.waitForTimeout(150);
      const preview = await page.evaluate(() => {
        const read = (selector) => {
          const table = document.querySelector(selector);
          if (!table) return null;
          const headers = [...table.querySelectorAll('thead th')];
          return {
            layout: table.dataset.wbTableLayout || '',
            colCount: [...table.querySelectorAll(':scope > colgroup > col')].length,
            widths: headers.map((cell) => Math.round(cell.getBoundingClientRect().width)),
            fontSize: getComputedStyle(table).fontSize,
            padding: getComputedStyle(headers[0] || table).paddingLeft,
            headerLines: headers.map((cell) => {
              const range = document.createRange();
              range.selectNodeContents(cell);
              return range.getClientRects().length;
            }),
          };
        };
        return { markdown: read('#markdownPreview table'), word: read('#wordPaper table') };
      });
      await page.locator('#wordPaper table').screenshot({ path: v4WideTableScreenshotPath });
      screenshotCreated.add(v4WideTableScreenshotPath);
      const output = await waitForDownload(page, '#exportWordBtn', exportedPaths.v4WideTableDocx);
      const summary = zipSummary(await readFile(output));
      const tableXml = (summary.documentXml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>/) || [])[0] || '';
      const gridWidths = [...tableXml.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"/g)].map((match) => Number(match[1]));
      assert(preview.markdown?.layout === 'fixed' && preview.word?.layout === 'fixed', `九列表格未进入 fixed 布局：${JSON.stringify(preview)}`);
      assert(preview.markdown.colCount === 9 && preview.word.colCount === 9, `九列表格 colgroup 列数错误：${JSON.stringify(preview)}`);
      assert(preview.markdown.headerLines.slice(2, 8).every((lines) => lines <= 1) && preview.word.headerLines.slice(2, 8).every((lines) => lines <= 1), `五字表头仍然多行折叠：${JSON.stringify(preview)}`);
      assert(preview.markdown.fontSize === '10px' && preview.word.fontSize === '10px' && preview.markdown.padding === '4px' && preview.word.padding === '4px', `宽表紧凑排版未生效：${JSON.stringify(preview)}`);
      assert(gridWidths.length === 9 && Math.min(...gridWidths.slice(2, 8)) >= 900, `DOCX 中间语义列宽度不足：${gridWidths.join('/')}`);
      assert(/<w:tblCellMar\b[\s\S]*?<w:left\b[^>]*w:w="60"[\s\S]*?<w:right\b[^>]*w:w="60"/.test(tableXml), `DOCX 宽表格没有显式紧凑单元格边距：${tableXml.slice(0, 500)}`);
      assert(!/<w:tcPr>[\s\S]*?<w:textDirection\b/.test(tableXml), 'DOCX 宽表格错误写入文字方向。');
      return `9 列 · 紧凑字号=${preview.markdown.fontSize} · 表头行数=${preview.markdown.headerLines.slice(2, 8).join('/')} · cellMar=60twips · DOCX=${formatBytes((await stat(output)).size)}`;
    } finally {
      await page.setViewportSize({ width: 1440, height: 1100 }).catch(() => {});
      await page.locator('#sourceTab').click().catch(() => {});
      await page.locator('#mdEditor').fill(original).catch(() => {});
      await page.waitForTimeout(350).catch(() => {});
    }
  });

  await check('XLSX 导入与结构展示', async () => {
    await page.locator('#excelInput').setInputFiles(fixturePaths.xlsx);
    await page.waitForFunction(() => document.querySelectorAll('#sheetTabs .sheet-tab').length === 3, null, { timeout: 20_000 });
    const summary = await textContent('#excelSummaryText', page);
    const sheetCount = await count('#sheetTabs .sheet-tab', page);
    const gridCellCount = await count('#gridWrap .grid-cell', page);
    assert(sheetCount === 3, `Sheet 数量不正确：${sheetCount}。`);
    assert(gridCellCount > 0, 'Excel 网格没有可编辑单元格。');
    assert(/\d+ 个样式单元格/.test(summary), `Excel 摘要缺少样式统计：${summary}`);
    assert(/[1-9]\d* 个样式单元格/.test(summary), `Excel 样式未被识别：${summary}`);
    assert(/[1-9]\d* 个合并区域/.test(summary), `Excel 合并区域未被识别：${summary}`);
    assert(summary.includes('已保留列宽/行高'), `Excel 尺寸元数据未保留：${summary}`);
    await page.locator('#excelStage').screenshot({ path: screenshotPaths.excel });
    screenshotCreated.add(screenshotPaths.excel);
    return `${summary} · 网格单元格=${gridCellCount}`;
  });

  await check('XLSX 导出包可被 ExcelJS 回读', async () => {
    await waitForDownload(page, '#exportExcelBtn', exportedPaths.xlsx);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(exportedPaths.xlsx);
    const names = workbook.worksheets.map((worksheet) => worksheet.name);
    assert(names.length === 3, `导出 XLSX Sheet 数量错误：${names.join(', ')}`);
    assert(names.join('|') === '销售|库存|备注', `导出 XLSX Sheet 名称错误：${names.join(', ')}`);
    const sales = workbook.getWorksheet('销售');
    assert(sales.getCell('A1').value === '销售测试工作簿', '导出 XLSX 合并标题丢失。');
    assert(sales.model.merges.includes('A1:E1'), `导出 XLSX 合并区域丢失：${sales.model.merges.join(', ')}`);
    assert(sales.getColumn(1).width > 0 && sales.getRow(5).height > 0, '导出 XLSX 列宽/行高丢失。');
    assert(String(sales.getCell('E5').value).includes('\n'), '导出 XLSX 换行单元格丢失。');
    assert(String(sales.getCell('B3').numFmt || '').includes('¥'), '导出 XLSX 金额格式丢失。');
    assert(String(sales.getCell('C3').numFmt || '').includes('%'), '导出 XLSX 百分比格式丢失。');
    return `文件=${formatBytes((await stat(exportedPaths.xlsx)).size)} · sheets=${names.join('/')}`;
  });

  await check('XLSX 导出结果回导', async () => {
    await page.locator('#excelInput').setInputFiles(exportedPaths.xlsx);
    await page.waitForFunction(() => document.querySelectorAll('#sheetTabs .sheet-tab').length === 3, null, { timeout: 20_000 });
    const summary = await textContent('#excelSummaryText', page);
    assert(await count('#gridWrap .grid-cell', page) > 0, '回导 XLSX 网格为空。');
    assert(summary.includes('销售'), `回导 XLSX 未激活销售 Sheet：${summary}`);
    assert(/个样式单元格/.test(summary), `回导 XLSX 摘要缺少样式统计：${summary}`);
    return summary;
  });

  await check('DOCX 导入与结构展示', async () => {
    await page.locator('#wordInput').setInputFiles(fixturePaths.docx);
    await page.waitForFunction(() => {
      const stage = document.querySelector('#wordStage');
      return stage && !stage.hidden && document.querySelector('#wordPaper').textContent.trim().length > 0;
    }, null, { timeout: 20_000 });
    const headingCounts = {};
    for (let level = 1; level <= 6; level += 1) headingCounts[`h${level}`] = await count(`#wordPaper h${level}`, page);
    const tableCount = await count('#wordPaper table', page);
    const imageCount = await count('#wordPaper img', page);
    const linkCount = await count('#wordPaper a[href^="http"]', page);
    const footnoteCount = await count('#wordPaper .footnote-ref', page);
    const titleCount = await count('#wordPaper .md-title', page);
    assert(titleCount >= 1, 'DOCX 导入缺少 Title 样式内容。');
    assert(headingCounts.h2 >= 1 && headingCounts.h3 >= 1 && headingCounts.h4 >= 1 && headingCounts.h5 >= 1 && headingCounts.h6 >= 1, `DOCX 标题层级导入不完整：${JSON.stringify(headingCounts)}`);
    assert(Object.values(headingCounts).reduce((total, value) => total + value, 0) >= 6, `DOCX 导入标题数量不足：${JSON.stringify(headingCounts)}`);
    assert(tableCount >= 1, 'DOCX 导入缺少表格。');
    assert(imageCount >= 1, 'DOCX 导入缺少图片资源。');
    assert(linkCount >= 1, 'DOCX 导入缺少外部链接。');
    assert(footnoteCount >= 1, 'DOCX 导入缺少脚注引用。');
    return `title=${titleCount} · H1-H6=${Object.values(headingCounts).join('/')} · table=${tableCount} · image=${imageCount} · link=${linkCount} · footnote=${footnoteCount}`;
  });

  await check('DOCX 资源可解析为语义图片嵌入', async () => {
    const original = await page.locator('#mdEditor').inputValue();
    const match = original.match(/!\[([^\]]*)\]\((images\/[^)\s]+)\)/);
    assert(match, 'DOCX 导入结果缺少可替换的图片 Markdown。');
    const semanticSource = original.replace(match[0], `![[${match[2]}]]`);
    await page.locator('#sourceTab').click();
    await page.locator('#mdEditor').fill(semanticSource);
    await page.waitForTimeout(350);
    const embedImageCount = await count('#wordPaper .md-embed-image', page);
    const embeddedImgCount = await count('#wordPaper img[data-embedded="true"]', page);
    assert(embedImageCount === 1 && embeddedImgCount === 1, `本地图片嵌入未解析：container=${embedImageCount} · img=${embeddedImgCount}`);
    await page.locator('#mdEditor').fill(original);
    await page.waitForTimeout(350);
    return `资源池图片已识别为 md-embed-image · img=${embeddedImgCount}`;
  });

  await check('DOCX 导出包结构完整', async () => {
    await waitForDownload(page, '#exportWordBtn', exportedPaths.docx);
    const buffer = await readFile(exportedPaths.docx);
    const summary = zipSummary(buffer);
    assert(summary.zip.has('word/document.xml'), '导出 DOCX 缺少 word/document.xml。');
    assert(summary.zip.has('word/numbering.xml'), '导出 DOCX 缺少 word/numbering.xml。');
    assert(summary.zip.has('word/footnotes.xml'), '导出 DOCX 缺少 word/footnotes.xml。');
    assert(summary.names.some((name) => name.startsWith('word/media/')), '导出 DOCX 缺少图片资源。');
    assert(/w:hyperlink|r:id/.test(summary.documentXml), '导出 DOCX 外部链接关系未写入。');
    return `文件=${formatBytes(buffer.length)} · entries=${summary.names.length} · media=${summary.names.filter((name) => name.startsWith('word/media/')).length}`;
  });

  await check('DOCX 导出结果回导', async () => {
    await page.locator('#wordInput').setInputFiles(exportedPaths.docx);
    await page.waitForFunction(() => {
      const stage = document.querySelector('#wordStage');
      return stage && !stage.hidden && document.querySelector('#wordPaper').textContent.trim().length > 0;
    }, null, { timeout: 20_000 });
    const headingCount = await count('#wordPaper h1, #wordPaper h2, #wordPaper h3, #wordPaper h4, #wordPaper h5, #wordPaper h6', page);
    const tableCount = await count('#wordPaper table', page);
    const imageCount = await count('#wordPaper img', page);
    assert(headingCount >= 6, `回导 DOCX 标题不足：${headingCount}。`);
    assert(tableCount >= 1, '回导 DOCX 表格丢失。');
    assert(imageCount >= 1, '回导 DOCX 图片丢失。');
    return `headings=${headingCount} · tables=${tableCount} · images=${imageCount}`;
  });

  await check('无效 DOCX 错误恢复', async () => {
    await writeFile(invalidDocxPath, Buffer.from('not a valid docx package', 'utf8'));
    await page.locator('#wordInput').setInputFiles(invalidDocxPath);
    await page.waitForFunction(() => document.querySelector('#toast')?.textContent.includes('Word 导入失败'), null, { timeout: 20_000 });
    const reportText = await textContent('#reportBody', page);
    assert(reportText.includes('Word 解析失败'), `错误报告未出现：${reportText}`);
    await page.locator('#reportDrawer [data-close="reportDrawer"]').click();
    return '无效包被拦截，页面保留并展示转换报告。';
  });

  await check('批量转换 MD：P0 确认、ZIP 与编辑区零污染', async () => {
    const before = await page.locator('#mdEditor').inputValue();
    await page.locator('#batchBtn').click();
    await page.locator('#batchInput').setInputFiles([fixturePaths.md, fixturePaths.docx, fixturePaths.xlsx]);
    assert(await textContent('#batchFileCount', page) === '3 个文件', '批量文件清单数量错误。');
    assert(await count('#batchFileList .batch-file', page) === 3, '批量文件清单未列出三件夹具。');
    await page.locator('#batchRunBtn').click();
    await page.waitForFunction(() => {
      const node = document.querySelector('#batchConfirm');
      return node && !node.hidden;
    }, null, { timeout: 60_000 });
    const afterRun = await page.locator('#mdEditor').inputValue();
    assert(afterRun === before, '批量转换污染了编辑区 Markdown。');
    const downloadPromise = page.waitForEvent('download', { timeout: 30_000 });
    await page.locator('#batchConfirmBtn').click();
    const download = await downloadPromise;
    await download.saveAs(exportedPaths.batchMdZip);
    const summary = zipSummary(await readFile(exportedPaths.batchMdZip));
    const mdNames = summary.names.filter((name) => name.endsWith('.md'));
    assert(mdNames.length === 3, `批量 MD ZIP 文件数量错误：${mdNames.join(', ')}`);
    assert(summary.zip.has('批量转换报告.txt'), '批量 MD ZIP 缺少批量转换报告。');
    const source = await readFile(fixturePaths.md, 'utf8');
    assert(summary.zip.readText('sample.md') === source, 'Markdown 输入批量转换结果与源文件不一致。');
    assert(await page.locator('#mdEditor').inputValue() === before, '批量导出后编辑区 Markdown 发生变化。');
    return `files=${mdNames.length} · P0 确认通过 · 编辑区 MD diff=0 · report=1`;
  });

  await check('批量转换 Word：ZIP 产物可读取', async () => {
    await page.locator('#batchClearBtn').click();
    await page.locator('#batchInput').setInputFiles([fixturePaths.docx, fixturePaths.xlsx]);
    await page.locator('#batchTarget').selectOption('word');
    const before = await page.locator('#mdEditor').inputValue();
    const downloadPromise = page.waitForEvent('download', { timeout: 60_000 });
    await page.locator('#batchRunBtn').click();
    const download = await downloadPromise;
    await download.saveAs(exportedPaths.batchWordZip);
    const summary = zipSummary(await readFile(exportedPaths.batchWordZip));
    const docxNames = summary.names.filter((name) => name.endsWith('.docx'));
    assert(docxNames.length === 2, `批量 Word ZIP 文件数量错误：${docxNames.join(', ')}`);
    assert(docxNames.every((name) => summary.zip.readText(name).includes('document')), '批量 Word 产物缺少可读取的文档 XML。');
    assert(summary.zip.has('批量转换报告.txt'), '批量 Word ZIP 缺少批量转换报告。');
    assert(await page.locator('#mdEditor').inputValue() === before, '批量 Word 转换污染了编辑区 Markdown。');
    return `files=${docxNames.length} · document.xml 可读取 · 编辑区 MD diff=0`;
  });

  await check('批量转换 Excel：ZIP 产物可回读', async () => {
    await page.locator('#batchClearBtn').click();
    await page.locator('#batchInput').setInputFiles([fixturePaths.docx, fixturePaths.xlsx]);
    await page.locator('#batchTarget').selectOption('excel');
    const before = await page.locator('#mdEditor').inputValue();
    const downloadPromise = page.waitForEvent('download', { timeout: 60_000 });
    await page.locator('#batchRunBtn').click();
    const download = await downloadPromise;
    await download.saveAs(exportedPaths.batchExcelZip);
    const summary = zipSummary(await readFile(exportedPaths.batchExcelZip));
    const xlsxNames = summary.names.filter((name) => name.endsWith('.xlsx'));
    assert(xlsxNames.length === 2, `批量 Excel ZIP 文件数量错误：${xlsxNames.join(', ')}`);
    for (const name of xlsxNames) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(summary.zip.read(name));
      assert(workbook.worksheets.length >= 1, `批量 Excel 产物没有工作表：${name}`);
    }
    assert(summary.zip.has('批量转换报告.txt'), '批量 Excel ZIP 缺少批量转换报告。');
    assert(await page.locator('#mdEditor').inputValue() === before, '批量 Excel 转换污染了编辑区 Markdown。');
    return `files=${xlsxNames.length} · ExcelJS 回读通过 · 编辑区 MD diff=0`;
  });

  await check('批量输入 CSV：经 Excel 中枢转 MD', async () => {
    const csvPath = path.join(artifactDir, 'batch-input.csv');
    await writeFile(csvPath, '\uFEFF名称,金额\n华东,1280000\n华南,960000\n', 'utf8');
    await page.locator('#batchClearBtn').click();
    await page.locator('#batchInput').setInputFiles(csvPath);
    await page.locator('#batchTarget').selectOption('md');
    const before = await page.locator('#mdEditor').inputValue();
    const downloadPromise = page.waitForEvent('download', { timeout: 60_000 });
    await page.locator('#batchRunBtn').click();
    const download = await downloadPromise;
    await download.saveAs(exportedPaths.batchCsvZip);
    const summary = zipSummary(await readFile(exportedPaths.batchCsvZip));
    const mdNames = summary.names.filter((name) => name.endsWith('.md'));
    assert(mdNames.length === 1, `CSV 批量 MD ZIP 文件数量错误：${mdNames.join(', ')}`);
    assert(summary.zip.readText(mdNames[0]).includes('华东') && summary.zip.readText(mdNames[0]).includes('1280000'), 'CSV 内容未进入 Markdown 结果。');
    assert(summary.zip.has('批量转换报告.txt'), 'CSV 批量 ZIP 缺少批量转换报告。');
    assert(await page.locator('#mdEditor').inputValue() === before, 'CSV 批量转换污染了编辑区 Markdown。');
    return `csv=1 · Excel 中枢转 MD · 编辑区 MD diff=0`;
  });

  await check('F2 打开与只读关联', async () => {
    const closeBatch = page.locator('#batchDrawer [data-close="batchDrawer"]');
    if (await closeBatch.count() && await closeBatch.isVisible()) await closeBatch.click();
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.permissionMode = 'readwrite';
      test.content = '# Disk baseline\n\n原始内容';
      test.lastModified = 100;
      test.openCalls = 0;
    });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().supported)) === true, '测试浏览器未识别 File System Access API。');
    await page.locator('#openMdBtn').click();
    await page.waitForFunction(() => {
      const status = window.__WB_F2__?.getState();
      return status && status.associated && !status.busy;
    }, null, { timeout: 10_000 });
    const opened = await page.evaluate(() => ({
      value: document.querySelector('#mdEditor')?.value,
      association: document.querySelector('#sourceAssociation')?.textContent,
      fileName: window.__WB_F2__?.getState().fileName,
      permission: window.__WB_F2__?.getState().permission,
      openCalls: window.__WB_F2_TEST__?.openCalls,
    }));
    assert(opened.value.includes('# Disk baseline'), '打开后的 Markdown 未读取 mock 文件内容。');
    assert(opened.association === '已关联' && opened.fileName === 'linked.md', `关联状态不正确：${JSON.stringify(opened)}`);
    assert(opened.permission === 'readwrite' && opened.openCalls === 1, `读写权限或打开次数不正确：${JSON.stringify(opened)}`);
    assert(await page.locator('#refreshMdBtn').isEnabled(), '关联后刷新按钮未启用。');
    assert(await page.locator('#saveMdBtn').isEnabled(), '读写关联后保存按钮未启用。');

    await page.evaluate(() => {
      window.__WB_F2_TEST__.permissionMode = 'read';
    });
    await page.locator('#openMdBtn').click();
    await page.waitForFunction(() => {
      const status = window.__WB_F2__?.getState();
      return status && status.associated && status.permission === 'read' && !status.busy;
    }, null, { timeout: 10_000 });
    const readOnly = await page.evaluate(() => ({
      association: document.querySelector('#sourceAssociation')?.textContent,
      saveDisabled: document.querySelector('#saveMdBtn')?.disabled,
    }));
    assert(readOnly.association.includes('只读') && readOnly.saveDisabled, `只读关联状态不正确：${JSON.stringify(readOnly)}`);

    await page.evaluate(() => {
      window.__WB_F2_TEST__.permissionMode = 'readwrite';
    });
    await page.locator('#openMdBtn').click();
    await page.waitForFunction(() => {
      const status = window.__WB_F2__?.getState();
      return status && status.permission === 'readwrite' && !status.busy;
    }, null, { timeout: 10_000 });
    return `打开=1 · 文件=linked.md · 读写关联=1 · 只读降级=1`;
  });

  await check('F2 刷新与保存双向链路', async () => {
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# External refresh\n\n磁盘已更新';
      test.lastModified = 200;
    });
    await page.locator('#refreshMdBtn').click();
    await page.waitForFunction(() => document.querySelector('#mdEditor')?.value.includes('# External refresh'), null, { timeout: 10_000 });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === false, '刷新后仍被标记为未保存。');

    await page.locator('#mdEditor').fill('# Local save\n\n平台修改');
    await page.waitForTimeout(350);
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === true, '编辑 Markdown 后未进入未保存状态。');
    assert((await page.locator('#saveMdBtn').textContent()).includes('未保存'), '保存按钮未呈现未保存提示态。');
    await page.locator('#saveMdBtn').click();
    await page.waitForFunction(() => window.__WB_F2_TEST__?.content.includes('# Local save'), null, { timeout: 10_000 });
    const saved = await page.evaluate(() => ({
      content: window.__WB_F2_TEST__?.content,
      dirty: window.__WB_F2__?.getState().dirty,
      diskStamp: window.__WB_F2_TEST__?.lastModified,
    }));
    assert(saved.content.includes('平台修改') && saved.dirty === false && saved.diskStamp === 201, `保存结果不正确：${JSON.stringify(saved)}`);
    return `外部修改→刷新=1 · 本地修改→保存=1 · mock 内容已同步 · dirty=false`;
  });

  await check('F2 刷新与保存冲突保护', async () => {
    await page.locator('#mdEditor').fill('# Local save conflict');
    await page.waitForTimeout(350);
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# External before save';
      test.lastModified = 300;
    });
    await clickWithDialog(page, '#saveMdBtn', false, '磁盘上被修改');
    await page.waitForTimeout(120);
    assert((await page.evaluate(() => window.__WB_F2_TEST__?.content)) === '# External before save', '取消保存冲突后覆盖了磁盘内容。');
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === true, '取消保存冲突后未保留本地未保存状态。');

    await clickWithDialog(page, '#saveMdBtn', true, '磁盘上被修改');
    await page.waitForFunction(() => window.__WB_F2_TEST__?.content === '# Local save conflict', null, { timeout: 10_000 });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === false, '确认覆盖后未重置保存基线。');

    await page.locator('#mdEditor').fill('# Local refresh conflict');
    await page.waitForTimeout(350);
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# External before refresh';
      test.lastModified = 400;
    });
    await clickWithDialog(page, '#refreshMdBtn', false, '未保存修改');
    await page.waitForTimeout(120);
    assert((await page.locator('#mdEditor').inputValue()) === '# Local refresh conflict', '取消刷新冲突后覆盖了本地内容。');

    await clickWithDialog(page, '#refreshMdBtn', true, '未保存修改');
    await page.waitForFunction(() => document.querySelector('#mdEditor')?.value === '# External before refresh', null, { timeout: 10_000 });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === false, '确认刷新后未重置保存基线。');
    return `保存冲突取消/确认=1/1 · 刷新冲突取消/确认=1/1 · 双向均保护本地与磁盘内容`;
  });

  await check('F2 自动同步与 Ctrl+S', async () => {
    await page.locator('#autoSyncToggle').check();
    await page.waitForFunction(() => window.__WB_F2__?.getState().autoSync === true, null, { timeout: 5_000 });
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# Auto sync update';
      test.lastModified = 500;
    });
    await page.waitForFunction(() => document.querySelector('#mdEditor')?.value === '# Auto sync update', null, { timeout: 7_000 });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === false, '自动同步后未重置保存基线。');
    await page.evaluate(() => {
      window.__WB_F2_TEST__.permissionMode = 'denied';
    });
    await page.waitForFunction(() => {
      const status = window.__WB_F2__?.getState();
      const message = document.querySelector('#fileSyncStatus')?.textContent || '';
      return status && status.autoSync === false && message.includes('权限');
    }, null, { timeout: 7_000 });
    await page.evaluate(() => {
      window.__WB_F2_TEST__.permissionMode = 'readwrite';
    });
    await page.locator('#openMdBtn').click();
    await page.waitForFunction(() => window.__WB_F2__?.getState().permission === 'readwrite' && !window.__WB_F2__?.getState().busy, null, { timeout: 10_000 });
    await page.locator('#autoSyncToggle').check();
    await page.waitForFunction(() => window.__WB_F2__?.getState().autoSync === true, null, { timeout: 5_000 });

    await page.locator('#mdEditor').fill('# Local pause');
    await page.waitForTimeout(350);
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# Must not replace local';
      test.lastModified = 600;
    });
    await page.waitForTimeout(2_300);
    const paused = await page.evaluate(() => ({
      value: document.querySelector('#mdEditor')?.value,
      dirty: window.__WB_F2__?.getState().dirty,
      autoSync: window.__WB_F2__?.getState().autoSync,
    }));
    assert(paused.value === '# Local pause' && paused.dirty === true && paused.autoSync === true, `本地未保存时自动同步未暂停：${JSON.stringify(paused)}`);
    await page.locator('#autoSyncToggle').uncheck();
    await page.evaluate(() => {
      const test = window.__WB_F2_TEST__;
      test.content = '# Local pause';
      test.lastModified = 601;
    });
    await page.evaluate(() => window.__WB_F2__?.refresh({ silent: true }));
    await page.waitForFunction(() => document.querySelector('#mdEditor')?.value === '# Local pause' && window.__WB_F2__?.getState().dirty === false, null, { timeout: 10_000 });

    await page.locator('#mdEditor').fill('# Ctrl save');
    await page.waitForTimeout(350);
    await page.locator('#mdEditor').press('Control+s');
    await page.waitForFunction(() => window.__WB_F2_TEST__?.content === '# Ctrl save', null, { timeout: 10_000 });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().dirty)) === false, 'Ctrl+S 保存后未重置基线。');
    return `自动同步=2 秒轮询生效 · 权限失效自动停止=1 · 本地未保存时暂停=1 · Ctrl+S 写回=1`;
  });

  await check('F2 无 API 降级', async () => {
    await page.evaluate(() => {
      delete window.showOpenFilePicker;
      window.__WB_F2__?.refreshUi();
    });
    assert((await page.evaluate(() => window.__WB_F2__?.getState().supported)) === false, '移除 API 后仍报告为支持。');
    assert(await page.locator('#refreshMdBtn').isDisabled(), '无 API 时刷新按钮未禁用。');
    assert(await page.locator('#saveMdBtn').isDisabled(), '无 API 时保存按钮未禁用。');
    assert((await page.locator('#refreshMdBtn').getAttribute('title')).includes('不支持'), '刷新降级 tooltip 缺失。');
    await page.locator('#openMdBtn').click();
    await page.locator('#mdInput').setInputFiles(fixturePaths.md);
    await page.waitForFunction(() => document.querySelector('#sourceAssociation')?.textContent === '未关联' && !window.__WB_F2__?.getState().associated, null, { timeout: 10_000 });
    await page.waitForFunction(() => document.querySelector('#mdEditor')?.value.includes('M3.2 黄金样本文档'), null, { timeout: 10_000 });
    assert(await page.locator('#mdEditor').inputValue() === await readFile(fixturePaths.md, 'utf8'), '普通文件选择回退未载入 Markdown。');
    assert(await page.locator('#refreshMdBtn').isDisabled() && await page.locator('#saveMdBtn').isDisabled(), '普通文件回退后刷新/保存未保持禁用。');
    return '无 API → 普通文件选择 · 无句柄关联 · 刷新/保存禁用 · tooltip 可见';
  });

  await check('M1 首页选择规范与工作台主链路', async () => {
    await mkdir(m1Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    assert(await page.locator('#homeView').isVisible(), '首页未显示。');
    assert(await page.locator('#homeNextBtn').isDisabled(), '未上传文件时下一步应禁用。');
    assert(await page.locator('[data-app-view="workspace"]').count() === 0, '一级导航仍出现工作台。');
    await page.screenshot({ path: m1HomeScreenshotPath, fullPage: false });
    await page.locator('#homeInput').setInputFiles(fixturePaths.docx);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 1, null, { timeout: 8_000 });
    assert(await page.locator('#homeNextBtn').isEnabled(), '上传 Word 后下一步仍禁用。');
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceView').waitFor({ state: 'visible', timeout: 5_000 });
    assert(await page.locator('#choiceTemplateBtn').isVisible(), '缺少上传模板入口。');
    assert(await page.locator('#choiceBlankBtn').isVisible(), '缺少从空白开始入口。');
    await page.screenshot({ path: m1ChoiceScreenshotPath, fullPage: false });
    await page.locator('#choiceBlankBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    assert(await page.locator('#fileRail').isVisible(), '工作台左栏文件列表未显示。');
    assert(await page.locator('#wordPaper').isVisible(), '工作台预览未显示。');
    await page.screenshot({ path: m1WorkspaceScreenshotPath, fullPage: false });
    const task = await page.evaluate(() => window.__WB_M1__?.getTask());
    assert(task?.step === 3 && task?.fileCount >= 1, `任务状态不正确：${JSON.stringify(task)}`);
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `首页截图 · 选择规范 · 三栏工作台 · files=${task.fileCount}`;
  });

  await check('M2 任务生命周期与 M3 Specification', async () => {
    await mkdir(m2Dir, { recursive: true });
    await mkdir(m4Dir, { recursive: true });
    const secondDocx = path.join(m2Dir, 'sample-b.docx');
    await copyFile(fixturePaths.docx, secondDocx);
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.evaluate(() => localStorage.removeItem('markpivot-spec-library-v1'));
    const styleBefore = await page.evaluate(() => localStorage.getItem('document-workbench-style-v2'));
    await page.locator('#homeInput').setInputFiles([fixturePaths.docx, secondDocx]);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 2, null, { timeout: 8_000 });
    const firstTaskId = await page.evaluate(() => window.__WB_M1__?.getTask()?.id);
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceBlankBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    await page.waitForFunction(() => Array.from(document.querySelectorAll('#fileRailList .mp-file-meta')).some((node) => node.textContent.includes('已载入')), null, { timeout: 15_000 });
    await page.locator('[data-spec-tab="body"]').click();
    const model = await page.evaluate(() => {
      const spec = window.__WB_M3__?.getSpecification();
      const parsed = window.__WB_M3__?.parseSpecification(JSON.stringify({ body: { size: '18px' } }));
      return {
        roundtrip: window.__WB_M3__?.roundtripOk(),
        hasPage: !!spec?.page?.paper,
        hasHeader: spec?.header && Object.prototype.hasOwnProperty.call(spec.header, 'text'),
        hasFooter: spec?.footer && Object.prototype.hasOwnProperty.call(spec.footer, 'pageNumber'),
        missingFilled: !!(parsed?.body?.font && parsed?.body?.size === '18px' && parsed?.page?.paper),
        uiBound: !!document.querySelector('[data-spec-path="body.font"]'),
      };
    });
    assert(model.roundtrip, 'Specification JSON roundtrip 失败。');
    assert(model.hasPage && model.hasHeader && model.hasFooter && model.missingFilled && model.uiBound, `Specification 模型不完整：${JSON.stringify(model)}`);
    const mdBefore = await page.locator('#mdEditor').inputValue();
    await page.locator('#specBodySize').fill('18px');
    await page.waitForFunction(() => window.__WB_M3__?.getSpecification()?.body?.size === '18px', null, { timeout: 5_000 });
    assert(await page.locator('#mdEditor').inputValue() === mdBefore, '修改规范改写了 Markdown 语义。');
    assert(await page.evaluate(() => window.__WB_M2__?.library()?.length) === 0, '任务内修改自动写入了规范库。');
    assert(await page.evaluate(() => localStorage.getItem('document-workbench-style-v2')) === styleBefore, '任务修改污染了旧样式存储。');
    await page.locator('#specSaveLibraryBtn').click();
    await page.waitForFunction(() => window.__WB_M2__?.library()?.length >= 1, null, { timeout: 5_000 });
    await page.locator('[data-spec-scope="file"]').click();
    const fileIds = await page.evaluate(() => ({
      current: window.__WB_M2__?.getTask()?.currentFileId,
      files: Array.from(document.querySelectorAll('#fileRailList [data-file-id]')).map((node) => node.dataset.fileId),
    }));
    await page.locator('#specBodyColor').fill('#112233');
    await page.waitForFunction(() => window.__WB_M3__?.getSpecification()?.body?.color === '#112233', null, { timeout: 5_000 });
    assert((await page.evaluate(() => window.__WB_M2__?.getOverride()?.body?.color)) === '#112233', '当前文件 override 未写入。');
    const otherId = fileIds.files.find((id) => id && id !== fileIds.current);
    assert(otherId, '测试需要至少两个任务文件。');
    await page.locator(`#fileRailList [data-file-id="${otherId}"] .mp-file-name`).click();
    await page.waitForFunction(({ id, color }) => (
      window.__WB_M2__?.getTask()?.currentFileId === id
      && window.__WB_M3__?.getSpecification()?.body?.color !== color
    ), { id: otherId, color: '#112233' }, { timeout: 10_000 });
    await page.locator(`#fileRailList [data-file-id="${fileIds.current}"] .mp-file-name`).click();
    await page.waitForFunction((color) => window.__WB_M3__?.getSpecification()?.body?.color === color, '#112233', { timeout: 10_000 });
    await page.screenshot({ path: m2WorkspaceScreenshotPath, fullPage: false });
    await page.locator('[data-step="1"]').click();
    await page.locator('#homeView').waitFor({ state: 'visible', timeout: 5_000 });
    assert((await page.evaluate(() => window.__WB_M1__?.getTask()?.fileCount)) >= 2, '步骤条回到上传页时任务被清空。');
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceView').waitFor({ state: 'visible', timeout: 5_000 });
    page.once('dialog', async (dialog) => {
      assert(String(dialog.message()).includes('离开当前任务'), `离开确认文案不正确：${dialog.message()}`);
      await dialog.accept();
    });
    await page.locator('[data-app-view="home"]').click();
    await page.locator('#homeView').waitFor({ state: 'visible', timeout: 5_000 });
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount === 0, null, { timeout: 5_000 });
    await page.locator('#homeInput').setInputFiles(fixturePaths.docx);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount === 1, null, { timeout: 8_000 });
    const secondTaskId = await page.evaluate(() => window.__WB_M1__?.getTask()?.id);
    assert(secondTaskId && secondTaskId !== firstTaskId, `新任务没有独立 id：${firstTaskId} / ${secondTaskId}`);
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `独立 Task · override 切文件 · 规范库显式保存 · JSON roundtrip · files 隔离`;
  });

  await check('M4 模板提取生成 Specification 初稿', async () => {
    await mkdir(m4Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.locator('#homeInput').setInputFiles(fixturePaths.docx);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 1, null, { timeout: 8_000 });
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceView').waitFor({ state: 'visible', timeout: 8_000 });
    await page.locator('#templateInput').setInputFiles(fixturePaths.referenceDocx);
    await page.locator('#extractSummary').waitFor({ state: 'visible', timeout: 20_000 });
    const summary = await page.evaluate(() => ({
      extracted: Array.from(document.querySelectorAll('#extractExtractedList li')).map((node) => node.textContent),
      missing: Array.from(document.querySelectorAll('#extractMissingList li')).map((node) => node.textContent),
      pendingPaper: window.__WB_M4__?.pending()?.page?.paper,
      pendingHeader: window.__WB_M4__?.pending()?.header?.text,
    }));
    assert(summary.extracted.includes('页面') && summary.extracted.includes('正文') && summary.extracted.includes('一级标题') && summary.extracted.includes('二级标题') && summary.extracted.includes('三级标题') && summary.extracted.includes('表格') && summary.extracted.includes('页眉页脚'), `已提取分组不完整：${JSON.stringify(summary.extracted)}`);
    assert(summary.missing.includes('图片') && !summary.extracted.includes('图片'), `图片约束被猜测提取：${JSON.stringify(summary)}`);
    assert(summary.pendingPaper === 'A4' && summary.pendingHeader === 'MarkPivot 参考页眉', `待确认初稿不正确：${JSON.stringify(summary)}`);
    await page.screenshot({ path: m4ExtractScreenshotPath, fullPage: false });
    await page.locator('#extractContinueBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    const model = await page.evaluate(() => {
      const spec = window.__WB_M3__?.getSpecification();
      const banner = document.querySelector('#specExtractBanner');
      return {
        paper: spec?.page?.paper,
        orientation: spec?.page?.orientation,
        marginTop: spec?.page?.marginTop,
        bodyColor: spec?.body?.color,
        h1: spec?.heading1?.color,
        tableHead: spec?.table?.head,
        tableBorder: spec?.table?.border,
        header: spec?.header?.text,
        footerPage: spec?.footer?.pageNumber,
        extracted: spec?.coverage?.extracted || [],
        missing: spec?.coverage?.missing || [],
        imageStatus: spec?.coverage?.fields?.['image.maxWidth']?.status,
        imageWidth: spec?.image?.maxWidth,
        banner: banner && !banner.hidden ? banner.textContent : '',
        imageMarked: !!document.querySelector('[data-spec-path="image.maxWidth"]')?.classList.contains('is-spec-missing'),
        roundtrip: window.__WB_M3__?.roundtripOk(),
      };
    });
    assert(model.paper === 'A4' && model.orientation === 'portrait' && model.marginTop === '96px', `页面提取不正确：${JSON.stringify(model)}`);
    assert(model.bodyColor === '#000000' && model.h1 === '#f1975a' && model.tableHead === '#698ed0', `正文/标题/表头颜色不正确：${JSON.stringify(model)}`);
    assert(model.header === 'MarkPivot 参考页眉' && model.tableBorder === '#4472c4' && model.footerPage === true, `页眉页脚或表格边框不正确：${JSON.stringify(model)}`);
    assert(model.missing.includes('图片') && model.imageStatus === 'missing' && model.imageMarked, `未提取图片未标记：${JSON.stringify(model)}`);
    assert(model.roundtrip && String(model.banner).includes('未提取'), `提取结果未驱动工作台或 roundtrip 失败：${JSON.stringify(model)}`);
    const origin = new URL(page.appUrl).origin;
    const fingerprints = await page.evaluate(async ({ templateUrl, sampleUrl }) => {
      const load = async (url, name) => {
        const buffer = await fetch(url).then((response) => response.arrayBuffer());
        return window.__WB_M4__.extractFromFile(new File([buffer], name));
      };
      const first = await load(templateUrl, 'reference-template.docx');
      const firstFp = window.__WB_M4__.fingerprint();
      const second = await load(templateUrl, 'reference-template.docx');
      const secondFp = window.__WB_M4__.fingerprint();
      const sample = await load(sampleUrl, 'sample.docx');
      return {
        same: firstFp === secondFp && !!firstFp,
        paper: first.page.paper,
        paper2: second.page.paper,
        sampleMissing: sample.coverage?.missing || [],
        sampleImage: sample.coverage?.fields?.['image.maxWidth']?.status,
        sampleWarning: (sample.coverage?.warnings || []).some((item) => String(item).includes('图片')),
      };
    }, {
      templateUrl: `${origin}/tools/fixtures/reference-template.docx`,
      sampleUrl: `${origin}/tools/fixtures/sample.docx`,
    });
    assert(fingerprints.same && fingerprints.paper === 'A4' && fingerprints.paper2 === 'A4', `同一模板重复提取不稳定：${JSON.stringify(fingerprints)}`);
    assert(fingerprints.sampleMissing.includes('图片') && fingerprints.sampleImage === 'missing' && fingerprints.sampleWarning, `含图片文档被猜测了图片约束：${JSON.stringify(fingerprints)}`);
    await page.screenshot({ path: m4WorkspaceScreenshotPath, fullPage: false });
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `A4/96px · H1=${model.h1} · table=${model.tableHead} · 页眉已提取 · 图片未猜测 · fingerprint 稳定`;
  });

  await check('M5 语义识别不静默改写低置信度标题', async () => {
    await mkdir(m5Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    const origin = new URL(page.appUrl).origin;
    const mapped = await page.evaluate(async ({ sampleUrl, manualUrl }) => {
      const load = async (url, name) => {
        const buffer = await fetch(url).then((response) => response.arrayBuffer());
        return window.__WB_M5__.mapFromFile(new File([buffer], name));
      };
      const sample = await load(sampleUrl, 'sample.docx');
      const sampleAgain = await load(sampleUrl, 'sample.docx');
      const manual = await load(manualUrl, 'manual-heading.docx');
      const heading = (sample.blocks || []).find((block) => block.text === '一级标题');
      const collapsed = (sample.blocks || []).find((block) => block.text === '四级标题');
      const chapter = (manual.blocks || []).find((block) => block.text.includes('第一章'));
      const section = (manual.blocks || []).find((block) => block.text.includes('1.1'));
      const tableCaption = (manual.blocks || []).find((block) => block.semanticType === 'tableCaption');
      const figureCaption = (manual.blocks || []).find((block) => block.semanticType === 'figureCaption');
      return {
        same: sample.fingerprint === sampleAgain.fingerprint && !!sample.fingerprint,
        headingType: heading?.semanticType,
        headingConf: heading?.confidence,
        headingReview: heading?.needsReview,
        headingSource: heading?.source,
        collapsedType: collapsed?.semanticType,
        collapsedReview: collapsed?.needsReview,
        chapterType: chapter?.semanticType,
        chapterReview: chapter?.needsReview,
        chapterSource: chapter?.source,
        chapterConf: chapter?.confidence,
        sectionType: section?.semanticType,
        sectionReview: section?.needsReview,
        tableCaptionType: tableCaption?.semanticType,
        tableCaptionReview: tableCaption?.needsReview,
        figureCaptionType: figureCaption?.semanticType,
        figureCaptionReview: figureCaption?.needsReview,
        manualReview: manual.reviewCount,
      };
    }, {
      sampleUrl: `${origin}/tools/fixtures/sample.docx`,
      manualUrl: `${origin}/tools/fixtures/manual-heading.docx`,
    });
    assert(mapped.same, `同一文件重复识别不稳定：${JSON.stringify(mapped)}`);
    assert(mapped.headingType === 'heading1' && mapped.headingConf >= 0.9 && mapped.headingReview === false && mapped.headingSource === 'style', `Heading Style 未优先识别：${JSON.stringify(mapped)}`);
    assert(mapped.collapsedType === 'heading3' && mapped.collapsedReview === false, `四级标题折叠不正确：${JSON.stringify(mapped)}`);
    assert(mapped.chapterType === 'heading1' && mapped.chapterReview === true && mapped.chapterSource === 'heuristic' && mapped.chapterConf <= 0.84, `手工一级标题未进入需确认：${JSON.stringify(mapped)}`);
    assert(mapped.sectionType === 'heading2' && mapped.sectionReview === true, `手工二级标题未进入需确认：${JSON.stringify(mapped)}`);
    assert(mapped.tableCaptionType === 'tableCaption' && mapped.tableCaptionReview === false, `表题识别不正确：${JSON.stringify(mapped)}`);
    assert(mapped.figureCaptionType === 'figureCaption' && mapped.figureCaptionReview === true, `无图图题应保持需确认：${JSON.stringify(mapped)}`);
    assert(mapped.manualReview >= 2, `手工夹具待确认数量不足：${JSON.stringify(mapped)}`);

    await page.locator('#homeInput').setInputFiles([fixturePaths.docx, fixturePaths.manualHeading]);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 2, null, { timeout: 8_000 });
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceView').waitFor({ state: 'visible', timeout: 8_000 });
    await page.locator('#choiceBlankBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    await page.waitForFunction(() => {
      const map = window.__WB_M5__?.getMap();
      return map && /sample\.docx$/i.test(map.fileName || '') && Array.isArray(map.blocks) && map.blocks.length > 0;
    }, null, { timeout: 20_000 });
    await page.locator('#fileRailList .mp-file-name', { hasText: 'manual-heading.docx' }).click();
    await page.waitForFunction(() => {
      const map = window.__WB_M5__?.getMap();
      return map && /manual-heading\.docx$/i.test(map.fileName || '') && map.reviewCount > 0;
    }, null, { timeout: 20_000 });
    const before = await page.evaluate(() => ({
      md: document.querySelector('#mdEditor')?.value || '',
      status: Array.from(document.querySelectorAll('#fileRailList .mp-file')).map((node) => node.textContent),
      review: window.__WB_M5__?.reviewCount(),
      chapter: (window.__WB_M5__?.getMap()?.blocks || []).find((block) => String(block.text).includes('第一章')),
      visible: !document.querySelector('#semanticMap')?.hidden,
    }));
    assert(!/#\s*第一章/.test(before.md), `低置信度标题被静默改写成 Markdown：${before.md.slice(0, 200)}`);
    assert(before.status.some((text) => text.includes('需确认') && text.includes('待确认')), `文件栏未标出异常文件：${JSON.stringify(before.status)}`);
    assert(before.visible && before.chapter && before.chapter.needsReview, `工作台未展示待确认语义：${JSON.stringify(before)}`);
    await page.screenshot({ path: m5ReviewScreenshotPath, fullPage: false });
    await page.evaluate((id) => window.__WB_M5__.setType(id, 'heading1', true), before.chapter.id);
    await page.waitForFunction(() => /#\s*第一章/.test(document.querySelector('#mdEditor')?.value || ''), null, { timeout: 8_000 });
    const after = await page.evaluate(() => ({
      md: document.querySelector('#mdEditor')?.value || '',
      remaining: window.__WB_M5__?.reviewCount(),
    }));
    assert(/#\s*第一章/.test(after.md), `确认后未把标题写入 Markdown：${after.md.slice(0, 200)}`);
    assert(after.remaining < before.review, `确认后待确认数量未下降：${before.review} → ${after.remaining}`);
    await page.screenshot({ path: m5WorkspaceScreenshotPath, fullPage: false });
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `style H1=${mapped.headingConf} · heuristic H1 需确认 · 确认后写入 # 第一章 · fingerprint 稳定`;
  });

  await check('M6 工作台实时预览、作用域与模板恢复', async () => {
    await mkdir(m6Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.locator('#homeInput').setInputFiles([fixturePaths.docx, fixturePaths.manualHeading]);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 2, null, { timeout: 8_000 });
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceBlankBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    await page.locator('#previewToolbar').waitFor({ state: 'visible', timeout: 8_000 });
    await page.waitForFunction(() => {
      const map = window.__WB_M5__?.getMap();
      return map && /sample\.docx$/i.test(map.fileName || '') && Array.isArray(map.blocks) && map.blocks.length > 0;
    }, null, { timeout: 20_000 });
    const mdBefore = await page.locator('#mdEditor').inputValue();
    const baseline = await page.evaluate(() => ({
      size: window.__WB_M3__?.getSpecification()?.body?.size,
      font: window.__WB_M6__?.getPreview()?.computedFont,
      paper: window.__WB_M6__?.getPreview()?.paper,
    }));
    await page.locator('[data-spec-tab="body"]').click();
    await page.locator('#specBodySize').fill('18px');
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.computedFont === '18px', null, { timeout: 5_000 });
    assert(await page.locator('#mdEditor').inputValue() === mdBefore, '修改正文设置改写了 Markdown。');
    assert((await page.evaluate(() => window.__WB_M3__?.getSpecification()?.body?.size)) === '18px', '正文设置未写入 Specification。');
    await page.locator('[data-preview-mode="source"]').click();
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.mode === 'source', null, { timeout: 5_000 });
    const sourcePreview = await page.evaluate(() => ({
      font: window.__WB_M6__?.getPreview()?.computedFont,
      specSize: window.__WB_M3__?.getSpecification()?.body?.size,
    }));
    assert(sourcePreview.font !== '18px' && sourcePreview.specSize === '18px', `原文模式仍套用了目标规范或改写了 spec：${JSON.stringify(sourcePreview)}`);
    await page.locator('[data-preview-mode="effect"]').click();
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.mode === 'effect' && window.__WB_M6__?.getPreview()?.computedFont === '18px', null, { timeout: 5_000 });
    await page.locator('#previewZoomIn').click();
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.zoom === 125, null, { timeout: 5_000 });
    assert((await page.locator('#previewZoomLabel').textContent()) === '125%', '缩放标签未更新。');
    await page.locator('[data-spec-tab="page"]').click();
    await page.locator('#specPaper').selectOption('A3');
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.width === 1123, null, { timeout: 5_000 });
    await page.locator('#specRestoreBtn').click();
    await page.waitForFunction((size) => window.__WB_M3__?.getSpecification()?.body?.size === size && window.__WB_M6__?.getPreview()?.paper === 'A4', baseline.size, { timeout: 5_000 });
    assert((await page.evaluate(() => window.__WB_M6__?.getPreview()?.computedFont)) === baseline.font, '恢复模板值后预览未回到模板。');
    await page.locator('#templateInput').setInputFiles(fixturePaths.referenceDocx);
    await page.waitForFunction(() => window.__WB_M3__?.getSpecification()?.header?.text === 'MarkPivot 参考页眉', null, { timeout: 15_000 });
    const afterTemplate = await page.evaluate(() => ({
      paper: window.__WB_M3__?.getSpecification()?.header?.text,
      page: window.__WB_M3__?.getSpecification()?.page?.paper,
      header: document.querySelector('#paperHeader')?.textContent,
      pageLabel: window.__WB_M6__?.getPreview()?.pageLabel,
    }));
    assert(afterTemplate.page === 'A4' && afterTemplate.header === 'MarkPivot 参考页眉', `工作台上传模板未重生成规范初稿：${JSON.stringify(afterTemplate)}`);
    assert(/第\s*\d+\s*\/\s*\d+\s*页/.test(afterTemplate.pageLabel || ''), `页码文案不正确：${afterTemplate.pageLabel}`);
    await page.locator('#fileRailList .mp-file-name', { hasText: 'manual-heading.docx' }).click();
    await page.waitForFunction(() => {
      const map = window.__WB_M5__?.getMap();
      return map && /manual-heading\.docx$/i.test(map.fileName || '') && map.reviewCount > 0;
    }, null, { timeout: 20_000 });
    const chapter = await page.evaluate(() => (window.__WB_M5__?.getMap()?.blocks || []).find((block) => String(block.text).includes('第一章')));
    assert(chapter && chapter.id, '手工标题未进入语义列表。');
    await page.evaluate((id) => window.__WB_M5__.setType(id, 'heading1', true), chapter.id);
    await page.waitForFunction(() => /#\s*第一章/.test(document.querySelector('#mdEditor')?.value || ''), null, { timeout: 8_000 });
    await page.locator('#fileRailList .mp-file-name', { hasText: 'sample.docx' }).click();
    await page.waitForFunction(() => /sample\.docx$/i.test(window.__WB_M6__?.getPreview()?.fileName || ''), null, { timeout: 20_000 });
    await page.locator('#fileRailList .mp-file-name', { hasText: 'manual-heading.docx' }).click();
    try {
      await page.waitForFunction(() => /manual-heading\.docx$/i.test(window.__WB_M6__?.getPreview()?.fileName || '') && /#\s*第一章/.test(document.querySelector('#mdEditor')?.value || ''), null, { timeout: 20_000 });
    } catch (error) {
      const snap = await page.evaluate(() => ({
        fileName: window.__WB_M6__?.getPreview()?.fileName,
        md: String(document.querySelector('#mdEditor')?.value || '').slice(0, 240),
        stored: String(window.__WB_M6__?.getPreview()?.md || '').slice(0, 240),
      }));
      throw new Error(`切回文件未保留确认后的 Markdown：${error.message} · ${JSON.stringify(snap)}`);
    }
    await page.screenshot({ path: m6WorkspaceScreenshotPath, fullPage: false });
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `即时预览 18px · 原文不套用目标规范 · 恢复模板 · 上传模板页眉 · 切文件保留 # 第一章`;
  });

  await check('M7 Rule Engine 涂绘、作用域与原文重渲染', async () => {
    await mkdir(m7Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.locator('#homeInput').setInputFiles([fixturePaths.docx, fixturePaths.manualHeading]);
    await page.waitForFunction(() => window.__WB_M1__?.getTask()?.fileCount >= 2, null, { timeout: 8_000 });
    await page.locator('#homeNextBtn').click();
    await page.locator('#choiceBlankBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    await page.locator('#previewToolbar').waitFor({ state: 'visible', timeout: 8_000 });
    await page.waitForFunction(() => window.__WB_M5__?.getMap()?.blocks?.length > 0, null, { timeout: 20_000 });
    await page.locator('#fileRailList .mp-file-name', { hasText: 'manual-heading.docx' }).click();
    await page.waitForFunction(() => {
      const map = window.__WB_M5__?.getMap();
      const chapter = (map?.blocks || []).find((block) => String(block.text).includes('第一章'));
      return map && /manual-heading\.docx$/i.test(map.fileName || '') && chapter && chapter.semanticType === 'heading1';
    }, null, { timeout: 20_000 });
    await page.waitForFunction(() => window.__WB_M7__?.headingFont() === '27px', null, { timeout: 8_000 });
    const mdBefore = await page.locator('#mdEditor').inputValue();
    assert(!/#\s*第一章/.test(mdBefore), '未确认启发式标题已写入 Markdown。');
    const painted = await page.evaluate(() => (window.__WB_M7__?.painted() || []).find((item) => item.type === 'heading1'));
    assert(painted && painted.fontSize === '27px' && /第一章/.test(painted.text || ''), `启发式标题未按 H1 涂绘：${JSON.stringify(painted)}`);
    await page.locator('[data-spec-tab="heading"]').click();
    await page.locator('#specH1Size').fill('30px');
    await page.waitForFunction(() => window.__WB_M7__?.headingFont() === '30px', null, { timeout: 5_000 });
    assert(await page.locator('#mdEditor').inputValue() === mdBefore, '修改标题规范改写了 Markdown。');
    await page.locator('[data-preview-mode="source"]').click();
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.mode === 'source' && (window.__WB_M7__?.painted() || []).length === 0, null, { timeout: 5_000 });
    assert((await page.evaluate(() => window.__WB_M7__?.headingFont())) === '', '原文模式仍保留规则涂绘。');
    await page.locator('[data-preview-mode="effect"]').click();
    await page.waitForFunction(() => window.__WB_M6__?.getPreview()?.mode === 'effect' && window.__WB_M7__?.headingFont() === '30px', null, { timeout: 5_000 });
    const taskFingerprint = await page.evaluate(() => window.__WB_M7__?.fingerprint());
    await page.locator('[data-spec-scope="file"]').click();
    await page.locator('#specH1Size').fill('32px');
    await page.waitForFunction(() => window.__WB_M7__?.headingFont() === '32px', null, { timeout: 5_000 });
    assert((await page.evaluate(() => window.__WB_M3__?.getSpecification()?.heading1?.size)) === '32px', '当前文件 override 未写入 H1。');
    assert((await page.evaluate(() => window.__WB_M7__?.taskRules()?.heading1?.size)) === '30px', '文件 override 污染了任务规范。');
    await page.locator('#fileRailList .mp-file-name', { hasText: 'sample.docx' }).click();
    await page.waitForFunction(() => /sample\.docx$/i.test(window.__WB_M6__?.getPreview()?.fileName || '') && window.__WB_M3__?.getSpecification()?.heading1?.size === '30px', null, { timeout: 20_000 });
    await page.locator('#fileRailList .mp-file-name', { hasText: 'manual-heading.docx' }).click();
    await page.waitForFunction(() => /manual-heading\.docx$/i.test(window.__WB_M6__?.getPreview()?.fileName || '') && window.__WB_M7__?.headingFont() === '32px', null, { timeout: 20_000 });
    assert(await page.locator('#mdEditor').inputValue() === mdBefore, '切换规范或文件改写了 Markdown。');
    assert((await page.evaluate(() => window.__WB_M7__?.fingerprint())) === taskFingerprint, '涂绘过程改写了任务规则指纹。');
    await page.screenshot({ path: m7WorkspaceScreenshotPath, fullPage: false });
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return '启发式 H1 27px 不写 MD · 改 30px · 原文清除 · 文件 override 32px';
  });

  await check('M8 导出页、ZIP 结构与失败隔离', async () => {
    await mkdir(m8Dir, { recursive: true });
    const sample = await readFile(fixturePaths.docx);
    const heading = await readFile(fixturePaths.manualHeading);
    await startTaskToExport(page, [
      { name: 'sample.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: sample },
      { name: 'bad.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from('not a valid docx package', 'utf8') },
      { name: 'manual-heading.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: heading },
    ]);
    await page.screenshot({ path: m8ExportScreenshotPath, fullPage: false });
    await waitForDownload(page, '#exportAllBtn', exportedPaths.taskZip, 120_000);
    const summary = zipSummary(await readFile(exportedPaths.taskZip));
    const docxNames = summary.names.filter((name) => name.endsWith('.docx')).sort();
    assert(docxNames.includes('01 sample.docx') && docxNames.includes('03 manual-heading.docx'), `ZIP 缺少编号 Word：${docxNames.join(', ')}`);
    assert(docxNames.length === 2 && !docxNames.some((name) => /bad/i.test(name)), `失败文件不应进入 ZIP：${docxNames.join(', ')}`);
    assert(summary.zip.has('specification.json') && summary.zip.has('处理记录.txt'), 'ZIP 缺少 specification.json 或处理记录。');
    const spec = JSON.parse(summary.zip.readText('specification.json'));
    assert(spec.specification && spec.specification.body, 'specification.json 缺少任务规范。');
    assert(Array.isArray(spec.files) && spec.files.length === 3, `specification.json 未列出全部任务文件：${JSON.stringify(spec.files)}`);
    const record = summary.zip.readText('处理记录.txt');
    assert(/bad\.docx/.test(record) && /失败/.test(record), '处理记录未记录失败文件。');
    docxNames.forEach((name) => {
      const inner = zipSummary(summary.zip.read(name));
      assert(inner.zip.has('word/document.xml'), `${name} 缺少 word/document.xml`);
    });
    const snap = await page.evaluate(() => window.__WB_M8__.snapshot());
    assert(snap.results.filter((item) => item.status === 'done').length === 2, `成功数不正确：${JSON.stringify(snap.results)}`);
    assert(snap.results.filter((item) => item.status === 'error').length === 1, `失败隔离不正确：${JSON.stringify(snap.results)}`);
    return `ZIP ${docxNames.join(' / ')} · specification.json · 失败隔离 bad.docx`;
  });

  await check('M8 20 个文件可连续导出', async () => {
    const sample = await readFile(fixturePaths.docx);
    const files = Array.from({ length: 20 }, (_, index) => ({
      name: `doc-${String(index + 1).padStart(2, '0')}.docx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: sample,
    }));
    await startTaskToExport(page, files);
    await waitForDownload(page, '#exportAllBtn', exportedPaths.taskZip20, 180_000);
    const summary = zipSummary(await readFile(exportedPaths.taskZip20));
    const docxNames = summary.names.filter((name) => name.endsWith('.docx')).sort();
    assert(docxNames.length === 20, `20 文件 ZIP 数量错误：${docxNames.length} · ${docxNames.slice(0, 5).join(', ')}`);
    assert(docxNames[0] === '01 doc-01.docx' && docxNames[19] === '20 doc-20.docx', `编号不稳定：${docxNames[0]} / ${docxNames[19]}`);
    assert(summary.zip.has('specification.json') && summary.zip.has('处理记录.txt'), '20 文件 ZIP 缺少规范或记录。');
    const first = zipSummary(summary.zip.read('01 doc-01.docx'));
    assert(first.zip.has('word/document.xml'), '20 文件 ZIP 内 DOCX 缺少 document.xml。');
    return '20 个 Word 连续写入 ZIP · 01–20 编号稳定';
  });

  await check('M8 导出可取消', async () => {
    await startTaskToExport(page, [fixturePaths.docx, fixturePaths.manualHeading]);
    const downloadPromise = page.waitForEvent('download', { timeout: 90_000 }).catch(() => null);
    await page.locator('#exportAllBtn').click();
    await page.waitForFunction(() => window.__WB_M8__?.snapshot()?.running === true, null, { timeout: 8_000 });
    await page.evaluate(() => window.__WB_M8__.cancel());
    await page.waitForFunction(() => window.__WB_M8__?.snapshot()?.running === false, null, { timeout: 90_000 });
    const snap = await page.evaluate(() => window.__WB_M8__.snapshot());
    assert(snap.results.some((item) => item.status === 'cancelled') || snap.results.filter((item) => item.status === 'done').length < snap.results.length, `取消未打断剩余文件：${JSON.stringify(snap.results)}`);
    assert(snap.results.every((item) => item.status !== 'running'), `取消后仍有运行中文件：${JSON.stringify(snap.results)}`);
    const download = await downloadPromise;
    if (download) await download.saveAs(exportedPaths.taskCancelZip);
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return `取消后 ${snap.results.map((item) => item.status).join('/')}`;
  });

  await check('M9 调度档位与 Worker ping', async () => {
    const plans = await page.evaluate(() => {
      const plan = window.__WB_M9__.plan;
      return {
        normal: plan(0, 0).tier,
        p100: plan(0, 120).tier,
        mb5: plan(5242880, 0).tier,
        mb20: plan(20971520, 0).tier,
        mb50: plan(52428800, 2500).tier,
      };
    });
    assert(plans.normal === 'normal' && plans.p100 === '100p' && plans.mb5 === '5mb' && plans.mb20 === '20mb' && plans.mb50 === '50mb', `档位错误：${JSON.stringify(plans)}`);
    const ping = await page.evaluate(() => window.__WB_M9__.pingWorker());
    assert(ping && ping.ok, `Worker ping 失败：${JSON.stringify(ping)}`);
    return `档位 ${plans.normal}/${plans.p100}/${plans.mb5}/${plans.mb20}/${plans.mb50} · ping ok`;
  });

  await check('M9 大文件分片导入可见进度', async () => {
    await mkdir(m9Dir, { recursive: true });
    const large = await readFile(fixturePaths.largeDocx);
    await startTaskToWorkspace(page, [{
      name: 'large-doc.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: large,
    }]);
    await page.waitForFunction(() => window.__WB_M9__?.snapshot()?.running === false && (window.__WB_M9__?.snapshot()?.yields || 0) > 0, null, { timeout: 90_000 });
    const snap = await page.evaluate(() => window.__WB_M9__.snapshot());
    assert(snap.yields > 0, `大文件导入未让出主线程：${JSON.stringify(snap)}`);
    assert(snap.tier === '5mb' || snap.tier === '100p' || snap.tier === '20mb', `未命中大文件档位：${snap.tier}`);
    await page.screenshot({ path: m9WorkspaceScreenshotPath, fullPage: false });
    return `tier=${snap.tier} · yields=${snap.yields} · ${snap.files.map((item) => item.status).join('/')}`;
  });

  await check('M9 分片任务可取消', async () => {
    await page.evaluate(() => { window.__WB_M9__.runDummy(2000); });
    await page.waitForFunction(() => window.__WB_M9__?.snapshot()?.running === true, null, { timeout: 8_000 });
    await page.screenshot({ path: m9JobScreenshotPath, fullPage: false });
    await page.evaluate(() => window.__WB_M9__.cancel());
    await page.waitForFunction(() => window.__WB_M9__?.snapshot()?.running === false, null, { timeout: 30_000 });
    const snap = await page.evaluate(() => window.__WB_M9__.snapshot());
    assert(snap.cancelled === true, `取消标记未生效：${JSON.stringify(snap)}`);
    assert(snap.progress < snap.total || snap.kind === 'test', `取消后进度异常：${JSON.stringify(snap)}`);
    const bannerHidden = await page.locator('#jobBanner').evaluate((node) => node.hidden);
    assert(bannerHidden, '取消后进度条仍显示。');
    return `取消 · progress ${snap.progress}/${snap.total} · yields=${snap.yields}`;
  });

  await check('M9 单文件失败不拖垮任务', async () => {
    const sample = await readFile(fixturePaths.docx);
    await startTaskToWorkspace(page, [
      { name: 'bad.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from('not a valid docx package', 'utf8') },
      { name: 'sample.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: sample },
    ]);
    await page.waitForFunction(() => (window.__WB_M9__?.snapshot()?.files || []).some((item) => item.status === '导入失败'), null, { timeout: 60_000 });
    let snap = await page.evaluate(() => window.__WB_M9__.snapshot());
    assert(snap.files.length === 2, `任务文件数被失败文件打断：${JSON.stringify(snap.files)}`);
    assert(snap.files.some((item) => item.name === 'bad.docx' && item.status === '导入失败'), `损坏文件未隔离：${JSON.stringify(snap.files)}`);
    await page.locator('#fileRailList .mp-file-name', { hasText: 'sample.docx' }).click();
    await page.waitForFunction(() => (window.__WB_M9__?.snapshot()?.files || []).some((item) => item.name === 'sample.docx' && /已载入|需确认/.test(item.status)), null, { timeout: 60_000 });
    snap = await page.evaluate(() => window.__WB_M9__.snapshot());
    assert(snap.files.some((item) => item.name === 'sample.docx' && /已载入|需确认/.test(item.status)), `完好文件未能继续：${JSON.stringify(snap.files)}`);
    assert(snap.files.some((item) => item.name === 'bad.docx' && item.status === '导入失败'), `切换文件后失败状态丢失：${JSON.stringify(snap.files)}`);
    await page.goto(page.appUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    return snap.files.map((item) => `${item.name}=${item.status}`).join(' · ');
  });

  await check('M10 Windows Chrome 内核', async () => {
    const ua = await page.evaluate(() => navigator.userAgent);
    assert(/Windows NT/i.test(ua), `非 Windows 环境：${ua}`);
    assert(/Chrome\//i.test(ua), `非 Chrome 内核：${ua}`);
    const win = ua.match(/Windows NT [\d.]+/);
    const chrome = ua.match(/Chrome\/[\d.]+/);
    return `${win ? win[0] : 'Windows'} · ${chrome ? chrome[0] : 'Chrome'} · vendor 路径已验证`;
  });

  await check('M11 规范库增删改复制导入导出', async () => {
    await mkdir(m11Dir, { recursive: true });
    const homeUrl = String(page.appUrl).replace(/\?qa=.*$/, '');
    const importPath = path.join(m11Dir, 'import-spec.json');
    await writeFile(importPath, `${JSON.stringify({ name: '导入规范', body: { size: '19px' }, page: { paper: 'A3' } }, null, 2)}\n`, 'utf8');
    await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
    await waitForApp(page, { navigate: false });
    await page.evaluate(() => localStorage.removeItem('markpivot-spec-library-v1'));
    await page.locator('[data-app-view="specs"]').click();
    await page.locator('#specsView').waitFor({ state: 'visible', timeout: 5_000 });
    assert(await page.locator('#libraryNewBtn').isVisible(), '规范库缺少新建。');
    assert(await page.locator('#libraryEmpty').isVisible(), '空库提示未显示。');
    await page.screenshot({ path: m11LibraryScreenshotPath, fullPage: false });
    await page.locator('#libraryNewBtn').click();
    await page.locator('#libraryEditor').waitFor({ state: 'visible', timeout: 5_000 });
    await page.locator('#libraryEditorName').fill('投标模板');
    await page.locator('#libraryEditorTabs [data-lib-tab="body"]').click();
    await page.locator('#lib-specBodySize').fill('20px');
    await page.screenshot({ path: m11EditorScreenshotPath, fullPage: false });
    await page.locator('#libraryEditorSaveBtn').click();
    await page.waitForFunction(() => (window.__WB_M11__?.list() || []).some((item) => item.name === '投标模板' && item.bodySize === '20px'), null, { timeout: 5_000 });
    await page.locator('#specLibraryList [data-lib-action="copy"]').click();
    await page.waitForFunction(() => (window.__WB_M11__?.list() || []).length >= 2, null, { timeout: 5_000 });
    page.once('dialog', (dialog) => dialog.accept('正式投标'));
    await page.locator('#specLibraryList [data-lib-action="rename"]').first().click();
    await page.waitForFunction(() => (window.__WB_M11__?.list() || []).some((item) => item.name === '正式投标'), null, { timeout: 5_000 });
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8_000 }),
      page.locator('#specLibraryList [data-lib-action="export"]').first().click(),
    ]);
    const downloadPath = path.join(m11Dir, download.suggestedFilename() || 'exported-spec.json');
    await download.saveAs(downloadPath);
    const exported = JSON.parse(await readFile(downloadPath, 'utf8'));
    assert(exported.body && exported.page, `导出 JSON 不是 Specification：${Object.keys(exported).join(',')}`);
    page.once('dialog', (dialog) => dialog.accept());
    await page.locator('#specLibraryList .mp-library-card', { hasText: '副本' }).locator('[data-lib-action="delete"]').click();
    await page.waitForFunction(() => !(window.__WB_M11__?.list() || []).some((item) => String(item.name).includes('副本')), null, { timeout: 5_000 });
    await page.locator('#libraryJsonInput').setInputFiles(importPath);
    await page.waitForFunction(() => (window.__WB_M11__?.list() || []).some((item) => item.name === '导入规范' && item.bodySize === '19px' && item.paper === 'A3'), null, { timeout: 5_000 });
    await page.locator('#libraryWordInput').setInputFiles(fixturePaths.referenceDocx);
    await page.locator('#libraryEditor').waitFor({ state: 'visible', timeout: 20_000 });
    assert((await page.evaluate(() => window.__WB_M11__?.mode())) === 'create', '从 Word 创建未打开编辑器。');
    await page.locator('#libraryEditorSaveBtn').click();
    await page.waitForFunction(() => (window.__WB_M11__?.list() || []).some((item) => item.source === 'word'), null, { timeout: 5_000 });
    const names = await page.evaluate(() => (window.__WB_M11__?.list() || []).map((item) => item.name));
    return `保存=${names.join(' / ')} · JSON roundtrip · Word 创建`;
  });

  await check('M11 规范库不覆盖当前任务', async () => {
    await startTaskToWorkspace(page, fixturePaths.docx);
    const before = await page.evaluate(() => ({
      size: window.__WB_M3__?.getSpecification()?.body?.size,
      lib: (window.__WB_M2__?.library() || []).length,
    }));
    await page.locator('[data-spec-tab="body"]').click();
    await page.locator('#specBodySize').fill('18px');
    await page.waitForFunction(() => window.__WB_M3__?.getSpecification()?.body?.size === '18px', null, { timeout: 5_000 });
    assert((await page.evaluate(() => (window.__WB_M2__?.library() || []).length)) === before.lib, '任务内修改写入了规范库。');
    await page.locator('[data-app-view="specs"]').click();
    await page.locator('#specLibraryList [data-lib-action="edit"]').first().click();
    await page.locator('#libraryEditor').waitFor({ state: 'visible', timeout: 5_000 });
    await page.locator('#libraryEditorTabs [data-lib-tab="body"]').click();
    await page.locator('#lib-specBodySize').fill('22px');
    await page.locator('#libraryEditorSaveBtn').click();
    await page.locator('#resumeTaskBtn').click();
    await page.locator('#specRail').waitFor({ state: 'visible', timeout: 8_000 });
    const after = await page.evaluate(() => ({
      size: window.__WB_M3__?.getSpecification()?.body?.size,
      libraryHas22: (window.__WB_M11__?.list() || []).some((item) => item.bodySize === '22px'),
    }));
    assert(after.size === '18px', `编辑规范库改写了任务规范：${after.size}`);
    assert(after.libraryHas22, '规范库编辑未保存。');
    return `任务 18px 未变 · 规范库已更新 22px · 库 ${before.lib} 条`;
  });
}

async function smokeBrowserChannel(appUrl, channel, label) {
  let extra;
  try {
    extra = await chromium.launch({ channel, headless: true });
  } catch (error) {
    skip(`M10 ${label} 本地运行库`, `${label} 通道不可用，已用 Playwright Chromium 覆盖 Windows Chrome 内核`);
    return;
  }
  try {
    await check(`M10 ${label} 本地运行库`, async () => {
      const page = await extra.newPage({ viewport: { width: 1440, height: 900 } });
      try {
        const homeUrl = String(appUrl).replace(/\?qa=.*$/, '');
        await page.goto(homeUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
        await waitForApp(page, { navigate: false });
        const sources = await page.evaluate(() => (window.__WB_M10__ && window.__WB_M10__.sources()) || []);
        assert(sources.length === 9 && sources.every((source) => source === 'vendor'), `${label} 未走本地运行库：${sources.join('/')}`);
        const text = await page.locator('#libraryStatus').textContent();
        assert(text.includes('本地运行库可用'), `${label} 状态不正确：${text}`);
        return `${label} · vendor 9/9 · 可离线`;
      } finally {
        await page.close().catch(() => {});
      }
    });
  } finally {
    if (extra) await extra.close().catch(() => {});
  }
}

async function main() {
  await mkdir(baselineDir, { recursive: true });
  await mkdir(artifactDir, { recursive: true });
  await mkdir(m1Dir, { recursive: true });
  await mkdir(m2Dir, { recursive: true });
  await mkdir(m10Dir, { recursive: true });
  await mkdir(m11Dir, { recursive: true });
  await checkStaticContracts();

  const server = makeStaticServer();
  let browser;
  let page;
  let appUrl = '';
  try {
    const port = await listen(server);
    appUrl = `http://127.0.0.1:${port}/${encodeURIComponent('MarkPivot.html')}?qa=m1-${Date.now()}`;
    try {
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1440, height: 1100 }, deviceScaleFactor: 1 });
      page.appUrl = appUrl;
      await page.addInitScript(() => {
        const store = {
          name: 'linked.md',
          content: '# Disk baseline\n\n原始内容',
          lastModified: 100,
          permissionMode: 'readwrite',
          openCalls: 0,
          pending: '',
        };
        const handle = {
          kind: 'file',
          get name() {
            return store.name;
          },
          async queryPermission() {
            if (store.permissionMode === 'denied') return 'denied';
            return store.permissionMode === 'readwrite' ? 'granted' : 'prompt';
          },
          async requestPermission() {
            return store.permissionMode === 'readwrite' ? 'granted' : 'denied';
          },
          async getFile() {
            return new File([store.content], store.name, { type: 'text/markdown', lastModified: store.lastModified });
          },
          async createWritable() {
            return {
              async write(value) {
                store.pending = typeof value === 'string' ? value : await value.text();
              },
              async close() {
                store.content = store.pending;
                store.pending = '';
                store.lastModified += 1;
              },
              async abort() {
                store.pending = '';
              },
            };
          },
        };
        const picker = async () => {
          store.openCalls += 1;
          return [handle];
        };
        store.handle = handle;
        store.picker = picker;
        window.__WB_F2_TEST__ = store;
        try {
          Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, writable: true, value: picker });
        } catch (error) {
          window.showOpenFilePicker = picker;
        }
      });
      page.on('console', (message) => {
        if (message.type() === 'error') browserEvents.consoleErrors.push(message.text());
        if (message.type() === 'warning') browserEvents.consoleWarnings.push(message.text());
      });
      page.on('pageerror', (error) => browserEvents.pageErrors.push(error.message));
      page.on('response', (response) => {
        if (response.status() >= 400 && !/\.map(?:\?|$)/i.test(response.url())) browserEvents.httpErrors.push(`${response.status()} ${response.url()}`);
      });
      await runBrowserChecks(page);
      await smokeBrowserChannel(appUrl, 'chrome', 'Chrome');
      await smokeBrowserChannel(appUrl, 'msedge', 'Edge');
    } catch (error) {
      const detail = error?.message || String(error);
      if (/executable doesn't exist|browserType\.launch|Executable doesn't exist/i.test(detail)) {
        fail('浏览器运行时可用', `${detail}。请先执行 npx playwright install chromium。`);
      } else {
        fail('浏览器 QA 执行', detail);
      }
    }
  } catch (error) {
    fail('QA 静态服务启动', error);
  } finally {
    if (page) await page.close().catch(() => {});
    if (browser) await browser.close().catch(() => {});
    await closeServer(server);
  }

  await check('浏览器无 console/pageerror', () => {
    assert(browserEvents.consoleErrors.length === 0, `console.error=${browserEvents.consoleErrors.join(' | ')}`);
    assert(browserEvents.pageErrors.length === 0, `pageerror=${browserEvents.pageErrors.join(' | ')}`);
    assert(browserEvents.httpErrors.length === 0, `HTTP 错误=${browserEvents.httpErrors.join(' | ')}`);
    return `console.error=0 · pageerror=0 · http=0 · warning=${browserEvents.consoleWarnings.length}`;
  });

  const screenshotResults = await Promise.all(Object.entries(screenshotPaths).map(async ([name, filePath]) => {
    const exists = existsSync(filePath);
    return { name, path: filePath, exists: exists && screenshotCreated.has(filePath), bytes: exists ? (await stat(filePath)).size : 0 };
  }));
  await check('三张基线截图已落盘', () => {
    assert(screenshotResults.every((item) => item.exists && item.bytes > 0), `截图缺失：${screenshotResults.filter((item) => !item.exists).map((item) => item.name).join(', ')}`);
    return screenshotResults.map((item) => `${item.name}=${formatBytes(item.bytes)}`).join(' · ');
  });

  const passed = checks.filter((item) => item.status === 'PASS').length;
  const failed = checks.filter((item) => item.status === 'FAIL').length;
  const skipped = checks.filter((item) => item.status === 'SKIP').length;
  const htmlStat = await stat(appFile).catch(() => ({ size: 0 }));
  const report = {
    version: 'M11',
    generatedAt: new Date().toISOString(),
    appUrl,
    html: { path: appFile, bytes: htmlStat.size, kb: Number((htmlStat.size / 1024).toFixed(1)) },
    fixtures: fixturePaths,
    exportedArtifacts: exportedPaths,
    screenshots: screenshotResults,
    visualArtifacts: {
      v2MobileScreenshot: v2MobileScreenshotPath,
      v2StyleScreenshot: v2StyleScreenshotPath,
      f3CoverDesktopScreenshot: f3CoverDesktopScreenshotPath,
      f3CoverMobileScreenshot: f3CoverMobileScreenshotPath,
      v4TableDesktopScreenshot: v4TableDesktopScreenshotPath,
      v4TableMobileScreenshot: v4TableMobileScreenshotPath,
      v4WideTableScreenshot: v4WideTableScreenshotPath,
      m1HomeScreenshot: m1HomeScreenshotPath,
      m1ChoiceScreenshot: m1ChoiceScreenshotPath,
      m1WorkspaceScreenshot: m1WorkspaceScreenshotPath,
      m2WorkspaceScreenshot: m2WorkspaceScreenshotPath,
      m4ExtractScreenshot: m4ExtractScreenshotPath,
      m4WorkspaceScreenshot: m4WorkspaceScreenshotPath,
      m5ReviewScreenshot: m5ReviewScreenshotPath,
      m5WorkspaceScreenshot: m5WorkspaceScreenshotPath,
      m6WorkspaceScreenshot: m6WorkspaceScreenshotPath,
      m7WorkspaceScreenshot: m7WorkspaceScreenshotPath,
      m8ExportScreenshot: m8ExportScreenshotPath,
      m9WorkspaceScreenshot: m9WorkspaceScreenshotPath,
      m9JobScreenshot: m9JobScreenshotPath,
      m10HomeScreenshot: m10HomeScreenshotPath,
      m11LibraryScreenshot: m11LibraryScreenshotPath,
      m11EditorScreenshot: m11EditorScreenshotPath,
    },
    browser: browserEvents,
    summary: { passed, failed, skipped },
    checks,
  };
  await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  console.log(JSON.stringify(report, null, 2));
  if (failed > 0) process.exitCode = 1;
}

await main();
