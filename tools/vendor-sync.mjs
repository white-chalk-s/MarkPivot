import { copyFile, mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const vendorDir = path.join(projectRoot, 'vendor');
const templatesDir = path.join(projectRoot, 'templates');
const fixtureTemplate = path.join(projectRoot, 'tools', 'fixtures', 'reference-template.docx');
const productTemplate = path.join(templatesDir, 'reference-template.docx');

export const VENDOR_LIBRARIES = [
  {
    id: 'JSZip',
    global: 'JSZip',
    version: '3.10.1',
    file: 'jszip-3.10.1.min.js',
    url: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
    minBytes: 40_000,
  },
  {
    id: 'docx',
    global: 'docx',
    version: '8.5.0',
    file: 'docx-8.5.0.umd.js',
    url: 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
    minBytes: 200_000,
  },
  {
    id: 'XLSX',
    global: 'XLSX',
    version: '0.18.5',
    file: 'xlsx-0.18.5.full.min.js',
    url: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    minBytes: 400_000,
  },
  {
    id: 'markdownit',
    global: 'markdownit',
    version: '14.1.0',
    file: 'markdown-it-14.1.0.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it@14.1.0/dist/markdown-it.min.js',
    minBytes: 20_000,
  },
  {
    id: 'markdownitFootnote',
    global: 'markdownitFootnote',
    version: '4.0.0',
    file: 'markdown-it-footnote-4.0.0.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it-footnote@4.0.0/dist/markdown-it-footnote.min.js',
    minBytes: 1_000,
  },
  {
    id: 'markdownitTaskLists',
    global: 'markdownitTaskLists',
    version: '2.1.1',
    file: 'markdown-it-task-lists-2.1.1.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it-task-lists@2.1.1/dist/markdown-it-task-lists.min.js',
    minBytes: 1_000,
  },
  {
    id: 'markdownitMark',
    global: 'markdownitMark',
    version: '4.0.0',
    file: 'markdown-it-mark-4.0.0.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it-mark@4.0.0/dist/markdown-it-mark.min.js',
    minBytes: 400,
  },
  {
    id: 'markdownitSub',
    global: 'markdownitSub',
    version: '2.0.0',
    file: 'markdown-it-sub-2.0.0.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it-sub@2.0.0/dist/markdown-it-sub.min.js',
    minBytes: 200,
  },
  {
    id: 'markdownitSup',
    global: 'markdownitSup',
    version: '2.0.0',
    file: 'markdown-it-sup-2.0.0.min.js',
    url: 'https://cdn.jsdelivr.net/npm/markdown-it-sup@2.0.0/dist/markdown-it-sup.min.js',
    minBytes: 200,
  },
];

async function download(url) {
  const response = await fetch(url, { cache: 'no-store' });
  if (!response.ok) throw new Error(`${url} → HTTP ${response.status}`);
  return Buffer.from(await response.arrayBuffer());
}

function stripSourceMap(buffer) {
  const text = buffer.toString('utf8').replace(/\n\/\/[#@][ \t]*sourceMappingURL=.*$/gm, '');
  return Buffer.from(text, 'utf8');
}

async function ensureLibrary(item) {
  const target = path.join(vendorDir, item.file);
  if (existsSync(target)) {
    const current = await readFile(target);
    const stripped = stripSourceMap(current);
    if (stripped.length >= item.minBytes) {
      if (!current.equals(stripped)) await writeFile(target, stripped);
      return { id: item.id, file: item.file, bytes: stripped.length, action: current.equals(stripped) ? 'keep' : 'strip' };
    }
  }
  const body = stripSourceMap(await download(item.url));
  if (body.length < item.minBytes) {
    throw new Error(`${item.id} 下载过小：${body.length} B，预期至少 ${item.minBytes} B`);
  }
  await writeFile(target, body);
  return { id: item.id, file: item.file, bytes: body.length, action: 'download' };
}

async function ensureTemplate() {
  await mkdir(templatesDir, { recursive: true });
  if (!existsSync(fixtureTemplate)) {
    return { copied: false, path: productTemplate };
  }
  await copyFile(fixtureTemplate, productTemplate);
  const fileStat = await stat(productTemplate);
  return { copied: true, path: productTemplate, bytes: fileStat.size };
}

export async function syncVendor() {
  await mkdir(vendorDir, { recursive: true });
  const libraries = [];
  for (const item of VENDOR_LIBRARIES) {
    libraries.push(await ensureLibrary(item));
  }
  const template = await ensureTemplate();
  const manifest = {
    version: 1,
    generatedAt: new Date().toISOString(),
    libraries: VENDOR_LIBRARIES.map((item, index) => ({
      id: item.id,
      global: item.global,
      version: item.version,
      file: item.file,
      local: `vendor/${item.file}`,
      bytes: libraries[index].bytes,
    })),
    template: template.copied ? { path: 'templates/reference-template.docx', bytes: template.bytes } : null,
  };
  await writeFile(path.join(vendorDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  return { vendorDir, libraries, template, manifest };
}

const isMain = process.argv[1] && path.normalize(path.resolve(process.argv[1])) === path.normalize(fileURLToPath(import.meta.url));
if (isMain) {
  const result = await syncVendor();
  console.log(JSON.stringify(result, null, 2));
}
