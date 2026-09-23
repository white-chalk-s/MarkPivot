import { mkdir, stat, writeFile, copyFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  FootnoteReferenceRun,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import ExcelJS from 'exceljs';

const toolsDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtureDir = path.join(toolsDir, 'fixtures');
const imageDir = path.join(fixtureDir, 'images');
const mdPath = path.join(fixtureDir, 'sample.md');
const coverPath = path.join(fixtureDir, 'cover.md');
const v4TablePath = path.join(fixtureDir, 'v4-table.md');
const v4WideTablePath = path.join(fixtureDir, 'v4-wide-table.md');
const docxPath = path.join(fixtureDir, 'sample.docx');
const referenceDocxPath = path.join(fixtureDir, 'reference-template.docx');
const manualHeadingDocxPath = path.join(fixtureDir, 'manual-heading.docx');
const largeDocxPath = path.join(fixtureDir, 'large-doc.docx');
const xlsxPath = path.join(fixtureDir, 'sample.xlsx');
const imagePath = path.join(imageDir, 'sample.png');
const codeTick = String.fromCharCode(96);
const fence = codeTick.repeat(3);
const imageBuffer = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
);

const sampleMarkdown = [
  '---',
  'title: M3.1 语义黄金样本',
  'tags:',
  '  - qa',
  '  - 文档互转',
  'aliases: [语义回归]',
  '---',
  '',
  '# M3.2 黄金样本文档',
  '',
  '::title 可回归的真实文件样本',
  '',
  '::desc 这个样本覆盖 Markdown、Word 和 Excel 的关键互转路径。',
  '',
  '## H2 结构',
  '',
  '### H3 结构',
  '',
  '#### H4 结构',
  '',
  '##### H5 结构',
  '',
  '###### H6 结构',
  '',
  '这是一段包含 **加粗**、*斜体*、***粗斜体***、~~删除线~~、' + codeTick + '行内代码' + codeTick + '、<u>下划线</u>、<sup>上标</sup>、<sub>下标</sub>、==高亮==、[外部链接](https://example.com) 和脚注[^1] 的正文。#回归 #文档，跳转到 [[#H2 结构|H2 结构]]，打开 [[项目说明|项目说明]]，图片 ![[images/missing.png]]，笔记 ![[相关笔记#结论|相关笔记]]。',
  '',
  '- 一级列表',
  '  - 二级列表',
  '    - 三级列表',
  '- [ ] 待处理任务',
  '- [x] 已完成任务',
  '',
  '1. 有序项目',
  '   1. 嵌套项目',
  '',
  '| 指标 | 数值 | 说明 |',
  '| :--- | ---: | :---: |',
  '| 销售额 | 1280000 | 82% |',
  '| 毛利率 | 0.32 | 稳定 |',
  '',
  '> 这是多段引用的第一段。',
  '>',
  '> 这是引用的第二段。',
  '',
  fence + 'text',
  'const result = convert(document);',
  '┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐',
  '│ WPS 等宽字符画回归：中文与 ASCII 必须保持 1:2 宽度并且整行不折行。ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 │',
  '└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘',
  fence,
  '',
  '::tip 工作台扩展 Callout',
  '',
  '> [!tip] 标准 Callout',
  '> 标准语法的多行正文会被统一渲染。',
  '',
  '> [!warn]',
  '> 公式表达式降级时必须进入转换报告。',
  '',
  '![测试图片](images/sample.png)',
  '',
  '---',
  '',
  '::pagebreak',
  '',
  '[^1]: 这是黄金样本中的 GFM 脚注。',
].join('\n');

const coverMarkdown = [
  '---',
  'title: F3 封面格式回归',
  'tags: [cover, qa]',
  '---',
  '',
  '::title F3 封面页样本',
  '',
  '::desc 这是封面说明：标题与说明居中，封面之后强制分页。',
  '',
  '# 封面后的正文',
  '',
  '正文首段用于确认封面后的普通段落仍按默认首行缩进处理。',
  '',
  '::title 文档中部普通标题',
  '',
  '文档中部的 title 不应再次触发封面。',
].join('\n');

const v4TableMarkdown = [
  '# V4 表格布局与缩进回归',
  '',
  '正文首段用于确认导出后只有 2 字符首行缩进。',
  '',
  '- 一级列表不应继承正文首行缩进',
  '  - 二级列表的编号位置需要紧凑',
  '',
  '| 组织机构 | 负责人 | 说明 |',
  '| --- | --- | --- |',
  '| 石家庄地铁集团 | 张三 | 负责城市轨道交通线路建设和运营协调 |',
  '| 站务中心 | 李四 | 短 |',
  '|  | 王五 | 另一个短值 |',
].join('\n');

const v4WideTableMarkdown = [
  '# 9 列主数据识别记录表',
  '',
  '15. 交易稳定性：与交易数据相比主数据是相对稳定的，变化频率较低。',
  '',
  '表 4-1 主数据识别记录表',
  '',
  '| 序号 | 数据项 | 特征一致性 | 识别唯一性 | 数据共享性 | 长期有效性 | 交易稳定性 | 是否主数据 | 来源系统 |',
  '| --- | --- | --- | --- | --- | --- | --- | --- | --- |',
  '| 1 | 员工编码 | ✓ | ✓ | ✓ | ✓ | ✓ | 是 | 人力系统 |',
  '| 2 | 姓名 | ✓ | ✓ | ✓ | ✓ | ✓ | 否 | ... |',
].join('\n');

const cellBorders = {
  top: { style: 'thin', color: { argb: 'FFD5D9CF' } },
  left: { style: 'thin', color: { argb: 'FFD5D9CF' } },
  bottom: { style: 'thin', color: { argb: 'FFD5D9CF' } },
  right: { style: 'thin', color: { argb: 'FFD5D9CF' } },
};

function styleHeader(row) {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF38553B' } };
  row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  row.border = cellBorders;
}

function styleBodyRow(row) {
  row.eachCell((cell) => {
    cell.border = cellBorders;
    cell.alignment = { vertical: 'middle', wrapText: true };
  });
}

async function buildDocx() {
  const children = [
    new Paragraph({ text: 'M3.2 真实文件回归样本', heading: HeadingLevel.TITLE }),
    new Paragraph({
      children: [
        new TextRun({ text: '这是 Word 黄金夹具，覆盖 ' }),
        new TextRun({ text: '加粗', bold: true }),
        new TextRun({ text: '、' }),
        new TextRun({ text: '斜体', italics: true }),
        new TextRun({ text: '、' }),
        new ExternalHyperlink({
          children: [new TextRun({ text: '外部链接' })],
          link: 'https://example.com/qa-fixture',
        }),
        new TextRun({ text: ' 和脚注' }),
        new FootnoteReferenceRun(1),
        new TextRun({ text: '。' }),
      ],
    }),
    new Paragraph({ text: '一级标题', heading: HeadingLevel.HEADING_1 }),
    new Paragraph({ text: '二级标题', heading: HeadingLevel.HEADING_2 }),
    new Paragraph({ text: '三级标题', heading: HeadingLevel.HEADING_3 }),
    new Paragraph({ text: '四级标题', heading: HeadingLevel.HEADING_4 }),
    new Paragraph({ text: '五级标题', heading: HeadingLevel.HEADING_5 }),
    new Paragraph({ text: '六级标题', heading: HeadingLevel.HEADING_6 }),
    new Paragraph({ text: '这是 Word 夹具正文，用来验证段落、行距和导出后的可读性。' }),
    new Paragraph({ text: '一级编号', numbering: { reference: 'qa-numbering', level: 0 } }),
    new Paragraph({ text: '二级编号', numbering: { reference: 'qa-numbering', level: 1 } }),
    new Paragraph({ text: '引用段落：转换报告必须保留降级原因。', style: 'Quote' }),
    new Paragraph({
      children: [
        new ImageRun({ data: imageBuffer, type: 'png', transformation: { width: 160, height: 80 } }),
      ],
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: '字段' })] }),
            new TableCell({ children: [new Paragraph({ text: '值' })] }),
            new TableCell({ children: [new Paragraph({ text: '备注' })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: '合并标题' })], columnSpan: 2 }),
            new TableCell({ children: [new Paragraph({ text: '保留 colspan' })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: '脚注' })] }),
            new TableCell({ children: [new Paragraph({ text: '原生脚注' })], rowSpan: 2 }),
            new TableCell({ children: [new Paragraph({ text: 'Word OOXML' })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: '图片' })] }),
            new TableCell({ children: [new Paragraph({ text: '资源池' })] }),
          ],
        }),
      ],
    }),
    new Paragraph({ text: '分页后的段落', pageBreakBefore: true }),
  ];

  const document = new Document({
    creator: '文档互转工作台 QA',
    description: 'M3.2 DOCX fixture',
    numbering: {
      config: [
        {
          reference: 'qa-numbering',
          levels: [
            { level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT },
            { level: 1, format: 'lowerLetter', text: '%2.', alignment: AlignmentType.LEFT },
          ],
        },
      ],
    },
    footnotes: {
      1: { children: [new Paragraph({ text: 'DOCX 黄金夹具脚注内容。' })] },
    },
    sections: [{ children }],
  });

  await writeFile(docxPath, await Packer.toBuffer(document));
}

async function buildLargeDocx() {
  const children = [
    new Paragraph({ text: 'M9 大文件性能夹具', heading: HeadingLevel.HEADING_1 }),
  ];
  for (let i = 1; i <= 220; i += 1) {
    children.push(new Paragraph({ text: `段落 ${i}：用于分片解析与进度让出，不作为黄金语义样本。` }));
  }
  for (let t = 1; t <= 20; t += 1) {
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: `表${t}A` })] }),
            new TableCell({ children: [new Paragraph({ text: `表${t}B` })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: '1' })] }),
            new TableCell({ children: [new Paragraph({ text: '2' })] }),
          ],
        }),
      ],
    }));
  }
  for (let i = 1; i <= 50; i += 1) {
    children.push(new Paragraph({
      children: [
        new ImageRun({ data: imageBuffer, type: 'png', transformation: { width: 8, height: 8 } }),
        new TextRun({ text: ` 图 ${i}` }),
      ],
    }));
  }
  const document = new Document({
    creator: '文档互转工作台 QA',
    description: 'M9 large-doc fixture',
    sections: [{ children }],
  });
  await writeFile(largeDocxPath, await Packer.toBuffer(document));
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function createStoredZip(entries) {
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, 'utf8');
    const checksum = crc32(data);
    const local = Buffer.alloc(30 + name.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    name.copy(local, 30);
    localParts.push(local, data);

    const central = Buffer.alloc(46 + name.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(localOffset, 42);
    name.copy(central, 46);
    centralParts.push(central);
    localOffset += local.length + data.length;
  }
  const centralDirectory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(localOffset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, centralDirectory, end]);
}

async function buildReferenceDocx() {
  const documentXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>',
    '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>V2 Word 模板参考标题</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>正文无显式颜色，应按 OOXML 默认黑处理。</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>主题色一级标题</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>无色二级标题</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>显式紫色三级标题</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading4"/></w:pPr><w:r><w:t>无色四级标题</w:t></w:r></w:p>',
    '<w:tbl><w:tblPr><w:tblStyle w:val="V2Table"/><w:tblBorders>',
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>',
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>',
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>',
    '</w:tblBorders></w:tblPr>',
    '<w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr><w:p><w:r><w:t>字段</w:t></w:r></w:p></w:tc>',
    '<w:tc><w:tcPr><w:shd w:val="clear" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc></w:tr>',
    '<w:tr><w:tc><w:p><w:r><w:t>主题色</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>已定义</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    '<w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    '</w:body></w:document>',
  ].join('');
  const stylesXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>',
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="24"/></w:rPr><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="36"/><w:color w:themeColor="accent1" w:themeShade="66"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="30"/><w:color w:themeColor="accent2" w:themeTint="33"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="26"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="24"/><w:color w:val="7030A0"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="Heading 4"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>',
    '<w:style w:type="table" w:styleId="V2Table"><w:name w:val="V2 Table"/><w:tblStylePr w:type="firstRow"><w:shd w:val="clear" w:themeFill="accent1" w:themeFillTint="33"/></w:tblStylePr></w:style>',
    '</w:styles>',
  ].join('');
  const themeXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>',
    '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>',
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>',
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>',
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="黑体"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="微软雅黑"/></a:minorFont></a:fontScheme>',
    '</a:themeElements></a:theme>',
  ].join('');
  const contentTypes = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
  ].join('');
  const rootRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
  const documentRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>';
  const headerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>MarkPivot 参考页眉</w:t></w:r></w:p></w:hdr>';
  const footerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>第 </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> 页</w:t></w:r></w:p></w:ftr>';
  const settingsXml = '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>';
  const coreXml = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>V2 Word 模板参考</dc:title><dc:creator>文档互转工作台 QA</dc:creator></cp:coreProperties>';
  const appXml = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>文档互转工作台 QA</Application></Properties>';
  await writeFile(referenceDocxPath, createStoredZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: documentRels },
    { name: 'word/header1.xml', data: headerXml },
    { name: 'word/footer1.xml', data: footerXml },
    { name: 'word/styles.xml', data: stylesXml },
    { name: 'word/theme/theme1.xml', data: themeXml },
    { name: 'word/settings.xml', data: settingsXml },
    { name: 'docProps/core.xml', data: coreXml },
    { name: 'docProps/app.xml', data: appXml },
  ]));
}

async function buildManualHeadingDocx() {
  const documentXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>本规范适用于内部文档。</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>第一章 总则</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>1.1 适用范围</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>本文件用于验证手工加粗标题与图表题识别，不应被误判为标题。</w:t></w:r></w:p>',
    '<w:tbl><w:tblPr><w:tblBorders>',
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
    '</w:tblBorders></w:tblPr>',
    '<w:tr><w:tc><w:p><w:r><w:t>字段</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc></w:tr>',
    '<w:tr><w:tc><w:p><w:r><w:t>标题样式</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>仅 Normal</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>表 1 字段说明</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>图 1 示意图</w:t></w:r></w:p>',
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    '</w:body></w:document>',
  ].join('');
  const stylesXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>',
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    '</w:styles>',
  ].join('');
  const contentTypes = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
  ].join('');
  const rootRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
  const documentRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const coreXml = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>手工标题语义夹具</dc:title><dc:creator>文档互转工作台 QA</dc:creator></cp:coreProperties>';
  const appXml = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>文档互转工作台 QA</Application></Properties>';
  await writeFile(manualHeadingDocxPath, createStoredZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: documentRels },
    { name: 'word/styles.xml', data: stylesXml },
    { name: 'docProps/core.xml', data: coreXml },
    { name: 'docProps/app.xml', data: appXml },
  ]));
}

async function buildXlsx() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = '文档互转工作台 QA';
  workbook.created = new Date('2026-09-02T00:00:00Z');

  const sales = workbook.addWorksheet('销售');
  sales.columns = [
    { key: 'month', width: 14 },
    { key: 'amount', width: 16 },
    { key: 'rate', width: 12 },
    { key: 'date', width: 16 },
    { key: 'result', width: 16 },
  ];
  sales.mergeCells('A1:E1');
  sales.getCell('A1').value = '销售测试工作簿';
  sales.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  sales.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF38553B' } };
  sales.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  sales.getRow(1).height = 28;
  sales.getRow(2).values = ['月份', '销售额', '完成率', '日期', '公式结果'];
  styleHeader(sales.getRow(2));
  sales.addRow(['2026-01', 1280000, 0.82, new Date('2026-01-31T00:00:00Z'), { formula: 'B3*1.1', result: 1408000 }]);
  sales.addRow(['2026-02', 1435000, 0.91, new Date('2026-02-28T00:00:00Z'), { formula: 'B4*1.1', result: 1578500 }]);
  sales.addRow(['换行单元格', 188000, 0.76, new Date('2026-03-31T00:00:00Z'), '第一行\n第二行']);
  sales.getCell('B3').numFmt = '¥#,##0.00';
  sales.getCell('B4').numFmt = '¥#,##0.00';
  sales.getCell('B5').numFmt = '¥#,##0.00';
  sales.getCell('C3').numFmt = '0.00%';
  sales.getCell('C4').numFmt = '0.00%';
  sales.getCell('C5').numFmt = '0.00%';
  sales.getCell('D3').numFmt = 'yyyy-mm-dd';
  sales.getCell('D4').numFmt = 'yyyy-mm-dd';
  sales.getCell('D5').numFmt = 'yyyy-mm-dd';
  sales.getCell('E3').numFmt = '#,##0.00';
  sales.getCell('E4').numFmt = '#,##0.00';
  sales.getCell('E5').alignment = { wrapText: true, vertical: 'middle' };
  sales.getRow(3).height = 22;
  sales.getRow(4).height = 22;
  sales.getRow(5).height = 38;
  styleBodyRow(sales.getRow(3));
  styleBodyRow(sales.getRow(4));
  styleBodyRow(sales.getRow(5));

  const inventory = workbook.addWorksheet('库存');
  inventory.columns = [{ key: 'sku', width: 16 }, { key: 'count', width: 12 }, { key: 'status', width: 18 }];
  inventory.addRow(['SKU', '库存', '状态']);
  styleHeader(inventory.getRow(1));
  inventory.addRow(['QA-001', 42, '正常']);
  inventory.addRow(['QA-002', 0, '需补货']);
  styleBodyRow(inventory.getRow(2));
  styleBodyRow(inventory.getRow(3));

  const notes = workbook.addWorksheet('备注');
  notes.columns = [{ key: 'item', width: 20 }, { key: 'content', width: 44 }];
  notes.addRow(['项目', '内容']);
  styleHeader(notes.getRow(1));
  notes.addRow(['说明', '此 Sheet 用于验证多 Sheet 映射与回导。']);
  notes.addRow(['换行', '第一行\n第二行']);
  notes.getCell('B3').alignment = { wrapText: true, vertical: 'middle' };
  notes.getRow(3).height = 34;
  styleBodyRow(notes.getRow(2));
  styleBodyRow(notes.getRow(3));

  await workbook.xlsx.writeFile(xlsxPath);
  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(xlsxPath);
  if (check.worksheets.length !== 3 || check.getWorksheet('销售')?.getCell('A1').value !== '销售测试工作簿') {
    throw new Error('XLSX fixture 自检失败：Sheet 或合并标题未写入。');
  }
}

await mkdir(imageDir, { recursive: true });
await writeFile(imagePath, imageBuffer);
await writeFile(mdPath, sampleMarkdown, 'utf8');
await writeFile(coverPath, coverMarkdown, 'utf8');
await writeFile(v4TablePath, v4TableMarkdown, 'utf8');
await writeFile(v4WideTablePath, v4WideTableMarkdown, 'utf8');
await buildDocx();
await buildReferenceDocx();
await buildManualHeadingDocx();
await buildLargeDocx();
await buildXlsx();

const templatesDir = path.join(toolsDir, '..', 'templates');
const productTemplatePath = path.join(templatesDir, 'reference-template.docx');
await mkdir(templatesDir, { recursive: true });
await copyFile(referenceDocxPath, productTemplatePath);

  const [mdStat, coverStat, v4TableStat, v4WideTableStat, docxStat, referenceDocxStat, manualHeadingDocxStat, largeDocxStat, xlsxStat, imageStat, templateStat] = await Promise.all([
  stat(mdPath),
  stat(coverPath),
  stat(v4TablePath),
  stat(v4WideTablePath),
  stat(docxPath),
  stat(referenceDocxPath),
  stat(manualHeadingDocxPath),
  stat(largeDocxPath),
  stat(xlsxPath),
  stat(imagePath),
  stat(productTemplatePath),
]);

console.log(JSON.stringify({
  fixtureDir,
  files: {
    md: { path: mdPath, bytes: mdStat.size },
    cover: { path: coverPath, bytes: coverStat.size },
    v4Table: { path: v4TablePath, bytes: v4TableStat.size },
    v4WideTable: { path: v4WideTablePath, bytes: v4WideTableStat.size },
    docx: { path: docxPath, bytes: docxStat.size },
    referenceDocx: { path: referenceDocxPath, bytes: referenceDocxStat.size },
    manualHeadingDocx: { path: manualHeadingDocxPath, bytes: manualHeadingDocxStat.size },
    largeDocx: { path: largeDocxPath, bytes: largeDocxStat.size },
    xlsx: { path: xlsxPath, bytes: xlsxStat.size },
    image: { path: imagePath, bytes: imageStat.size },
    template: { path: productTemplatePath, bytes: templateStat.size },
  },
}, null, 2));
