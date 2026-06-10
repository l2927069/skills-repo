#!/usr/bin/env node
/**
 * req2testcase.mjs — 需求文档转测试用例工具
 *
 * 用法:
 *   node req2testcase.mjs extract <input.docx>       提取 Word 文本 (输出 JSON)
 *   node req2testcase.mjs generate <input.json> <output.xlsx>  从 JSON 生成 Excel
 *
 * 依赖:
 *   npm install mammoth xlsx
 */

import mammoth from "mammoth";
import * as XLSX from "xlsx";
import fs from "node:fs";
import path from "node:path";

// ====== 配置 ======
const HEADERS = [
  "用例编号",
  "所属模块",
  "用例名称",
  "优先级",
  "前置条件",
  "测试步骤",
  "预期结果",
  "用例类型",
];

const COL_WIDTHS = [16, 16, 30, 10, 30, 40, 40, 14];

// ====== Excel 样式 ======
function buildXBorders() {
  const thin = { style: "thin", color: { rgb: "999999" } };
  return {
    top: thin,
    bottom: thin,
    left: thin,
    right: thin,
  };
}

// ====== 命令: extract ======
async function cmdExtract(docxPath) {
  if (!fs.existsSync(docxPath)) {
    console.error(`文件不存在: ${docxPath}`);
    process.exit(1);
  }

  const buffer = fs.readFileSync(docxPath);

  // 用 mammoth 提取 HTML 结构文本
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value;

  // 同时提取纯文本（用于 AI 分析）
  const rawText = await mammoth.extractRawText({ buffer });

  // 尝试提取标题层级
  const titlePattern = /<h(\d)[^>]*>(.*?)<\/h\d>/gi;
  const headings = [];
  let m;
  while ((m = titlePattern.exec(html)) !== null) {
    headings.push({ level: parseInt(m[1]), text: m[2].replace(/<[^>]+>/g, "") });
  }

  // 提取表格数据
  const tablePattern = /<table[^>]*>.*?<\/table>/gis;
  const tables = [];
  let tableMatch;
  while ((tableMatch = tablePattern.exec(html)) !== null) {
    const tableHtml = tableMatch[0];
    const rows = [];
    const rowPattern = /<tr[^>]*>.*?<\/tr>/gis;
    let rowMatch;
    while ((rowMatch = rowPattern.exec(tableHtml)) !== null) {
      const cells = [];
      const cellPattern = /<t[dh][^>]*>(.*?)<\/t[dh]>/gis;
      let cellMatch;
      while ((cellMatch = cellPattern.exec(rowMatch[0])) !== null) {
        cells.push(cellMatch[1].replace(/<[^>]+>/g, "").trim());
      }
      if (cells.length > 0) rows.push(cells);
    }
    if (rows.length > 0) tables.push(rows);
  }

  const output = {
    title: path.basename(docxPath, ".docx"),
    headings,
    tables,
    rawText: rawText.value.trim(),
    html,
    warnings: result.messages,
  };

  console.log(JSON.stringify(output, null, 2));
}

// ====== 命令: generate ======
function cmdGenerate(jsonPath, outputPath) {
  if (!fs.existsSync(jsonPath)) {
    console.error(`文件不存在: ${jsonPath}`);
    process.exit(1);
  }

  const jsonRaw = fs.readFileSync(jsonPath, "utf-8");
  const testcases = JSON.parse(jsonRaw);
  if (!Array.isArray(testcases)) {
    console.error("JSON 必须是测试用例数组");
    process.exit(1);
  }

  // 映射字段名（支持中英文 key）
  const fieldMap = {
    id: "用例编号",
    caseId: "用例编号",
    case_id: "用例编号",
    用例编号: "用例编号",
    module: "所属模块",
    所属模块: "所属模块",
    title: "用例名称",
    caseName: "用例名称",
    name: "用例名称",
    用例名称: "用例名称",
    priority: "优先级",
    优先级: "优先级",
    precondition: "前置条件",
    preCondition: "前置条件",
    pre_condition: "前置条件",
    前置条件: "前置条件",
    steps: "测试步骤",
    testSteps: "测试步骤",
    test_steps: "测试步骤",
    测试步骤: "测试步骤",
    expected: "预期结果",
    expectedResult: "预期结果",
    expected_result: "预期结果",
    预期结果: "预期结果",
    type: "用例类型",
    testType: "用例类型",
    test_type: "用例类型",
    用例类型: "用例类型",
  };

  const wsData = [HEADERS];
  for (const tc of testcases) {
    const row = HEADERS.map((header) => {
      // 查找映射中的值
      for (const [key, val] of Object.entries(fieldMap)) {
        if (val === header && tc[key] !== undefined) {
          return String(tc[key]);
        }
      }
      return "";
    });
    wsData.push(row);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // 列宽
  ws["!cols"] = COL_WIDTHS.map((w) => ({ wch: w }));

  // 冻结首行
  ws["!freeze"] = { x: 0, y: 1 };

  // ---- 样式 ----
  // 注意: xlsx-style 扩展不在标准库中，我们用标准 API 能达到的最佳效果
  const headerStyle = {
    font: { name: "微软雅黑", bold: true, sz: 11, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "4472C4" }, patternType: "solid" },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: buildXBorders(),
  };
  const cellStyle = {
    font: { name: "微软雅黑", sz: 10 },
    alignment: { vertical: "top", wrapText: true },
    border: buildXBorders(),
  };

  // 应用表头样式
  for (let c = 0; c < HEADERS.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (!ws[addr]) continue;
    ws[addr].s = headerStyle;
  }

  // 应用数据行样式
  for (let r = 1; r < wsData.length; r++) {
    for (let c = 0; c < HEADERS.length; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!ws[addr]) continue;
      const cell = ws[addr];
      cell.s = cellStyle;
      // 优先级列特殊着色
      if (c === 3) {
        const priority = String(cell.v).trim().toUpperCase();
        if (priority === "P0") {
          cell.s.font = { bold: true, color: { rgb: "FF0000" }, sz: 10, name: "微软雅黑" };
        } else if (priority === "P1") {
          cell.s.font = { color: { rgb: "FF8C00" }, sz: 10, name: "微软雅黑" };
        }
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "测试用例");
  XLSX.writeFile(wb, outputPath);
  console.log(`✅ 测试用例已生成: ${outputPath}`);
  console.log(`📊 共 ${testcases.length} 条用例`);
}

// ====== 主入口 ======
async function main() {
  const [, , command, ...args] = process.argv;

  switch (command) {
    case "extract": {
      if (!args[0]) {
        console.error("用法: node req2testcase.mjs extract <input.docx>");
        process.exit(1);
      }
      await cmdExtract(args[0]);
      break;
    }
    case "generate": {
      if (!args[0] || !args[1]) {
        console.error("用法: node req2testcase.mjs generate <input.json> <output.xlsx>");
        process.exit(1);
      }
      cmdGenerate(args[0], args[1]);
      break;
    }
    default:
      console.error(`
用法:
  node req2testcase.mjs extract <input.docx>       提取 Word 文本
  node req2testcase.mjs generate <input.json> <output.xlsx>  生成 Excel

示例:
  node req2testcase.mjs extract 需求文档.docx > extracted.json
  node req2testcase.mjs generate testcases.json 测试用例.xlsx
`);
      process.exit(1);
  }
}

main().catch((err) => {
  console.error("错误:", err.message);
  process.exit(1);
});
