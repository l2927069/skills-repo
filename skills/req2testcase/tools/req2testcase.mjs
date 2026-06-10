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

// ====== 命令: metersphere ======
function cmdMeterSphere(jsonPath, outputPath) {
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

  // 字段映射（支持中英文 key）
  function getVal(tc, ...keys) {
    for (const k of keys) {
      const v = tc[k];
      if (v !== undefined && v !== null && v !== "") return String(v);
    }
    return "";
  }

  // MeterSphere 导入模板列头（步骤描述模式）
  const MS_HEADERS = [
    "用例名称",
    "用例等级",
    "用例类型",
    "模块",
    "前置条件",
    "步骤描述",
    "预期结果",
    "标签",
    "备注",
  ];

  const wsData = [MS_HEADERS];

  for (const tc of testcases) {
    // 步骤和预期结果处理：确保带序号
    const rawSteps = getVal(tc, "steps", "testSteps", "test_steps", "测试步骤");
    const rawExpected = getVal(tc, "expected", "expectedResult", "expected_result", "预期结果");

    // 将步骤和预期按行分割，加上序号
    const stepLines = rawSteps.split("\n").filter(Boolean);
    const expectedLines = rawExpected.split("\n").filter(Boolean);

    const numberedSteps = stepLines
      .map((s, i) => {
        // 如果已经有 "N. " 或 "N、 " 开头就不加序号
        if (/^\d+[.、．]/.test(s.trim())) return s.trim();
        return `${i + 1}. ${s.trim()}`;
      })
      .join("\n");

    const numberedExpected = expectedLines
      .map((e, i) => {
        if (/^\d+[.、．]/.test(e.trim())) return e.trim();
        return `${i + 1}. ${e.trim()}`;
      })
      .join("\n");

    // 模块路径处理：所属模块可能包含 "/" 用于 MeterSphere 多级模块
    const modulePath = getVal(tc, "module", "所属模块");

    // 标签：从用例类型 + 优先级生成
    const testType = getVal(tc, "type", "testType", "test_type", "用例类型");
    const priority = getVal(tc, "priority", "优先级");
    const tags = [testType, `P${priority.replace(/[Pp]/g, "")}`]
      .filter(Boolean)
      .join("、");

    // 备注：放用例编号
    const caseId = getVal(tc, "id", "caseId", "case_id", "用例编号");

    const row = [
      getVal(tc, "title", "caseName", "name", "用例名称"), // 用例名称
      priority.replace(/^P/i, "P"),                          // 用例等级
      testType,                                              // 用例类型
      modulePath,                                            // 模块
      getVal(tc, "precondition", "preCondition", "pre_condition", "前置条件"), // 前置条件
      numberedSteps,                                         // 步骤描述
      numberedExpected,                                      // 预期结果
      tags,                                                  // 标签
      caseId ? `原用例编号: ${caseId}` : "",                  // 备注
    ];
    wsData.push(row);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  const MS_COL_WIDTHS = [30, 10, 14, 20, 30, 40, 40, 20, 20];
  ws["!cols"] = MS_COL_WIDTHS.map((w) => ({ wch: w }));
  ws["!freeze"] = { x: 0, y: 1 };

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

  for (let c = 0; c < MS_HEADERS.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) ws[addr].s = headerStyle;
  }

  for (let r = 1; r < wsData.length; r++) {
    for (let c = 0; c < MS_HEADERS.length; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!ws[addr]) continue;
      ws[addr].s = cellStyle;
      // 优先级列特殊着色（第2列）
      if (c === 1) {
        const p = String(ws[addr].v).trim().toUpperCase();
        if (p === "P0") {
          ws[addr].s.font = { bold: true, color: { rgb: "FF0000" }, sz: 10, name: "微软雅黑" };
        } else if (p === "P1") {
          ws[addr].s.font = { color: { rgb: "FF8C00" }, sz: 10, name: "微软雅黑" };
        }
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "测试用例");
  XLSX.writeFile(wb, outputPath);
  console.log(`✅ MeterSphere 格式已生成: ${outputPath}`);
  console.log(`📊 共 ${testcases.length} 条用例`);
  console.log("💡 导入路径: 测试跟踪 → 功能用例 → 导入 → Excel → 步骤描述");
}

// ====== 命令: upload ======
async function cmdUpload(jsonPath, configPath) {
  if (!fs.existsSync(jsonPath)) {
    console.error(`测试用例文件不存在: ${jsonPath}`);
    process.exit(1);
  }
  if (!fs.existsSync(configPath)) {
    console.error(`配置文件不存在: ${configPath}`);
    process.exit(1);
  }

  const testcases = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
  const config = JSON.parse(fs.readFileSync(configPath, "utf-8"));

  if (!Array.isArray(testcases) || testcases.length === 0) {
    console.error("测试用例为空");
    process.exit(1);
  }
  if (!config.url) {
    console.error("配置缺少 url (MeterSphere 地址)");
    process.exit(1);
  }
  if (!config.projectId) {
    console.error("配置缺少 projectId");
    process.exit(1);
  }

  const baseUrl = config.url.replace(/\/+$/, "");
  const batchSize = config.batchSize || 10;
  let token;

  // ---- 登录 ----
  let csrfToken = "";
  let cookieJar = "";

  if (config.apiToken) {
    token = config.apiToken;
    console.log("🔑 使用 API Token 认证");
  } else if (config.auth) {
    console.log(`🔑 登录 ${baseUrl}/login ...`);

    try {
      const loginRes = await fetch(`${baseUrl}/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          username: config.auth.username,
          password: config.auth.password,
        }),
      });

      const text = await loginRes.text();

      if (!loginRes.ok) {
        console.error(`❌ 登录失败 (${loginRes.status})`);
        console.error("💡 检查用户名密码是否正确");
        process.exit(1);
      }

      const data = JSON.parse(text);
      if (data.code !== 100200 || !data.data) {
        console.error(`❌ 登录失败: ${JSON.stringify(data)}`);
        process.exit(1);
      }

      // X-AUTH-TOKEN 在响应头中
      token = loginRes.headers.get("X-AUTH-TOKEN") || "";
      // csrfToken 在响应体中
      csrfToken = data.data.csrfToken || data.data.token || "";
      const sessionId = data.data.sessionId || "";

      // 从 Set-Cookie 提取 session cookie
      const setCookie = loginRes.headers.get("Set-Cookie") || "";
      cookieJar = setCookie.split(";")[0]; // JSESSIONID=xxx

      if (!csrfToken && !sessionId) {
        console.error("❌ 登录返回无 csrfToken，可能版本不匹配");
        process.exit(1);
      }
      console.log(`✅ 登录成功 (用户: ${data.data.name}, session: ${sessionId.slice(0,8)}...)`);
    } catch (err) {
      console.error(`❌ 登录请求失败: ${err.message}`);
      process.exit(1);
    }
  } else {
    console.error("配置缺少 auth 或 apiToken");
    process.exit(1);
  }

  // 通用请求头（所有 API 调用都需要）
  const headers = {
    Accept: "application/json, text/plain, */*",
    "X-AUTH-TOKEN": token,
    "csrf-token": csrfToken,
  };
  if (config.organization) headers["organization"] = config.organization;
  if (config.projectId) headers["project"] = config.projectId;
  if (cookieJar) headers["Cookie"] = cookieJar;

  // 字段映射辅助函数
  function getVal(tc, ...keys) {
    for (const k of keys) {
      const v = tc[k];
      if (v !== undefined && v !== null && v !== "") return String(v);
    }
    return "";
  }

  // ---- 模块查找 ----
  async function ensureModule(modulePath) {
    // MeterSphere v3.x 模块 API 已变更，暂时直接用根模块
    return config.projectId;
  }

  // ---- 创建单条用例（multipart/form-data） ----
  async function createTestCase(tc) {
    const rawSteps = getVal(tc, "steps", "testSteps", "test_steps", "测试步骤");
    const rawExpected = getVal(tc, "expected", "expectedResult", "expected_result", "预期结果");

    const stepLines = rawSteps.split("\n").filter(Boolean);
    const expectedLines = rawExpected.split("\n").filter(Boolean);

    // 构造 MeterSphere 步骤格式（带 id）
    const steps = stepLines.map((s, i) => ({
      id: `step_${Date.now()}_${i}`,
      num: i,
      desc: s.replace(/^\d+[.、．]\s*/, ""),
      result: expectedLines[i]
        ? expectedLines[i].replace(/^\d+[.、．]\s*/, "")
        : "",
    }));

    const moduleId = await ensureModule(getVal(tc, "module", "所属模块"));
    const caseName = getVal(tc, "title", "caseName", "name", "用例名称");
    const priority = getVal(tc, "priority", "优先级").replace(/^P/i, "P") || "P3";

    // 构造完整的请求体（与浏览器捕获一致）
    const requestBody = {
      id: "",
      projectId: config.projectId,
      templateId: config.templateId || "",
      name: caseName,
      prerequisite: getVal(tc, "precondition", "preCondition", "pre_condition", "前置条件"),
      caseEditType: "STEP",
      steps: JSON.stringify(steps),
      textDescription: "",
      expectedResult: "",
      description: "",
      publicCase: false,
      moduleId: moduleId,
      versionId: "",
      tags: [getVal(tc, "type", "testType", "test_type", "用例类型") || "功能测试"],
      customFields: [
        { fieldId: config.priorityFieldId || "", value: priority },
      ],
      relateFileMetaIds: [],
      functionalPriority: "",
      reviewStatus: "UN_REVIEWED",
      caseDetailFileIds: [],
      aiCreate: false,
    };

    if (config.debug) {
      console.log("  请求体:", JSON.stringify(requestBody, null, 2));
    }

    // 构造 multipart/form-data（与浏览器一致）
    const ts = Date.now();
    const boundary = `----FormBoundary${ts}`;
    const multipartBody = [
      `------FormBoundary${ts}\r\nContent-Disposition: form-data; name="request"; filename="blob"\r\nContent-Type: application/json;charset=utf-8\r\n\r\n${JSON.stringify(requestBody)}\r\n`,
      `------FormBoundary${ts}--\r\n`,
    ].join("");

    let res;
    try {
      res = await fetch(`${baseUrl}/functional/case/add`, {
        method: "POST",
        headers: {
          ...headers,
          "Content-Type": `multipart/form-data; boundary=----FormBoundary${ts}`,
        },
        body: multipartBody,
      });
    } catch (err) {
      return { ok: false, name: caseName, status: 0, error: err.message };
    }

    if (res.ok) {
      const text = await res.text();
      try {
        const result = JSON.parse(text);
        return { ok: true, name: caseName, id: result.data?.id || result.id };
      } catch (e) {
        return { ok: false, name: caseName, status: res.status, error: "返回非 JSON: " + text.slice(0, 100) };
      }
    } else {
      const text = await res.text();
      return { ok: false, name: caseName, status: res.status, error: text.slice(0, 300) };
    }
  }

  // ---- 批量上传 ----
  console.log(`\n📤 开始上传 ${testcases.length} 条用例到 ${baseUrl} ...\n`);

  let success = 0,
    fail = 0;
  for (let i = 0; i < testcases.length; i += batchSize) {
    const batch = testcases.slice(i, i + batchSize);
    const results = await Promise.all(batch.map(createTestCase));

    for (const r of results) {
      if (r.ok) {
        success++;
      } else {
        fail++;
        console.error(`  ❌ [${r.name}] ${r.status}: ${(r.error || "").slice(0, 200)}`);
      }
    }

    // 批量进度
    const pct = Math.min(100, Math.round(((i + batch.length) / testcases.length) * 100));
    console.log(`  📊 进度: ${i + batch.length}/${testcases.length} (${pct}%)  ✅${success} ❌${fail}`);
  }

  console.log(`\n✅ 上传完成: 成功 ${success}, 失败 ${fail}`);
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
    case "metersphere": {
      if (!args[0] || !args[1]) {
        console.error("用法: node req2testcase.mjs metersphere <input.json> <output.xlsx>");
        process.exit(1);
      }
      cmdMeterSphere(args[0], args[1]);
      break;
    }
    case "upload": {
      if (!args[0] || !args[1]) {
        console.error("用法: node req2testcase.mjs upload <testcases.json> <config.json>");
        console.error("  config.json 包含: url, projectId, auth(username/password) 或 apiToken");
        process.exit(1);
      }
      await cmdUpload(args[0], args[1]);
      break;
    }
    default:
      console.error(`
用法:
  node req2testcase.mjs extract <input.docx>                      提取 Word 文本
  node req2testcase.mjs generate <input.json> <output.xlsx>       生成通用 Excel
  node req2testcase.mjs metersphere <input.json> <output.xlsx>    生成 MeterSphere 格式
  node req2testcase.mjs upload <testcases.json> <config.json>     上传到 MeterSphere

示例:
  node req2testcase.mjs extract 需求文档.docx > extracted.json
  node req2testcase.mjs metersphere testcases.json 测试用例.xlsx
  node req2testcase.mjs upload testcases.json ms-config.json
`);
      process.exit(1);
  }
}

main().catch((err) => {
  console.error("错误:", err.message);
  process.exit(1);
});
