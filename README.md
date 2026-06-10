# Req2TestCase — 需求文档转测试用例 Skill

> 读取 Word (.docx) 需求文档，AI 自动分析并生成结构化 Excel 测试用例。

## 功能

- 📄 解析任意格式的 Word 需求文档（自由文本、表格、标题层级）
- 🧠 AI 自动识别功能模块、业务规则、边界条件
- 📊 生成格式化的 Excel 测试用例（.xlsx）
- 🎨 蓝色表头、冻结首行、P0/P1 颜色标记、自动筛选器
- ✅ 覆盖功能测试、边界测试、异常测试、业务规则测试、安全测试

## 安装

### 前置条件

- [Node.js](https://nodejs.org/) >= 18
- Claude Code

### 安装 Skill

在 Claude Code 中执行：

```bash
/install-skill github:你的用户名/skills-repo skills/req2testcase/SKILL.md
```

### 安装依赖

```bash
cd .claude/skills/req2testcase/tools && npm install
```

## 使用方法

在 Claude Code 中说：

```
读取 D:\需求文档.docx，生成测试用例
帮我分析这份 Word 需求文档并生成测试用例
```

Claude 会自动执行三步：

1. **提取文本** — 读取 .docx 文件，提取标题、表格、正文
2. **AI 分析** — 识别功能模块、业务规则，生成测试用例
3. **生成 Excel** — 输出格式化的 .xlsx 文件

## 输出字段

| 列名 | 说明 | 示例 |
|------|------|------|
| 用例编号 | TC-{模块缩写}-{序号} | TC-LOGIN-001 |
| 所属模块 | 对应的功能模块 | 登录模块 |
| 用例名称 | 简明描述测试场景 | 验证正确用户名密码登录成功 |
| 优先级 | P0/P1/P2/P3 | P0 |
| 前置条件 | 测试前必须满足的条件 | 用户已注册，账号已激活 |
| 测试步骤 | 操作的详细步骤 | 1. 打开登录页 2. 输入用户名 3. 点击登录 |
| 预期结果 | 期望的系统行为 | 登录成功，跳转首页 |
| 用例类型 | 功能测试/边界测试/异常测试/性能测试/安全测试 | 功能测试 |

## 测试类型覆盖

- **功能测试** — 正常流程、分支流程
- **边界测试** — 输入上限/下限、空值、特殊字符
- **异常测试** — 错误输入、网络中断、超时、权限不足
- **业务规则测试** — 条件组合、状态流转
- **安全测试** — 鉴权绕过、SQL注入、XSS（如适用）

## 优先级定义

| 级别 | 定义 |
|------|------|
| P0 | 核心功能，不通过则阻断发布 |
| P1 | 重要功能，建议修复后发布 |
| P2 | 次要功能，可延期修复 |
| P3 | 边缘场景 / 体验优化 |

## 仓库结构

```
skills/
└── req2testcase/
    ├── SKILL.md             # Skill 定义（触发词 + 工作流程）
    └── tools/
        ├── req2testcase.mjs # 工具脚本（提取 docx + 生成 xlsx）
        ├── package.json     # npm 依赖声明
        └── package-lock.json
```

## Python 替代方案

如需使用 Python 而非 Node.js：

```bash
pip install python-docx openpyxl
```

然后在 SKILL.md 中将 `node tools/req2testcase.mjs` 替换为对应的 Python 调用即可（工具逻辑相同）。

## License

MIT
