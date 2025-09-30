# 测试用例生成器

基于需求文档自动生成结构化测试用例的工具，支持从Word文档到XMind思维导图的完整转换流程。

## 功能特性

- 🚀 **智能分析**：基于需求文档自动分析核心功能、业务流程和关键风险点
- 📝 **结构化输出**：生成符合Tab缩进格式的测试用例，适配XMind思维导图
- 🧠 **思维导图**：自动转换为XMind格式，支持逻辑图展示
- 🔄 **多格式支持**：支持Word文档转Markdown，图片自动提取和引用
- 🎯 **高覆盖率**：覆盖正常流程、边界条件、异常情况和权限验证
- 📱 **标准化流程**：统一的提示词模板，确保测试用例质量一致性

## 项目结构

```
testcase-generator/
├── docs/                          # 文档目录
│   ├── input/                     # 输入的docx文件
│   └── markdown/                  # 转换后的markdown文件
├── prompts/                       # 提示词模板目录
│   └── testcase_generation.md    # 测试用例生成提示词
├── output/                        # 输出目录
│   ├── txt-case/                # 生成的txt测试用例
│   └── xmind-case/               # 生成的xmind文件
├── utils/                         # 工具脚本
│   ├── docx2md.py                # docx转markdown工具
│   └── convert_to_xmind.py       # txt转xmind工具
├── requirements.txt               # 依赖包列表
└── README.md                      # 项目说明文档
```
## 快速开始

### 1. 准备需求文档

将Word格式的需求文档放入 `docs/input/` 目录：

```bash
# 将需求文档复制到input目录
cp 你的需求文档.docx docs/input/
```

### 2. 使用AI助手生成测试用例

在Cursor中：

1. 打开 `prompts/testcase_generation.md` 文件
2. 使用 `@testcase_generation.md 按照指令生成测试用例` 命令
3. AI助手会自动：
   - 转换docx为markdown格式
   - 分析需求文档内容
   - 生成结构化测试用例
   - 输出txt和xmind格式文件

### 3. 查看结果

生成的文件位置：
- **Markdown文档**：`docs/markdown/文档名.md`
- **测试用例文本**：`output/testcases/文档名.txt`
- **XMind思维导图**：`output/xmind/文档名.xmind`


**注意**：本工具专为QA团队设计，用于提高测试用例生成效率。生成的测试用例建议经过人工审查后再使用。