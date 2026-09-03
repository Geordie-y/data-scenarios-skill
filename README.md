# 数据应用场景材料 Skill

根据 PPT、Excel 数据目录、PDF、Word 等输入，生成固定七部分结构、可编辑的正式 Word 应用材料。

## 安装与使用

将本仓库链接交给 Codex：

```text
请安装此仓库 skills/artifact-template-data-scenarios 中的 skill，保留全部模板与脚本，检查运行依赖并执行自带验证。
```

私有仓库需要使用具有访问权限的 GitHub 账号。安装后提供业务文件并调用：

```text
使用 $artifact-template-data-scenarios，根据附件生成正式的数据应用场景材料。缺少的方案内容结合业务合理补全，沿用固定样式，交付可编辑 Word，来源与编制说明单独保存。
```

## 输出与编写规则

每批输入默认生成一份 Word。多个场景按行业归类，每个场景依次包括场景一览表、背景和痛点、场景方案、数据来源（数据提供方）、目标客户（场景买单方）、收益测算、相关案例。

正文采用正式业务表达；资料页码、内容依据和核验事项保存在独立说明。缺少的方案流程、客户定位、交付方式和商业模式按业务逻辑补全；规划参数通过工具计算。实际事实需来源支持，自拟应用示例不冒充已落地项目，收入与利润分开。

## 目录

- `skills/artifact-template-data-scenarios/SKILL.md`：技能入口。
- `assets/`：参考 Word、填充模板与预览。
- `references/`：内容契约、读取规则和样式证据。
- `scripts/`：文件提取、Word 生成、结构检查、渲染及测试。
- `requirements.txt`：Python 依赖。

后四项均位于上述 skill 文件夹内。完整安装该文件夹，不要只复制 SKILL.md。

## 运行环境

需要 Python 3 及 skill 文件夹中 requirements.txt 列出的依赖。渲染检查需要 LibreOffice、Poppler 和中文字体；扫描 PDF 自动 OCR 另外需要 Tesseract 及对应语言包。可使用已有的兼容运行环境。

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r skills/artifact-template-data-scenarios/requirements.txt
.venv/bin/python skills/artifact-template-data-scenarios/scripts/test_invariants.py
```

以上命令适用于 macOS/Linux。Windows 使用虚拟环境 Scripts 目录中的 Python。

内容编写由调用 skill 的模型完成，脚本负责提取、验证和稳定排版；脚本不是独立的模型服务。

## 已完成验证

2026-09-03 在 macOS 上完成 10 项脚本测试，并分别使用方案型 PPT 与字段目录型 Excel 生成正式 Word；两份成品共 8 页，已逐页检查。原始业务输入、生成材料和临时文件未收入本仓库。

部分中文字体在测试机器上使用替代字体渲染，DOCX 保留参考文件中的字体声明。不同机器的字体和分页可能存在差异。测试机器没有 Tesseract，扫描 PDF 只验证页图读取，没有宣称自动 OCR 通过。

## 模板资源

`assets/reference.docx` 保留原始排版样例的内容，只用作格式依据；`assets/template.docx` 用于填充。仓库按私有方式共享。未指定开源许可证。
