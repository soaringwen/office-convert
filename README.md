# Office-Convert：Office 文档深度解析与 Markdown 转换服务

基于 **Java Spring Boot + Apache POI + PDFBox** 的实现，配套可选的 **Python Sidecar（MarkItDown + oletools）** 兜底转换与 OLE 分析能力。

对应需求文档：`Office 文档及嵌入附件转换为 Markdown 需求说明书`（以下按"需求 X.Y"引用条款）。

## 1. 功能总览（对照 P0/P1 需求）

| 能力 | 说明 |
|---|---|
| DOCX/XLSX/PPTX 转 Markdown | Apache POI 原生实现（P0） |
| 旧版 DOC/XLS/PPT | poi-scratchpad 实现（业务确认：首期必须支持） |
| PDF 转 Markdown | PDFBox，按页输出，提取 PDF 内嵌附件 |
| 嵌入附件识别与提取 | OOXML `/embeddings/` 部件 + OLE2 `\u0001Ole10Native`/`CONTENTS` 流（需求 7.6） |
| 附件递归解析 | 主文档 → 附件 → 子附件，manifest 表达父子关系（需求 3.4、7.7） |
| 部分成功机制 | 单附件失败不阻断主文档，任务状态 `partial_success`（需求 6、8） |
| 清单与报告 | 每个文档目录输出 `manifest.json` + `conversion-report.json`（需求 7.12、7.13） |
| 资源限制 | 20MB 单文件、10 个附件、2 层递归、任务超时、附件计数、大小限制（业务确认 + 需求 7.7） |
| 安全 | 宏禁用（仅静态检测，不执行）、可执行文件拦截、路径穿越防护、错误信息脱敏（需求 10） |
| Word 细节 | 标题层级/列表/加粗斜体删除线/超链接/表格（复杂合并降级 HTML）/图片/页眉页脚/批注/脚注（需求 7.3） |
| Excel 细节 | 多工作表/隐藏表/公式+计算值双输出/合并单元格/批注/超链接/错误值/冻结窗格/命名区域/大表分块+CSV（需求 7.4） |
| PPT 细节 | 按页输出/阅读顺序排序/表格/图片/备注/隐藏页/图表降级说明（需求 7.5） |
| 幂等与恢复 | 哈希幂等、重启恢复、失败重试（需求 7.1、11.2、12.4） |
| 外部链接 | 仅记录链接文字与地址，标记 `not_embedded`，不下载（需求 7.9） |

## 2. 目录结构

```
office-convert/
├── pom.xml                          # Spring Boot 3.3 + POI 5.2.5 + PDFBox 2.0.31
├── src/main/java/com/example/officeconvert/
│   ├── api/                         # REST 接口 + 全局异常处理
│   ├── config/                      # 全局配置（需求 13）
│   ├── core/                        # 任务管理（并发/超时/取消/重试/恢复/幂等）
│   ├── convert/                     # 可插拔转换器（需求 11.3）
│   ├── detect/                      # 文件类型识别（魔数 + 容器结构，需求 7.2）
│   ├── extract/                     # OOXML / OLE 嵌入附件提取
│   ├── model/                       # Manifest / Report / 状态枚举
│   ├── pipeline/                    # 递归转换编排
│   └── util/                        # 文件名安全 / 哈希 / Markdown 工具
├── python-sidecar/                  # MarkItDown + oletools 兜底服务（可选）
│   ├── app.py
│   └── requirements.txt
└── data/tasks/<taskId>/             # 运行时产物
    ├── input/                       # 上传原件副本
    ├── task.json                    # 任务状态（重启恢复）
    └── output/01-<name>/
        ├── content.md               # 主 Markdown
        ├── manifest.json            # 清单（需求 7.12）
        ├── conversion-report.json   # 报告（需求 7.13）
        ├── original/source.docx     # 原始文件
        ├── assets/image-001.png     # 图片等资源
        └── attachments/
            ├── att-001/original.xlsx + content.md + manifest.json + assets/
            ├── att-002/original.pdf + manifest.json
            └── overflow/            # 超过附件数量上限的原始附件（保留不解析）
```

## 3. 构建与运行

### 3.1 Java 主服务

要求 JDK 17+。若本机无 Maven，可用仓库内 `tools/apache-maven-*/bin/mvn`（见 3.3）。

```bash
cd office-convert
mvn spring-boot:run          # 或 mvn package && java -jar target/office-convert-1.0.0.jar
```

服务默认监听 `http://127.0.0.1:8080`。

### 3.2 Python Sidecar（可选，MarkItDown + oletools）

```bash
cd office-convert/python-sidecar
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --host 127.0.0.1 --port 8900
```

然后在 `application.yml` 中开启：

```yaml
office-convert:
  sidecar:
    enabled: true
    base-url: http://127.0.0.1:8900
```

启用后，`.rtf / .odt / .ods / .odp / .xlsb` 等格式由 MarkItDown 兜底转换；
sidecar 未启用时这些格式明确返回 `unsupported` 状态（不静默丢失）。

### 3.3 本地无 Maven 时

```bash
# 仓库已附带 tools/apache-maven-3.9.9（若已下载）
office-convert/tools/apache-maven-3.9.9/bin/mvn -f office-convert/pom.xml spring-boot:run
```

## 4. 接口（需求 12）

| 方法 | 路径 | 说明 |
|---|---|---|
| POST | `/api/v1/conversion-tasks` | 创建任务（multipart `files` 多文件；选项走 query/form 参数） |
| GET | `/api/v1/conversion-tasks` | 任务列表 |
| GET | `/api/v1/conversion-tasks/{taskId}` | 任务状态/进度/阶段 |
| GET | `/api/v1/conversion-tasks/{taskId}/result` | 结果索引（各文档 markdown/manifest/report 路径） |
| GET | `/api/v1/conversion-tasks/{taskId}/result/markdown` | 首文档 Markdown 文本 |
| GET | `/api/v1/conversion-tasks/{taskId}/result/zip` | 结果 ZIP 包 |
| GET | `/api/v1/conversion-tasks/{taskId}/files/**` | 产物/附件原件下载（含穿越防护） |
| POST | `/api/v1/conversion-tasks/{taskId}/cancel` | 取消 |
| POST | `/api/v1/conversion-tasks/{taskId}/retry` | 重试（可带新配置） |

### 4.1 创建任务示例

```bash
curl -F "files=@/path/项目方案.docx" \
     "http://127.0.0.1:8080/api/v1/conversion-tasks?maxDepth=2&formulaMode=both&includeComments=true"
```

可选参数（需求 12.1）：`markdownDialect`、`parseAttachments`、`maxDepth`、`enableOcr`（接受但忽略，业务确认不需要 OCR）、`includeHeaderFooter`、`includeComments`、`includeFootnotes`、`revisionStrategy`、`formulaMode`（value/formula/both）、`includeHiddenSheets`、`includeHiddenSlides`、`notesStrategy`。

### 4.2 查询与下载

```bash
curl http://127.0.0.1:8080/api/v1/conversion-tasks/task-20260911-xxxxxxxx
curl http://127.0.0.1:8080/api/v1/conversion-tasks/task-20260911-xxxxxxxx/result/zip -o result.zip
```

## 5. 状态定义（需求 8）

- 任务：`pending / processing / success / partial_success / failed / cancelled`
- 附件提取：`success / failed / skipped`
- 附件转换：`success / failed / unsupported / encrypted / corrupted / limit_exceeded / security_blocked / not_embedded`

## 6. 安全设计（需求 10）

- 宏文档（docm/xlsm/pptm 及携带 `vbaProject.bin` 的文档）：仅输出内容 + 告警，宏永不执行；
- `MZ`/ELF 可执行附件：`security_blocked`，不转换；
- 不自动破解密码；加密文件返回 `encrypted` 与明确原因（不暴露细节）；
- 压缩包按业务确认**不解析**，统一按 `unsupported` 保留原件并说明；
- 产物/接口下载路径规范化 + 前缀校验，防路径穿越；
- 全局异常仅返回通用错误信息，不含服务器路径与调用栈；
- 日志不记录正文内容，错误信息截断 200 字符。

## 7. 已知边界（P2 / 后续扩展）

- Word 文本框、目录（TOC）字段、PPT SmartArt 与 Excel 图表渲染为图片：以占位说明 + 降级记录表达（报告可追溯）；
- 旧版 .doc/.ppt 为文本近似转换（表格/图片尽力提取）；
- 附件数量超限时，超限附件原件保留于 `attachments/overflow/`，不做解析；
- OCR、音视频转写、云端处理：按业务确认不启用；
- PDF 加密文件不解密（可通过扩展接入用户补充密码重试）。

## 8. 验证清单（对照需求 14）

- [x] Word 标题/段落/列表/表格/图片/链接转换
- [x] Excel 工作表/数据区域/公式与计算值/合并单元格按配置输出
- [x] PPT 页序/标题/正文/表格/图片/备注
- [x] UTF-8 输出、相对路径资源链接
- [x] 三类 Office 文档 OLE/Package 附件提取 + 递归转换
- [x] 附件重名不覆盖（Deduper）、重复附件去重（sha256）
- [x] manifest 父子关系、单附件失败不阻断
- [x] 加密/损坏/超限文件明确状态
- [x] 深度/数量/大小/超时限制
