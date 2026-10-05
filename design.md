# FilesAsKM 系统架构与工程设计规范

## 1. 系统目标与运行约束

### 1.1 系统目标
针对企业技术文档（Word、PDF、PPT、Excel）与架构图纸（Draw.io、Visio），提取非结构化文本内容、结构化数据字典与矢量拓扑关系，构建系统级架构知识图谱，并提供基于事实证据索引的问答检索功能。

### 1.2 运行环境与约束
- **运行平台**：单机操作系统（Windows 11 / macOS）。
- **数据源存储**：OneDrive、Google Drive、Box 本地同步目录。
- **模型接口**：本地 Ollama 服务接口（`http://localhost:11434`），提供文本向量化与文本生成能力。
- **计算资源限制**：不部署常驻微服务集群（如独立的 Elasticsearch、MySQL、MinIO、分布式图数据库），使用嵌入式存储库。

---

## 2. 总体架构设计

系统由四个逻辑模块构成：

```
+-----------------------------------------------------------------------------+
|                          1. 网盘接入与文件过滤模块                          |
|  - 监控网盘本地同步目录 (OneDrive / Google Drive / Box)                    |
|  - 过滤临时文件与锁文件 (~$*, *.tmp, 冲突副本)                              |
|  - 文件写入状态检测 (排他锁检查与哈希稳定校验)                              |
+--------------------------------------|--------------------------------------+
                                       | 就绪文件路径
+--------------------------------------v--------------------------------------+
|                          2. 格式解析与拓扑逆向模块                          |
|  +---------------------------------+   +---------------------------------+  |
|  |     Draw.io / Visio 拓扑解析    |   |    Office / PDF / Excel 解析    |  |
|  | - Draw.io: XML/Deflate 解码     |   | - 文本提取与标题层级切片        |  |
|  |   提取 mxCell(vertex, edge)     |   | - Word/PDF: 表格转 Markdown     |  |
|  | - Visio: OpenXML 结构解压       |   | - Excel: Sheet/表头/数据字典    |  |
|  |   提取 Shape 与 Connect 标签    |   | - 提取页码/行列号/坐标          |  |
|  +----------------|----------------+   +----------------|----------------+  |
+-------------------|-------------------------------------|-------------------+
                    | 实体与拓扑关系                      | 文本切片与位置
+-------------------v-------------------------------------v-------------------+
|                        3. 嵌入式图谱与双向索引中枢                          |
|  - 架构 Schema 规范化与实体对齐                                             |
|  - 双向互引映射 (切片 <-> 图节点/边)                                        |
|  +-------------------------------+     +---------------------------------+  |
|  | Kùzu 嵌入式图数据库           | <-> | LanceDB 文件向量库              |  |
|  | - 节点、边、版本演进关系      |     | - 切片文本、向量、来源索引      |  |
|  +-------------------------------+     +---------------------------------+  |
+--------------------------------------|--------------------------------------+
                                       | 子图结构 + 证据切片
+--------------------------------------v--------------------------------------+
|                          4. 检索推理与本地交付模块                          |
|  - 查询意图解析与多跳图漫游检索                                             |
|  - 组装带证据编号的上下文 Prompt                                            |
|  - 调用本地 Ollama API 生成结论                                             |
|  - 本地 Web 页面展示回答、对应证据切片及 Mermaid 拓扑图                    |
+-----------------------------------------------------------------------------+
```

---

## 3. 领域模式定义（Domain Schema）

### 3.1 实体类型（Node Labels）

| 实体类型 | 说明 | 属性字段 |
| :--- | :--- | :--- |
| `System` | 顶层系统或平台 | `id`, `name`, `description`, `version`, `source_chunks` |
| `Service` | 服务或应用组件 | `id`, `name`, `type`, `version`, `source_chunks` |
| `Interface` | 接口规约（REST/RPC/MQ） | `id`, `name`, `protocol`, `path_or_topic`, `source_chunks` |
| `Database` | 数据库实例或表 | `id`, `name`, `engine`, `tables`, `source_chunks` |
| `DeploymentUnit` | 部署单元（容器、节点） | `id`, `name`, `platform`, `source_chunks` |
| `DocumentChunk` | 文档切片实体 | `chunk_id`, `doc_path`, `location_type`, `location_ref`, `content` |

### 3.2 关系类型（Edge Types）

| 关系类型 | 源实体 -> 目标实体 | 属性字段 |
| :--- | :--- | :--- |
| `CALLS` | `Service` -> `Service` / `Interface` | `protocol`, `sync_type`, `source_chunks` |
| `READS_FROM` | `Service` -> `Database` | `table_name`, `source_chunks` |
| `WRITES_TO` | `Service` -> `Database` | `table_name`, `source_chunks` |
| `DEPLOYED_AT` | `Service` -> `DeploymentUnit` | `environment`, `source_chunks` |
| `DEPENDS_ON` | 通用实体间依赖 | `dependency_type`, `source_chunks` |
| `EVOLVED_TO` | `Service` -> `Service`（版本演进） | `version_from`, `version_to`, `date`, `source_chunks` |
| `DEPRECATED_BY`| 旧实体 -> 新实体（替代关系） | `reason`, `date`, `source_chunks` |

### 3.3 双向互引数据结构（Mutual Indexing）
每个图元素与文档切片建立映射：
- **切片记录（LanceDB）**：
  ```json
  {
    "chunk_id": "chunk_doc1_p3_c1",
    "doc_path": "docs/architecture_v2.docx",
    "location": {
      "type": "page | sheet",
      "index": "3 | Sheet:接口清单",
      "range": [100.0, 200.0, 500.0, 320.0]
    },
    "heading": "3. 认证网关设计",
    "text": "Auth-Gateway 负责统一鉴权并路由请求至下游服务...",
    "vector": [0.012, -0.043, 0.089, "..."],
    "entity_ids": ["service_auth_gateway"],
    "relation_ids": ["rel_calls_user_service"]
  }
  ```
- **图节点与边记录（Kùzu）**：
  每个节点与边包含 `source_chunks` 字符串列表，存储关联的 `chunk_id`。

---

## 4. 模块实现设计

### 4.1 网盘接入模块 (`DriveWatcher`)
- **监控机制**：使用 Python `watchdog` 库监听目标目录的创建、修改与移动事件。
- **文件过滤规则**：
  - 忽略以 `~$` 开头的文件（Office 临时锁文件）。
  - 忽略后缀为 `.tmp`、`.crdownload` 的文件。
  - 忽略隐藏文件及隐藏目录（以 `.` 开头）。
- **文件就绪检测**：
  - 收到写入事件后延迟 1.5 秒。
  - 尝试以只读模式打开目标文件；若文件被占用则继续轮询。
  - 比对两次采样间隔的文件大小与哈希值，一致后进入处理流程。

### 4.2 架构图逆向解析模块 (`DiagramParser`)

#### A. Draw.io 解析 (`DrawioParser`)
1. 读取 `.drawio` 或包含图表数据的 `.xml` 文件。
2. 提取 `<diagram>` 标签内部数据。如果内容经压缩，进行 URL 解码、Base64 解码与 zlib/deflate 算法解压，获取 `mxGraphModel` XML 文本。
3. 解析 `<mxCell>` 元素：
   - 提取 `vertex="1"` 的节点：读取 `id`、清洗 HTML 标签后的 `value`、`style` 属性中的图元类型及 `mxGeometry` 坐标。
   - 提取 `edge="1"` 的连接线：读取 `id`、`source` 属性（源节点 ID）、`target` 属性（目标节点 ID）以及连线上的文本标注 `value`。
4. 输出有向图数据结构：
   ```json
   {
     "nodes": [{"id": "n1", "name": "API-Gateway", "type": "rectangle"}],
     "edges": [{"id": "e1", "source": "n1", "target": "n2", "label": "HTTP/REST"}]
   }
   ```

#### B. Visio 解析 (`VisioParser`)
1. 使用 Python 标准库 `zipfile` 解压 `.vsdx` 文件。
2. 读取 `visio/pages/pages.xml` 确认页面列表，依次读取 `visio/pages/page[N].xml`。
3. 解析 `<Shape>` 元素提取图元 ID、文本内容及几何位置。
4. 解析 `<Connect>` 标签：
   - 读取 `FromSheet` 确定连接线图元 ID。
   - 读取 `ToSheet` 确定起点与终点图元 ID。
5. 递归遍历组合形状（Group Shape），建立图元父子隶属关系。
6. 输出包含源、目标与属性的三元组列表。

### 4.3 文档与表格解析切片模块 (`DocumentParser`)
- **Word (`.docx`)**：使用 `python-docx` 读取段落层级与表格，表格序列化为 Markdown 语法字符串。
- **PDF (`.pdf`)**：使用 `PyMuPDF` 提取文本、页码与边界框坐标（Bounding Box）。
- **PPTX (`.pptx`)**：使用 `python-pptx` 提取每页 Slide 的文本框内容与备注。
- **Excel (`.xlsx` / `.xls`)**：
  - 使用 `openpyxl` 遍历工作表（Sheet）。
  - **结构化模式提取**：
    - 数据字典表（字段名、数据类型、主外键、约束、说明）映射为 `Database` 节点及属性；
    - 接口清单表（接口名称、请求路径、协议、调用方、被调用方）映射为 `Interface` 节点及 `CALLS` 关系；
    - 部署映射表（服务名、IP、端口、容器名、部署环境）映射为 `DeploymentUnit` 节点及 `DEPLOYED_AT` 关系。
  - **表格文本切片**：
    - 按逻辑行窗口（如每 20~30 行并携带表头）切分为 Markdown Table 格式文本块，存入 LanceDB 供向量及全文检索；
    - 记录来源元数据：`sheet_name`（工作表名）与 `row_range`（行号范围，如 `A1:F35`）。
- **切片规则**：基于滑动窗口划分切片（500 字符/切片，重叠 100 字符），记录所属章节标题与原文件页码/工作表位置。

### 4.4 存储管理模块 (`StorageEngine`)
- **图存储（Kùzu）**：
  - 数据以文件形式保存在本地目录（如 `./data/kuzu_db`）。
  - 通过 Cypher 语句建表（`CREATE NODE TABLE ...`, `CREATE REL TABLE ...`）。
  - 提供多跳路径查询接口（如 `MATCH (s:Service)-[r:CALLS*1..2]->(t) RETURN s, r, t`）。
- **向量存储（LanceDB）**：
  - 数据以文件形式保存在本地目录（如 `./data/lance_db`）。
  - 存储字段包括 `chunk_id`、`doc_path`、`text`、`vector`（向量数据）与元数据。
  - 提供近似最近邻（ANN）向量检索接口。

### 4.5 检索与推理模块 (`GraphRAGRetriever`)
1. **实体识别与意图规划**：接收自然语言查询，提取涉及的系统、服务或接口名称。
2. **图拓扑漫游**：以目标实体在 Kùzu 中执行 1 至 3 跳遍历，获取直接调用、依赖与演进链路。
3. **关联证据召回**：收集遍历节点和边中的 `source_chunks` 标识，在 LanceDB 中读取对应的原始文本切片。
4. **上下文格式化**：
   - 将子图结构转为三元组列表或 Mermaid 语法字符串。
   - 将文本切片按照 `[序号] 文件名(页码/工作表): 切片内容` 排列。
5. **模型调用**：
   - 向本地 Ollama 接口（`POST /api/chat`）发送请求。
   - System Prompt 设定：仅依据提供的子图与文本证据回答问题，每个结论后标明对应的证据序号，不存在证据的链路不得推断。

### 4.6 本地交付模块 (`LocalWebUI`)
- **服务载体**：基于 Python 构建轻量 Web 服务。
- **展示功能**：
  - 文本对话界面，支持流式输出。
  - 回答内证据序号点击交互：展开关联的原始文件切片内容、文档路径与页码。
  - 架构图渲染：将检索返回的子图转换为 Mermaid 格式并在前端渲染，高亮展示查询涉及的调用路径。

---

## 5. 项目工程结构

```
FilesAsKM/
├── config/
│   ├── config.yaml               # 基础配置 (网盘路径、Ollama接口配置)
│   └── schema.yaml               # 实体与关系类型定义
├── src/
│   ├── __init__.py
│   ├── drive/                    # 网盘目录监控与过滤
│   │   ├── watcher.py
│   │   └── filter.py
│   ├── parsers/                  # 格式解析与逆向
│   │   ├── base.py
│   │   ├── drawio_parser.py      # Draw.io XML 解析
│   │   ├── visio_parser.py       # Visio OpenXML 解析
│   │   ├── excel_parser.py       # Excel (.xlsx/.xls) 结构与切片解析
│   │   └── doc_parser.py         # Word / PDF / PPTX 解析
│   ├── storage/                  # 存储读写
│   │   ├── kuzu_store.py         # Kùzu 属性图操作
│   │   └── lancedb_store.py      # LanceDB 向量操作
│   ├── core/                     # 模式对齐与索引绑定
│   │   ├── schema_manager.py
│   │   └── mutual_indexer.py
│   ├── rag/                      # 检索与 Ollama 交互
│   │   ├── ollama_client.py      # Ollama API 封装
│   │   └── retriever.py          # 图漫游与证据组装
│   └── ui/                       # 本地页面展示
│       ├── app.py
│       └── templates/
├── data/                         # 本地数据持久化目录
│   ├── kuzu_db/
│   └── lance_db/
├── tests/                        # 单元测试
│   ├── test_drawio.py
│   ├── test_visio.py
│   └── test_excel.py
├── design.md                     # 本设计规范
├── requirements.txt              # 项目依赖库
└── README.md
```

---

## 6. 实施路线图

| 阶段 | 交付目标 |
| :--- | :--- |
| **阶段 1：解析模块与网盘监控** | 1. 实现网盘文件过滤与就绪检测逻辑；<br>2. 实现 Draw.io 与 Visio 的 XML 解析脚本，输出节点与连线拓扑；<br>3. 实现 Word、PDF、PPT 与 Excel 结构化表格及切片提取；<br>4. 连通本地 Ollama API 接口。 |
| **阶段 2：存储集成与数据入库** | 1. 初始化 Kùzu 图结构与 LanceDB 数据表；<br>2. 实现解析数据向图数据库与向量库的写入；<br>3. 建立并固化双向互引映射字段。 |
| **阶段 3：图检索与问答推理** | 1. 实现基于 Cypher 的多跳拓扑检索与切片反查；<br>2. 组装输入上下文，完成 Ollama 结构化回答生成。 |
| **阶段 4：界面集成与结果展示** | 1. 运行本地 WebUI；<br>2. 实现问答交互、证据切片查看与 Mermaid 拓扑图渲染。 |
