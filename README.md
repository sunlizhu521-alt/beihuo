# 优乐步对账框架

这是从 `gysxx/youlebu-duizhang` 分析并迁移到 `beihuo` 仓库的静态 GitHub Pages 框架。

## 页面结构

- `index.html`: 对账工作台入口，包含文件槽位、规则说明、统计指标和生成按钮。
- `styles.css`: 侧边栏、指标卡、文件槽位和响应式布局样式。
- `file-library.js`: 文件上传、IndexedDB 缓存、对账生成、Excel 保真导出逻辑。
- `shared-library.js`: 从 `data/shared-library.json` 导入共享文件包。
- `data/shared-library.json`: 当前为空共享包，后续可发布已应用文件数据。
- `vendor/`: SheetJS 和 JSZip，本地加载以便 GitHub Pages 直接运行。

## 核心业务逻辑

- 槽位 1 上传运营登记表。
- 槽位 2 上传优乐步对账表，要求 `.xlsx` 或 `.xlsm`。
- 生成时读取优乐步对账表第 3 行开始的数据。
- M 列作为姓名，N 列作为快递单号。
- 在运营登记表全工作簿内匹配姓名和快递单号。
- Q 列写核对结果，R 列写命中备注。
- A-N 列原格式、边框、列宽、筛选、合并单元格等尽量保持不变。

## 后续使用

把仓库发布到 GitHub Pages 后，访问 `https://sunlizhu521-alt.github.io/beihuo/` 即可使用。
