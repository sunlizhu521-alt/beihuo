# 备货需求分配框架

这是迁移到 `beihuo` 仓库的静态 GitHub Pages 框架，当前项目名称已调整为备货需求分配。

## 页面结构

- `index.html`: 备货需求分配工作台入口，包含文件槽位、规则说明、统计指标和生成按钮。
- `styles.css`: 侧边栏、指标卡、文件槽位和响应式布局样式。
- `file-library.js`: 文件上传、IndexedDB 缓存、备货需求分配生成、Excel 保真导出逻辑。
- `shared-library.js`: 从 `data/shared-library.json` 导入共享文件包。
- `data/shared-library.json`: 当前为空共享包，后续可发布已应用文件数据。
- `vendor/`: SheetJS 和 JSZip，本地加载以便 GitHub Pages 直接运行。

## 核心业务逻辑

- 4 个文件槽位都作为普通表格槽位使用，支持 `.xlsx`、`.xls`、`.xlsm`、`.csv`。
- 文件槽位依次为 `Dim-YL医疗器械商品分类`、`Dim-采购部分工明细`、`Dim-花名册（姓名&一级部门）`、`备货需求表`。
- 点击“一键刷新应用”后，页面会读取所有已应用表格。
- 指标卡显示备货总数和供应商数量。
- 筛选器包括事业部、销售产品线、采购组、采购订单下单人、供应商简称。
- 明细表字段包括采购组、采购单订单下单人、申请人、供应商简称、OA备货流程号、物料编码、SKU、物料名称、数量。
- 表格字段按表头名称自动匹配，匹配不到的列会留空。
- 明细表支持下载为 Excel，并按采购单订单下单人拆分为多个 sheet，sheet 名称来自采购单订单下单人。

## 后续使用

把仓库发布到 GitHub Pages 后，访问 `https://sunlizhu521-alt.github.io/beihuo/` 即可使用。
