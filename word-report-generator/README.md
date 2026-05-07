# 文件夹结构

## 项目结构

```
word-report-generator/
├── index.html          # 主页面（可直接在浏览器中打开）
├── 使用说明.md          # 详细使用指南
└── README.md           # 项目说明
```

## 所需文件

您需要准备以下三个文件：

### 1. Word模板文件
- 文件名：`附件4.退网方案模板.docx`（示例）
- 格式：`.docx`
- 包含占位符：`{{字段名}}` 和 `{{@图片字段名}}`

### 2. Excel数据表格
- 文件名：`退网报告填写字段.xlsx`（示例）
- 格式：`.xlsx` 或 `.xls`
- 第一行是表头（字段名）
- 从第二行开始是数据

### 3. 图片文件夹压缩包
- 文件夹主目录：`images/`
- 压缩为：`images.zip`
- 内部包含多个图片文件夹

## 示例结构

```
项目文件夹/
│
├── index.html                          # 生成工具
├── 使用说明.md                          # 使用指南
│
├── 附件4.退网方案模板.docx              # Word模板文件
├── 退网报告填写字段.xlsx                # Excel数据表格
│
└── images.zip                          # 图片文件夹压缩包
    │
    ├── photos_project001/
    │   ├── 01_entrance.jpg
    │   ├── 02_building.jpg
    │   └── 03_facility.jpg
    │
    ├── photos_project002/
    │   ├── 01_overview.jpg
    │   └── 02_detail.jpg
    │
    └── signatures/
        ├── sig_zhangsan.png
        └── sig_lisi.png
```

## 快速使用

1. 双击 `index.html` 在浏览器中打开
2. 上传 `附件4.退网方案模板.docx`
3. 上传 `退网报告填写字段.xlsx`（第一行说明，第二行表头，第三行开始数据）
4. 上传 `images.zip`
5. 点击「开始生成报告」
6. 下载生成的报告

## 占位符命名对应

Excel表格中的列名必须与Word模板中的占位符名称对应：

| Word模板占位符 | Excel列名 |
|---------------|----------|
| `{{report_no}}` | report_no |
| `{{project_name}}` | project_name |
| `{{@site_photos}}` | site_photos |
| `{{@signature}}` | signature |

**注意：** 列名中的 `{{}}` 和 `@` 符号不需要写在Excel列名中！