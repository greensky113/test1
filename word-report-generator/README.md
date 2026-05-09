# 退网报告生成工具（纯前端版）

## 功能特性

- 纯前端运行，无需后端服务
- 支持 Word 模板占位符替换
- 支持按小区批量插入图片
- 支持压缩包导入多个小区的图片文件夹

## 所需文件

您需要准备以下三个文件：

### 1. Word 模板文件
- 格式：`.docx`
- 包含占位符：
  - 文本占位符：`{{字段名}}`
  - 图片占位符：`{{@RemovedEquipment_images}}`、`{{@Environment_images}}`

### 2. Excel 数据表格
- 格式：`.xlsx` 或 `.xls`
- 第一行是表头（字段名）
- 必须包含 `CommunityName` 字段用于匹配小区图片
- 从第二行开始是数据

### 3. 图片压缩包
- 格式：`.zip`
- 压缩包结构如下：

```
退网报告.zip
└── 退网报告/
    ├── 小区名称1/
    │   ├── RemovedEquipment_images/
    │   │   ├── photo1.jpg
    │   │   └── photo2.jpg
    │   └── Environment_images/
    │       ├── photo3.jpg
    │       └── photo4.jpg
    ├── 小区名称2/
    │   ├── RemovedEquipment_images/
    │   └── Environment_images/
    └── ...
```

## 快速使用

1. 双击 `退网报告生成小工具.html` 在浏览器中打开
2. 上传 Word 模板文件
3. 上传 Excel 数据表格（第一行为表头）
4. 上传图片压缩包（按上述结构组织）
5. 点击「开始生成报告」
6. 下载生成的报告压缩包

## 图片占位符说明

| 占位符 | 对应图片来源 |
|--------|-------------|
| `{{@RemovedEquipment_images}}` | 小区文件夹下 `RemovedEquipment_images/` 中的所有图片 |
| `{{@Environment_images}}` | 小区文件夹下 `Environment_images/` 中的所有图片 |

## 匹配逻辑

1. 读取 Excel 每行的 `CommunityName` 字段
2. 在压缩包中查找与 `CommunityName` 同名的小区文件夹（支持大小写不敏感）
3. 将该小区下对应子文件夹的图片插入到报告中

## 注意事项

- 所有操作在浏览器本地完成，数据不会上传到服务器
- 图片尺寸默认固定为 6cm × 6cm
- 建议使用 Chrome、Edge 等现代浏览器
