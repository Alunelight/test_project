# PDF文件处理工具集

本项目提供了一系列Python脚本，用于处理PDF文件与Excel文件的匹配、重命名和移动操作。

## 功能概述

项目包含三个主要脚本：

1. **main.py** - PDF文件重命名脚本（按合同编号匹配）
2. **match_pdfs.py** - PDF文件匹配复制脚本（按身份证号匹配）
3. **match_pdfs_by_name.py** - PDF文件匹配移动脚本（按姓名匹配）

## 环境要求

- Python >= 3.13
- [uv](https://github.com/astral-sh/uv) - Python 包管理器（推荐）
- 依赖包：

  - pandas >= 2.0.0
  - xlrd >= 2.0.0
  - openpyxl >= 3.0.0
  - pandas-stubs >= 2.0.0

## 安装依赖

使用 `uv` 管理依赖：

```bash
uv sync
```

## 代码质量检查

项目使用 `ruff` 进行代码质量检查和格式化：

```bash
# 检查代码质量
uv tool run ruff check .

# 检查代码格式
uv tool run ruff format --check .

# 自动修复格式问题
uv tool run ruff format .
```

## 脚本说明

### 1. main.py - PDF文件重命名脚本

根据Excel文件中的合同编号信息，将PDF文件从"合同编号"格式重命名为"姓名+身份证号"格式。

**功能：**

- 读取目标文件夹下的所有PDF文件
- 识别PDF文件名格式：`协商解除劳动合同协议书_合同编号.pdf`
- 从Excel文件中查找对应的姓名和身份证号
- 重命名为：`协商解除劳动合同协议书_姓名身份证号.pdf`

**使用方法：**

```bash
uv run python main.py <目标文件夹路径> [--excel <Excel文件名>]
```

**示例：**

```bash
uv run python main.py /path/to/folder --excel "协商解除函签署名单-608人.xls"
```

**Excel文件要求：**

- 第一行为表头
- 必须包含"合同编号"、"姓名"、"身份证号"列

**PDF文件名格式：**

- 输入格式：`协商解除劳动合同协议书_4008070793657015304.pdf`
- 输出格式：`协商解除劳动合同协议书_张三110101199001011234.pdf`

---

### 2. match_pdfs.py - PDF文件匹配复制脚本（按身份证号）

根据Excel文件中的身份证号匹配PDF文件，并将匹配的文件复制到"匹配结果"文件夹。

**功能：**

- 读取PDF文件夹下的所有PDF文件
- 从PDF文件名中提取身份证号（支持末位为X）
- 在Excel文件中查找匹配的身份证号
- 将匹配的PDF文件复制到"匹配结果"文件夹
- 在Excel中标注匹配状态（成功/失败）

**使用方法：**

```bash
uv run python match_pdfs.py <PDF文件夹路径> <Excel文件路径> [--output-dir <输出文件夹名>]
```

**示例：**

```bash
uv run python match_pdfs.py /path/to/pdfs /path/to/excel.xlsx --output-dir "匹配结果"
```

**Excel文件要求：**

- 第一行为表头
- 必须包含"身份证号"列

**PDF文件名格式：**

- `协商解除劳动合同协议书_姓名身份证号.pdf`
- 身份证号支持18位数字，最后一位可能是0-9或X

**输出：**

- 匹配的文件复制到"匹配结果"文件夹
- Excel文件中添加"匹配状态"列，标注"成功"或"失败"

---

### 3. match_pdfs_by_name.py - PDF文件匹配移动脚本（按姓名）

根据Excel文件中的姓名匹配PDF文件，并将匹配的文件移动到"匹配结果"文件夹。

**功能：**

- 读取PDF文件夹下的所有PDF文件
- 从PDF文件名中提取员工姓名
- 在Excel文件中查找匹配的姓名
- 将匹配的PDF文件移动到"匹配结果"文件夹（原文件夹中不保留）
- 在Excel中标注匹配状态（成功/失败）

**使用方法：**

```bash
uv run python match_pdfs_by_name.py <PDF文件夹路径> <Excel文件路径> [--output-dir <输出文件夹名>]
```

**示例：**

```bash
uv run python match_pdfs_by_name.py /path/to/pdfs /path/to/excel.xlsx --output-dir "匹配结果"
```

**Excel文件要求：**

- 第一行为表头
- 必须包含"姓名"列

**PDF文件名格式（支持多种格式）：**

- `陈玲-承诺书.pdf`
- `承诺书-陈冬如.pdf`
- `吴慧贤-承诺书(2).pdf`
- `承诺书-姓名(数字).pdf`

**输出：**

- 匹配的文件移动到"匹配结果"文件夹（原文件夹中不再保留）
- Excel文件中添加"匹配状态"列，标注"成功"或"失败"

---

## 通用特性

### Excel文件格式支持

- 支持 `.xls` 和 `.xlsx` 格式
- 自动检测文件格式并选择合适的引擎
- 自动备份原Excel文件（添加 `.backup` 扩展名）

### 错误处理

- 文件不存在或路径错误
- Excel格式错误或找不到必要的列
- PDF文件名格式不匹配
- 文件操作权限错误

### 输出信息

- 详细的处理进度日志
- 每个文件的处理状态
- 统计信息（总数、成功、失败、错误）

## 开发说明

### 代码规范

- 项目使用 `ruff` 进行代码检查和格式化
- 所有代码已通过 `ruff check` 检查
- 代码格式符合 `ruff format` 标准

### 运行测试

```bash
# 代码质量检查
uv tool run ruff check .

# 代码格式检查
uv tool run ruff format --check .
```

## 注意事项

1. **文件备份**：所有脚本在修改Excel文件前会自动创建备份
2. **文件格式**：`.xls` 格式的文件在保存时可能会转换为 `.xlsx` 格式
3. **文件移动**：`match_pdfs_by_name.py` 使用移动操作，匹配的文件会从原文件夹移除
4. **文件复制**：`match_pdfs.py` 使用复制操作，原文件会保留
5. **代码质量**：项目已通过 ruff 代码检查，建议在提交代码前运行 `ruff check` 和 `ruff format`

## 许可证

详见 [LICENSE](LICENSE) 文件。
