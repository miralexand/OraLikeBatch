# OraLikeBatch

Oracle 批量模糊查询小工具，基于 `tkinter`、`python-oracledb`、`pandas` 和 `openpyxl`。

支持从 Excel 复制一列后批量粘贴查询，按批生成 `LIKE` 条件，查询结果合并去重，并导出包含“未匹配值”的 Excel 文件。数据库连接信息可在界面中保存，密码使用 Windows DPAPI 按当前用户加密。

## 功能特点

- 图形界面，无需命令行操作
- 支持测试数据库连接
- 每行一个待匹配值，自动忽略空行和重复值
- 支持从 Excel 复制整列后粘贴，也可直接 `Ctrl+V`
- 每批最多 200 个值，分批查询后合并去重
- 查询结果额外增加“命中值”列，标明该行命中了哪些粘贴值
- 导出 Excel 默认带时间戳，包含“查询结果”和“未匹配值”两个 sheet
- 支持记住密码，使用 Windows DPAPI 加密保存
- 测试连接、查询、导出均在后台线程执行，界面不卡死
- 查询值全部使用绑定变量；表名、字段名进行白名单校验
- 所有错误提示均为中文可读信息

## 环境要求

- Windows 10 / 11
- Python 3.10+（推荐 3.11）
- 依赖见 `requirements.txt`

安装依赖：

```powershell
pip install -r requirements.txt
```

## 直接运行

```powershell
python oracle_batch_query.py
```

## 打包为 exe

已提供 PyInstaller 打包脚本，会自动创建隔离的 `.venv-build` 环境，避免受本机已有 Python 包影响。

双击 `build_exe.bat`，或运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```

打包完成后生成：

```text
dist\OracleBatchQuery.exe
```

仓库中也已包含构建好的 Windows 单文件程序：`dist/OracleBatchQuery.exe`。单文件模式首次启动会解压运行文件，可能比直接运行源码慢几秒。

## 使用步骤

1. 填写数据库用户名、密码和连接串，点击“测试连接”。
2. 填写表名、匹配字段、查询字段。
   - 查询字段用英文或中文逗号分隔。
   - 查询字段留空表示 `SELECT *`。
3. 将待匹配值粘贴到多行文本框，每行一个值。
   - 可在 Excel 中复制整列后点击“从剪贴板粘贴一列”。
   - 也可直接在文本框中按 `Ctrl+V`。
4. 点击“执行查询”，结果会显示在下方 Treeview 中。
5. 点击“导出 Excel”，选择保存路径。

## 模糊匹配规则

对每个粘贴值生成一个 `LIKE` 条件，例如：

```sql
SELECT 字段1, 字段2, 匹配字段
FROM 表名
WHERE (匹配字段 LIKE :1 OR 匹配字段 LIKE :2 OR ...)
```

绑定参数统一传入 `%值%`，不会把用户输入拼接进 SQL。

- 每批最多 200 个值，避免 `OR` 条件过多导致查询变慢。
- 分批结果按整行内容合并去重。
- “命中值”列按行判断，一行命中多个值时会用 ` | ` 连接。
- 粘贴值中的 `%` 和 `_` 会被 Oracle 当作 `LIKE` 通配符，本工具不做转义。

## 密码保存

- 勾选“记住密码”后，点击“保存连接”即可保存。
- 测试连接或查询成功后，如果已勾选“记住密码”，也会自动保存。
- 如果不勾选“记住密码”，只保存用户名和连接串，不保存密码。
- 密码通过 Windows DPAPI 按当前 Windows 用户加密，配置文件位置：

```text
%APPDATA%\OracleBatchQuery\connection.json
```

- 换 Windows 用户或换电脑后，已保存的密码无法解密，需要重新输入。
- 点击“清除保存”可删除本机保存的连接信息。

## 表名和字段名限制

表名、匹配字段、查询字段仅允许字母、数字、下划线和点号，例如：

```text
SCOTT.EMP
EMP
EMP.NAME
```

不支持别名、函数、带引号的标识符或任意 SQL 片段。Oracle 中未加双引号创建的表名、字段名默认按大写存储；如果建表时使用了双引号，请按实际大小写填写。

## 结果与导出说明

- 界面最多显示前 2000 行，避免结果过大影响操作。
- 导出的 Excel 包含全部查询结果。
- “查询结果”sheet 包含查询结果列和“命中值”列。
- “未匹配值”sheet 列出所有一次都没有命中的粘贴值。
- 界面状态栏会显示总行数、未匹配值数量及分批执行进度。

## 常见问题

**提示 ORA-01017**

用户名或密码错误，请检查登录凭据。

**提示 ORA-00942**

表或视图不存在，请确认表名、schema 以及当前用户是否有查询权限。

**提示 ORA-00904**

字段名无效，请检查匹配字段和查询字段拼写。

**提示 DPY-6005 / ORA-12170**

无法连接数据库，请检查连接串、网络、防火墙和数据库监听状态。

**连接串怎么写**

常见格式为：

```text
主机:端口/服务名
```

例如：

```text
127.0.0.1:1521/orclpdb1
```

也可以填写本机 `tnsnames.ora` 中配置的 TNS 别名。

**是否需要安装 Oracle Instant Client**

默认使用 python-oracledb 的 Thin 模式，一般不需要。若必须使用 Thick 模式，需要初始化 Oracle Instant Client，并确保其位数与 Python 一致。

**exe 被杀毒软件拦截**

PyInstaller 单文件程序可能被部分杀毒软件误报，可添加信任或改用源码运行。

## 项目结构

```text
oracle_batch_query.py          主程序
requirements.txt               运行依赖
build_exe.ps1                  PyInstaller 打包脚本
build_exe.bat                  双击打包入口
dist/OracleBatchQuery.exe      已构建的 Windows 单文件程序
```
