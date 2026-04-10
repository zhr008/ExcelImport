# ExcelImport

ExcelImport 是一个基于 .NET 10 的 Excel 自动导入系统，用于按模板读取 Excel 文件、格式化字段、并将数据写入 SQL Server 或通过 WebApi 入库。

项目包含两个运行端：

- `ExcelImport.WinForms`：桌面端，提供配置界面、目录扫描、定时任务、手动执行、日志查看、托盘运行、开机启动
- `ExcelImport.WebApi`：服务端接口，接收导入记录、按模板校验和格式化后写入数据库
- `ExcelImport.Core`：共享模型、模板定义、公共格式化逻辑和路径解析逻辑

---

## 1. 项目目标

本项目用于解决固定格式 Excel 文件的自动采集与入库问题，支持两种典型场景：

1. **按行读取**：适合明细表、记录表、多行数据连续导入
2. **按单元格读取**：适合检测报告、汇总表、固定位置字段提取

系统通过 JSON 模板定义 Excel 字段与目标字段的映射关系，避免把导入逻辑写死在代码中，便于后续扩展不同表格格式。

---

## 2. 技术栈

- .NET 10
- WinForms
- ASP.NET Core WebApi
- ClosedXML
- Microsoft.Data.SqlClient
- log4net
- Swagger / OpenAPI

---

## 3. 解决方案结构

```text
D:\ExcelImport
├─ README.md
├─ ExcelImport.slnx
├─ ExcelImport.Core
│  ├─ ExcelImport.Core.csproj
│  ├─ Models
│  │  ├─ AppSettings.cs
│  │  ├─ DatabaseOptions.cs
│  │  ├─ DynamicEntityRecord.cs
│  │  ├─ ImportResult.cs
│  │  ├─ TemplateTaskConfig.cs
│  │  ├─ WebApiOptions.cs
│  │  ├─ WebApiPayload.cs
│  │  └─ Template
│  │     ├─ ExcelTemplateDefinition.cs
│  │     ├─ ExcelTemplateFieldDefinition.cs
│  │     └─ ExcelTemplateReadMode.cs
│  └─ Services
│     ├─ RecordFormatterService.cs
│     └─ SharedPathResolver.cs
├─ ExcelImport.WinForms
│  ├─ ExcelImport.WinForms.csproj
│  ├─ Form1.cs
│  ├─ Program.cs
│  ├─ appsettings.json
│  ├─ log4net.config
│  └─ Services
│     ├─ ConfigService.cs
│     ├─ ExcelReaderService.cs
│     ├─ ImportService.cs
│     ├─ LoggingService.cs
│     ├─ SchedulerService.cs
│     ├─ SqlServerService.cs
│     ├─ StartupService.cs
│     ├─ TrayService.cs
│     └─ WebApiService.cs
├─ ExcelImport.WebApi
│  ├─ ExcelImport.WebApi.csproj
│  ├─ Program.cs
│  ├─ appsettings.json
│  ├─ appsettings.Development.json
│  ├─ log4net.config
│  ├─ Controllers
│  │  └─ ExcelImportController.cs
│  ├─ Models
│  │  └─ ExcelImportRequest.cs
│  └─ Services
│     ├─ ConfigService.cs
│     └─ SqlServerService.cs
├─ Template
│  ├─ row-template.json
│  └─ cell-template.json
└─ OutPut
   ├─ WinForms
   ├─ WebApi
   ├─ Template
   └─ Logs
```

---

## 4. 各项目职责

### 4.1 ExcelImport.Core

公共核心层，供 WinForms 和 WebApi 共享。

主要内容：

- 应用配置模型
- 模板配置模型
- WebApi 传输模型
- 导入结果模型
- Excel 记录格式化逻辑
- 共享模板目录解析逻辑

关键文件：

- `ExcelImport.Core/Models/AppSettings.cs`
- `ExcelImport.Core/Models/TemplateTaskConfig.cs`
- `ExcelImport.Core/Models/Template/ExcelTemplateDefinition.cs`
- `ExcelImport.Core/Models/Template/ExcelTemplateFieldDefinition.cs`
- `ExcelImport.Core/Services/RecordFormatterService.cs`
- `ExcelImport.Core/Services/SharedPathResolver.cs`

### 4.2 ExcelImport.WinForms

桌面端配置和执行程序。

主要职责：

- 加载和保存配置
- 展示模板任务列表
- 选择监控目录和模板文件
- 定时扫描目录中的 Excel 文件
- 手动立即执行某个模板任务
- 调用本地数据库写入或调用 WebApi 上传
- 写运行日志
- 最小化到托盘运行
- 支持开机启动

关键文件：

- `ExcelImport.WinForms/Program.cs`
- `ExcelImport.WinForms/Form1.cs`
- `ExcelImport.WinForms/Services/ImportService.cs`
- `ExcelImport.WinForms/Services/ExcelReaderService.cs`
- `ExcelImport.WinForms/Services/SchedulerService.cs`
- `ExcelImport.WinForms/Services/SqlServerService.cs`
- `ExcelImport.WinForms/Services/WebApiService.cs`
- `ExcelImport.WinForms/Services/ConfigService.cs`

### 4.3 ExcelImport.WebApi

服务端导入接口。

主要职责：

- 接收 WinForms 提交的导入请求
- 按模板文件重新格式化记录
- 调用数据库写入逻辑
- 返回成功、插入数、跳过数等结果
- 提供 Swagger 接口文档

关键文件：

- `ExcelImport.WebApi/Program.cs`
- `ExcelImport.WebApi/Controllers/ExcelImportController.cs`
- `ExcelImport.WebApi/Models/ExcelImportRequest.cs`
- `ExcelImport.WebApi/Services/ConfigService.cs`
- `ExcelImport.WebApi/Services/SqlServerService.cs`

---

## 5. 系统工作流程

### 5.1 WinForms 导入流程

WinForms 端的导入主流程如下：

1. 从 `appsettings.json` 加载系统配置
2. 加载模板任务列表
3. 定时器按 `IntervalMinutes` 扫描监控目录
4. 根据 `FilePattern` 找到待处理 Excel 文件
5. 读取模板 JSON
6. 按模板读取 Excel 内容
7. 按模板字段类型进行格式化和校验
8. 根据配置选择导入目标：
   - 直接写入 SQL Server
   - 发送到 WebApi 接口
9. 成功文件移动到 `Processed` 目录
10. 失败文件移动到 `Failed` 目录
11. 写入日志并统计导入结果

对应核心逻辑主要位于：

- `ExcelImport.WinForms/Services/ImportService.cs:30`
- `ExcelImport.WinForms/Services/ExcelReaderService.cs:8`
- `ExcelImport.Core/Services/RecordFormatterService.cs:7`

### 5.2 WebApi 导入流程

1. 接收 `POST /api/excel-import`
2. 检查请求中的 `Records` 和 `TemplateFile`
3. 加载模板文件
4. 对传入记录重新按模板做格式化和校验
5. 写入数据库
6. 返回处理结果

对应入口位于：

- `ExcelImport.WebApi/Controllers/ExcelImportController.cs:33`

---

## 6. Excel 模板机制

模板文件位于顶层共享目录：

```text
Template/
```

当前示例模板：

- `Template/row-template.json`
- `Template/cell-template.json`

模板用于描述：

- Excel 读取模式
- 起始行
- 字段名
- Excel 列或单元格位置
- 字段类型
- 是否必填
- 最大长度
- 是否主键字段

### 6.1 模板读取模式

#### Row 模式

适用于按行读取连续数据。

示例：

```json
{
  "ReadMode": "Row",
  "StartRow": 2,
  "Fields": [
    { "Name": "TestId", "Column": "A", "Type": "string", "IsKey": true, "Required": true },
    { "Name": "TestTime", "Column": "E", "Type": "datetime", "IsKey": true, "Required": true }
  ]
}
```

含义：

- 从第 2 行开始读取
- 每个字段对应一个 Excel 列
- 遇到整行为空时停止读取

实现位置：

- `ExcelImport.WinForms/Services/ExcelReaderService.cs:21`

#### Cell 模式

适用于固定坐标提取。

示例：

```json
{
  "ReadMode": "Cell",
  "Fields": [
    { "Name": "TestId", "Cell": "A2", "Type": "string", "Required": true },
    { "Name": "TestTime", "Cell": "B6", "Type": "datetime", "Required": true }
  ]
}
```

含义：

- 每个字段直接指定单元格坐标
- 一份文件通常只生成一条记录

实现位置：

- `ExcelImport.WinForms/Services/ExcelReaderService.cs:65`

### 6.2 字段定义说明

模板字段定义见：

- `ExcelImport.Core/Models/Template/ExcelTemplateFieldDefinition.cs:3`

字段含义：

- `Name`：目标字段名
- `Column`：Row 模式下的 Excel 列名，如 `A`、`B`
- `Cell`：Cell 模式下的单元格地址，如 `A2`
- `Type`：字段类型，当前支持 `string`、`int`、`decimal`、`datetime`、`bool`
- `Required`：是否必填
- `Length`：字符串最大长度
- `IsKey`：是否关键字段

### 6.3 字段格式化规则

记录在导入前会统一经过格式化逻辑处理：

- 空值在必填字段中会报错
- `int` 使用 `int.Parse`
- `decimal` 使用 `decimal.Parse`
- `datetime` 使用 `DateTime.Parse`
- `bool` 支持以下文本：
  - true：`TRUE`、`1`、`Y`、`YES`、`PASS`
  - false：`FALSE`、`0`、`N`、`NO`、`FAIL`
- 字符串可按 `Length` 限制最大长度

实现位置：

- `ExcelImport.Core/Services/RecordFormatterService.cs:15`

---

## 7. WinForms 配置说明

WinForms 默认配置文件：

- `ExcelImport.WinForms/appsettings.json`

发布后配置位置：

- `OutPut/WinForms/appsettings.json`

说明：

- WinForms 不再自动生成默认 `appsettings.json`
- 如果发布目录缺少 `appsettings.json`，程序会在启动加载配置时直接报错

当前配置结构示例：

```json
{
  "Database": {
    "Enabled": false,
    "ConnectionString": "Server=.;Database=TEST;Trusted_Connection=True;TrustServerCertificate=True;"
  },
  "WebApi": {
    "Enabled": true,
    "BaseUrl": "http://localhost:7001",
    "Endpoint": "/api/excel-import"
  },
  "StartWithWindows": false,
  "LogDirectory": "Logs",
  "Templates": [
    {
      "Name": "RowTemplate",
      "Enabled": true,
      "WatchPath": "D:\\ExcelImport\\Excel\\Row",
      "IncludeSubdirectories": true,
      "IntervalMinutes": 3,
      "TemplateFile": "row-template.json",
      "TargetTable": "dbo.ExcelImportRow",
      "FilePattern": "*.xlsx"
    }
  ]
}
```

### 7.1 Database

- `Enabled`：是否启用数据库直写
- `ConnectionString`：SQL Server 连接字符串

### 7.2 WebApi

- `Enabled`：是否启用接口导入
- `BaseUrl`：WebApi 服务地址
- `Endpoint`：接口路径，默认 `/api/excel-import`

### 7.3 StartWithWindows

- 是否启用开机启动

### 7.4 LogDirectory

- 日志目录名，通常为 `Logs`

### 7.5 Templates

模板任务列表，每项对应一个目录扫描任务。

任务结构定义见：

- `ExcelImport.Core/Models/TemplateTaskConfig.cs:3`

字段说明：

- `Name`：任务名称
- `Enabled`：是否启用
- `WatchPath`：监控目录
- `IncludeSubdirectories`：是否扫描子目录
- `IntervalMinutes`：扫描间隔分钟数
- `TemplateFile`：模板文件名，相对共享模板目录
- `TargetTable`：目标数据库表名
- `FilePattern`：文件匹配模式，如 `*.xlsx`

---

## 8. WebApi 配置说明

WebApi 默认配置文件：

- `ExcelImport.WebApi/appsettings.json`

发布后配置位置：

- `OutPut/WebApi/appsettings.json`

当前配置示例：

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "Database": {
    "ConnectionString": "Server=.;Database=TEST;Trusted_Connection=True;TrustServerCertificate=True;"
  },
  "AllowedHosts": "*",
  "Kestrel": {
    "Endpoints": {
      "Http": {
        "Url": "http://localhost:7001"
      }
    }
  }
}
```

字段说明：

- `Database.ConnectionString`：数据库连接字符串，WebApi 导入时必须可用
- `Kestrel.Endpoints.Http.Url`：监听地址

说明：

- WebApi 不再使用 `Database.Enabled`
- `Database.ConnectionString` 为空时，请求会在入口处直接失败并抛出 `未配置 Database:ConnectionString。`
- 显式校验位置：`ExcelImport.WebApi/Controllers/ExcelImportController.cs:36`

---

## 9. API 说明

### 9.1 导入接口

- 方法：`POST`
- 路径：`/api/excel-import`

控制器位置：

- `ExcelImport.WebApi/Controllers/ExcelImportController.cs:11`

### 9.2 请求结构

请求模型定义：

- `ExcelImport.WebApi/Models/ExcelImportRequest.cs:4`

示例：

```json
{
  "templateName": "RowTemplate",
  "fileName": "sample.xlsx",
  "templateFile": "row-template.json",
  "targetTable": "dbo.ExcelImportRow",
  "records": [
    {
      "TestId": "T001",
      "MeasuringSide": "A",
      "TaskNumber": "TASK-001",
      "Description": "sample",
      "TestTime": "2026-04-09 10:00:00",
      "Operator": "Tom",
      "IsPass": true,
      "Wavelength1310_IL": 1.23,
      "Wavelength1310_RL": 2.34
    }
  ]
}
```

### 9.3 返回结果

成功时会返回：

- `Success`
- `Message`
- `TemplateName`
- `FileName`
- `TemplateFile`
- `TargetTable`
- `RecordCount`
- `InsertedCount`

---

## 10. 文件处理规则

WinForms 扫描目录中的 Excel 文件后，会按修改日期决定是否移动文件：

- 修改日期不是当天：处理完成后移动
  - 成功：移动到 `Succeed`
  - 失败：移动到 `Failed`
- 修改日期是当天：处理完成后不移动，保留在原目录，允许后续继续执行

规则实现位于：

- `ExcelImport.WinForms/Services/ImportService.cs:123`

说明：

- 若目标目录不存在，会自动创建
- 若目标文件同名，会自动追加时间戳避免覆盖
- `Succeed` 和 `Failed` 目录本身不会被重复扫描

---

## 11. 日志与运行行为

### 11.1 WinForms

WinForms 启动时会：

- 加载配置
- 输出配置路径、程序目录、Template 目录
- 启动调度器
- 在界面中展示日志

主要逻辑位于：

- `ExcelImport.WinForms/Form1.cs:50`

### 11.2 单实例运行

WinForms 通过命名互斥体保证单实例运行；若重复启动，会向已存在实例发送激活消息。

实现位置：

- `ExcelImport.WinForms/Program.cs:24`

### 11.3 托盘最小化启动

支持参数：

```text
--startup-minimized
```

程序在收到该参数时会启动后最小化到托盘。

实现位置：

- `ExcelImport.WinForms/Program.cs:7`
- `ExcelImport.WinForms/Form1.cs:291`

---

## 12. 模板目录与发布目录规则

### 12.1 模板目录

全项目统一使用顶层共享模板目录：

```text
Template/
```

发布后模板目录位于：

```text
OutPut/Template/
```

WinForms 发布时会自动把模板复制到：

```text
OutPut/Template/
```

相关发布逻辑位于：

- `ExcelImport.WinForms/ExcelImport.WinForms.csproj:54`

### 12.2 配置目录

- WinForms 配置：`OutPut/WinForms/appsettings.json`
- WebApi 配置：`OutPut/WebApi/appsettings.json`

### 12.3 不再采用的方式

以下方式不推荐继续使用：

- 为 WinForms 和 WebApi 分别维护模板副本
- 将两个项目发布到同一个平铺目录
- 通过多个不同命名的配置文件兼容同目录部署

---

## 13. 构建与发布

### 13.1 构建整个解决方案

```bash
dotnet build "/d/ExcelImport/ExcelImport.slnx"
```

### 13.2 发布 WinForms

```bash
dotnet publish "/d/ExcelImport/ExcelImport.WinForms/ExcelImport.WinForms.csproj" -c Release
```

输出目录：

```text
D:\ExcelImport\OutPut\WinForms
```

### 13.3 发布 WebApi

```bash
dotnet publish "/d/ExcelImport/ExcelImport.WebApi/ExcelImport.WebApi.csproj" -c Release
```

输出目录：

```text
D:\ExcelImport\OutPut\WebApi
```

### 13.4 同时发布两个项目

```bash
dotnet publish "/d/ExcelImport/ExcelImport.WinForms/ExcelImport.WinForms.csproj" -c Release && dotnet publish "/d/ExcelImport/ExcelImport.WebApi/ExcelImport.WebApi.csproj" -c Release
```

### 13.5 查看发布结果

```bash
ls -R "/d/ExcelImport/OutPut"
```

---

## 14. 运行方式

### 14.1 启动 WinForms

运行：

```text
OutPut/WinForms/ExcelImport.WinForms.exe
```

### 14.2 启动 WebApi

运行：

```text
OutPut/WebApi/ExcelImport.WebApi.exe
```

### 14.3 Swagger

WebApi 启动后可通过 Swagger 查看接口文档。

Swagger 已在程序启动时注册：

- `ExcelImport.WebApi/Program.cs:17`
- `ExcelImport.WebApi/Program.cs:23`

---

## 15. 二次开发建议

### 修改 WinForms 行为

优先查看：

- `ExcelImport.WinForms/Form1.cs`
- `ExcelImport.WinForms/Services/ImportService.cs`
- `ExcelImport.WinForms/Services/SchedulerService.cs`

### 修改 WebApi 入库逻辑

优先查看：

- `ExcelImport.WebApi/Controllers/ExcelImportController.cs`
- `ExcelImport.WebApi/Services/SqlServerService.cs`

### 修改模板结构或字段转换规则

优先查看：

- `ExcelImport.Core/Models/Template/*`
- `ExcelImport.Core/Services/RecordFormatterService.cs`
- `ExcelImport.WinForms/Services/ExcelReaderService.cs`

### 新增一种 Excel 模板

一般步骤：

1. 在 `Template/` 下新增 JSON 模板文件
2. 在 WinForms 配置中新增一个模板任务
3. 指定监控目录、目标表、文件匹配规则
4. 运行 WinForms 手动测试或等待定时任务执行

---

## 16. 当前默认示例

项目当前内置了两类模板任务：

### 16.1 RowTemplate

- 模板：`row-template.json`
- 监控目录：`D:\ExcelImport\Excel\Row`
- 目标表：`dbo.ExcelImportRow`
- 文件类型：`*.xlsx`

### 16.2 CellTemplate

- 模板：`cell-template.json`
- 监控目录：`D:\ExcelImport\Excel\Cell`
- 目标表：`dbo.ExcelImportCell`
- 文件类型：`*.xlsx`

---

## 17. 注意事项

1. `TemplateFile` 应始终指向共享模板目录下的文件
2. `WatchPath` 必须是真实存在的目录，否则任务会被跳过
3. 至少启用一种导入方式：`Database.Enabled` 或 `WebApi.Enabled`
4. WebApi 模式下，服务端仍会再次按模板做校验和格式化
5. Excel 文件处理完成后会被移动，因此监控目录建议只放待处理文件
6. `Processed` 和 `Failed` 目录不要手动作为监控目录
7. 数据库表结构需要与模板字段保持一致

---

## 18. 后续可扩展方向

如果后续要继续增强，比较适合扩展的方向包括：

- 支持更多字段类型
- 支持自定义字段转换器
- 支持多工作表读取
- 支持失败重试
- 支持归档策略
- 支持按模板配置不同数据库连接
- 支持导入前预检查和预览

---

## 19. 关键源码入口速查

- WinForms 启动入口：`ExcelImport.WinForms/Program.cs:22`
- WinForms 主界面：`ExcelImport.WinForms/Form1.cs:26`
- Excel 读取：`ExcelImport.WinForms/Services/ExcelReaderService.cs:8`
- 导入主流程：`ExcelImport.WinForms/Services/ImportService.cs:30`
- 记录格式化：`ExcelImport.Core/Services/RecordFormatterService.cs:7`
- WebApi 启动入口：`ExcelImport.WebApi/Program.cs:5`
- WebApi 导入接口：`ExcelImport.WebApi/Controllers/ExcelImportController.cs:33`

---

## 20. 总结

ExcelImport 当前已经形成了较清晰的三层结构：

- `Core` 负责共享定义与公共逻辑
- `WinForms` 负责本地扫描、调度、人工配置和执行
- `WebApi` 负责远程接收与统一入库

通过共享模板目录和统一发布结构，项目可以在保持部署简单的同时，支持本地直写和接口导入两种模式。后续如需新增模板或扩展字段规则，也可以在现有结构上继续演进。
