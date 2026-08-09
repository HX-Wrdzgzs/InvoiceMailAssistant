# 开票邮件助手

Windows 10/11 本地桌面工具，用于读取腾讯企业邮箱中的指定中外运开票申请邮件，解析结构化字段，经 SQLite 去重和持久化后写入现有 `开票登记表.xlsx` 的 `中外运` 工作表。

## 当前实现

- .NET 8 + WPF。
- MailKit 通过 IMAP 以只读方式读取收件箱，不依赖未读状态。
- 发件人必须标准化后等于 `sino-esign@sinotrans.com`。
- 主题必须以 `中外运向您提交了开票申请` 开头。
- 标签式正文解析，兼容中英文冒号、空格、HTML `<br>` 与 `mailto:`。
- SQLite 作为事实记录层。
- Message-ID、邮箱身份 + Folder + UIDVALIDITY + IMAP UID、fallback hash 三层去重；SQLite UNIQUE 冲突会归类为重复而不会中断监听任务。
- 解析失败的目标邮件也会落库，避免静默丢失。
- ClosedXML 只写 A/B/C/D/F；E/G/H 始终不自动填写。
- Excel 写入采用同目录临时文件、校验后替换。
- Excel 写入会校验已预留行；重启时复用空的计划行或识别已写入的相同数据，发现人工占用时重新规划，不自动修改 E/G/H 人工填写列。
- Excel 被占用时进入 `PendingExcel`，关闭 Excel 后可以自动或手动重试。
- 自动监听默认每 60 秒执行一次，可配置 30–300 秒。
- 首次配置邮箱默认从当前时刻开始监听，不会自动补扫多年历史邮件；设置页提供 7 天、30 天和自定义历史补扫入口，补扫仍经过同一套 SQLite 去重。
- 登录凭据使用 Windows 当前用户的数据保护机制后保存到 LocalAppData，不进入 `settings.json`。
- `settings.json` 使用临时文件、写入刷新、JSON 校验后替换；普通设置变化不会重置监听起始时间，只有更换邮箱账户才会建立新的起始点。
- 基础界面包括概览、处理记录、等待处理和设置；支持系统托盘、关闭到托盘、第二次启动激活已有窗口、开机启动和失败记录重新解析。

## 本地数据目录

`%LOCALAPPDATA%\InvoiceMailAssistant`

包含：

- `invoice-mail.db`
- `settings.json`
- `credentials\`

这些运行时数据均已通过 `.gitignore` 排除。

## 构建与测试

```powershell
dotnet restore
dotnet build .\InvoiceMailAssistant.App\InvoiceMailAssistant.App.csproj -c Release
dotnet test .\InvoiceMailAssistant.Tests\InvoiceMailAssistant.Tests.csproj -c Release
```

## 当前未完成

以下内容仍属于后续阶段，不能把当前版本描述为完整生产版：

- 处理记录搜索、筛选和详情页。
- SQLite 自动备份。
- Serilog 结构化日志与保留策略。
- Excel 崩溃恢复幂等逻辑已实现；真实登记表兼容性验证必须由使用者提供隔离副本路径执行，不把业务文件放入仓库。
- self-contained 单文件 EXE 已提供；Windows 安装包仍未实现。

## 当前验证状态

已使用 Windows 原生 .NET SDK 执行验证：根目录解决方案可执行 `dotnet restore`，App 的 Release build 成功（0 个错误、0 个警告）；未配置真实工作簿时测试为 30 通过、1 跳过，配置隔离副本后为 31 通过、0 跳过。

自动测试覆盖 parser、金额/HTML/mailto、SQLite 并发 UNIQUE 去重、UIDVALIDITY、MonitorFromUtc 边界、邮箱重连退避、200+ 积压、严格发件人、Excel E/G/H 保留、计划行恢复、已写行幂等、人工占用/插行重规划、跨年日期，以及真实工作簿的工作表集合、公式、列宽、合并单元格、冻结窗格、打印设置和原有人工列内容。

真实工作簿验证应使用本机隔离副本，原始文件不应被测试写入。当前版本仍需真实企业邮箱登录、断网恢复和长时间无人值守观察后，才能提升为生产状态。
