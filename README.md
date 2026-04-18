# LiteEMLTOPST - 邮件转换工具

一款专业的 Windows 邮件转换工具，支持 EML、PST、OST 等格式之间的转换，以及 IMAP 邮件收取功能。

## 功能特点

### 1. EML 转 PST 转换
- 支持批量转换 EML 文件为 PST 格式
- **智能分类功能**：
  - 按年份归类（创建 `Inbox/2024` 等文件夹）
  - 按发件人域名归类（创建 `Inbox/gmail.com` 等文件夹）
  - 按月份归类（创建 `Inbox/2024-01` 等文件夹）
- 支持跳过已归类文件夹，避免重复处理
- 保留原始目录结构
- 预估邮件数量，转换进度实时显示

### 2. IMAP 邮件收取
- 支持从主流邮箱服务器收取邮件（QQ邮箱、网易邮箱等）
- **收取到 EML**：按文件夹结构保存到本地
- **收取到 PST**：直接导入 PST 文件
- **智能分类**：勾选归类选项后，自动按年份/发件人/月份创建文件夹
- 支持选择特定文件夹收取
- 文件夹多选功能

### 3. PST/OST 文件提取
- 从 PST 文件提取邮件为 EML 格式
- 从 OST 文件（Outlook 缓存）提取邮件
- 支持按文件夹结构保留导出

### 4. EML 日期分割
- 按指定日期将邮件分割到不同目录
- 区分已发送和收件箱邮件
- 支持分类模式下的目录结构保留

### 5. Office 365 同步
- 支持 Microsoft 365 云端同步功能

## 系统要求

- Windows 10/11
- .NET 8.0 Runtime
- Microsoft Outlook（部分功能需要）
- IMAP 邮件服务器访问权限

## 安装说明

1. 下载最新版本的 ZIP 包
2. 解压到任意目录
3. 运行 `MailConvertPrivateUser.exe`

## 使用指南

### EML 转 PST

1. 切换到 **📧 EML 转 PST** 标签页
2. 选择 EML 文件所在输入目录
3. 选择 PST 输出目录
4. 勾选归类方式（可选）：
   - ☑️ 按年份 - 按邮件年份放入对应文件夹
   - ☐ 按发件人 - 按发件人域名分类
   - ☐ 按月份 - 按年份-月份分类
5. 点击 **扫描** 预估邮件数量
6. 点击 **开始转换**

### IMAP 收取邮件

1. 切换到 **📥 IMAP 收取** 标签页
2. 选择或添加邮箱配置
3. 输入邮箱地址和密码（或授权码）
4. 点击 **刷新** 获取文件夹列表
5. 按住 Ctrl 多选需要收取的文件夹
6. 勾选归类方式（可选）
7. 选择输出目录
8. 点击 **收取到 PST** 或 **收取到 EML**

### 归类说明

归类功能会在指定目录下创建类似以下结构：

```
输出目录/
├── Inbox/
│   ├── 2024/
│   │   ├── gmail.com/
│   │   │   └── 邮件1.eml
│   │   └── qq.com/
│   │       └── 邮件2.eml
│   └── 2023/
│       └── 邮件3.eml
```

## 技术架构

### 核心技术栈

| 组件 | 技术 | 说明 |
|------|------|------|
| 框架 | .NET 8.0 + WPF | 现代化桌面应用 |
| 邮件解析 | MimeKit | 强大的 MIME 邮件解析库 |
| IMAP 客户端 | MailKit | 支持 IMAP/SMTP 协议 |
| PST 操作 | Outlook COM Interop | 通过 Outlook 操作 PST 文件 |
| 配置管理 | Microsoft.Extensions.Configuration | JSON 配置文件支持 |
| 日志 | Serilog | 结构化日志记录 |

### 项目结构

```
LiteEMLTOPST/
├── src/MailConvertPrivateUser/
│   ├── Models/              # 数据模型
│   ├── Services/            # 核心服务
│   │   ├── ImapExtractService.cs    # IMAP 收取
│   │   ├── PstExtractService.cs     # PST 提取
│   │   ├── OstExtractService.cs     # OST 提取
│   │   ├── PstWriterService.cs      # PST 写入
│   │   ├── EmailParserService.cs    # 邮件解析
│   │   └── ClassificationService.cs # 分类服务
│   ├── MainWindow.xaml      # 主窗口
│   └── App.xaml             # 应用入口
├── tests/                   # 单元测试
└── README.md
```

## 注意事项

1. **Outlook 依赖**：EML 转 PST 和 OST 提取功能需要本地安装 Microsoft Outlook
2. **IMAP 限制**：部分邮件服务器默认限制单次收取数量，建议分批收取
3. **授权码**：使用 QQ邮箱、网易邮箱等需要开启 IMAP 并获取授权码
4. **分类覆盖**：如果勾选多个分类选项，会按 年份+发件人 → 年份+月份 → 发件人+月份 的优先级组合

## 日志说明

程序运行日志保存在：
```
%LOCALAPPDATA%\MailConvertPrivateUser\logs\
```

按功能模块分目录存储：
- `EML2PST/` - EML 转 PST 日志
- `IMAP/` - IMAP 收取日志
- `PST/` - PST 操作日志
- `OST/` - OST 操作日志
- `O365/` - Office 365 日志

## 常见问题

### Q: 转换时界面冻结？
A: 程序已优化为异步处理，如仍出现冻结可能是 Outlook COM 操作繁忙，请耐心等待。

### Q: PST 文件创建失败？
A: 请确保：
- Outlook 已正确安装
- 有权在目标目录创建文件
- Outlook 未以管理员权限运行（可能导致权限问题）

### Q: IMAP 收取失败？
A: 请检查：
- 邮箱是否开启了 IMAP 服务
- 密码是否使用授权码而非登录密码
- 网络能否访问邮箱服务器

### Q: 如何获得软件订阅？
A: 请访问 [https://www.booming.one](https://www.booming.one) 了解订阅详情。

### Q: 试用期满后还能继续使用吗？
A: 试用期结束后需要订阅才能继续使用。请在试用期结束前通过官网续费。

## 关于订阅模式

LiteEMLTOPST 是一款商业化邮件转换工具。

### 注册与订阅流程

1. **注册软件**：下载安装后，输入邮箱进行注册
2. **免费试用**：注册成功后自动获得 **30 天试用版**
3. **订阅续费**：试用期结束后，通过以下方式订阅：
   - 网站：[https://www.booming.one](https://www.booming.one)
   - 续费后获得正式授权

### 订阅权益
- 持续功能更新
- 技术支持服务
- 优先体验新功能

## 许可证

本文档仅供已订阅用户内部使用。

## 联系方式

- 商务合作：raymond.xu@booming.one

---

如有任何问题或建议，欢迎通过邮箱联系。
