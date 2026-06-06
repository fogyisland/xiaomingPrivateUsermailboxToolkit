# LiteEMLTOPST - 小铭邮件百宝箱（个人版）

一款专业的 Windows 邮件转换工具，支持 EML、PST、OST 等格式之间的转换，以及 IMAP 邮件收取和 Office 365 云端同步功能。

## 功能特点

### 1. 邮件导出管理
- **EML 转 PST 转换**：批量转换 EML 文件为 PST 格式
- **O365 邮箱导出**：从 Microsoft 365 导出邮件到 PST 文件
  - 支持按时间范围筛选
  - 支持导出全部邮件
  - 可选按年月归类（Inbox/2026/04）
  - 保留原始文件夹结构（Inbox、mail1、mail2 等）
- **PST 文件提取**：从 PST 文件提取邮件为 EML 格式
- **OST 文件提取**：从 Outlook 缓存文件（OST）提取邮件
- **智能分类功能**：
  - 按年份归类（创建 `Inbox/2024` 等文件夹）
  - 按月份归类（创建 `Inbox/2024-01` 等文件夹）
  - 保留原始目录结构
- 预估邮件数量，转换进度实时显示

### 2. IMAP 邮件收取
- 支持从主流邮箱服务器收取邮件（QQ邮箱、网易邮箱、企业邮箱等）
- **收取到 EML**：按文件夹结构保存到本地
- **收取到 PST**：直接导入 PST 文件
- **智能分类**：勾选归类选项后，自动按年份/月份创建文件夹
- 支持选择特定文件夹收取
- 文件夹多选功能

### 3. 联系人格式转换
- 支持批量转换联系人格式
- 支持多种格式映射
- 预览和确认功能

### 4. Office 365 同步
- **设备代码登录**：安全的 OAuth 认证方式
- **交互式登录**：支持多种认证方式
- **邮件同步**：在云端和本地之间同步邮件
- **联系人同步**：将联系人同步到 O365
- **日历同步**：将日历事件同步到 O365

## 国际化 / 多语言支持

支持中文（简体）和英文两种界面语言。

- 切换路径：**常规设置 → Language** 标签页
- 切换后需重启软件生效
- 语言文件位置：`Language/zh-cn.xml` 和 `Language/en-us.xml`（与应用同级目录）
- 翻译词条统一在 `docs/superpowers/translations/keys.csv` 维护

## 系统要求

- Windows 10/11
- .NET 8.0 Runtime
- Microsoft Outlook（EML 转 PST 和 OST 提取功能需要）
- IMAP 邮件服务器访问权限（IMAP 收取功能需要）
- Microsoft 365 账户（O365 同步功能需要）

## 安装说明

1. 下载最新版本的 ZIP 包
2. 解压到任意目录
3. 运行 `MailConvertPrivateUser.exe`

## 使用指南

### O365 邮箱导出

1. 切换到 **邮件管理** 标签页
2. 登录 Office 365 账户
3. 选择时间范围或勾选"导出全部邮件"
4. 选择输出文件路径
5. 可选：勾选"按年月归类"
6. 点击 **导出 PST**

### EML 转 PST

1. 切换到 **EML 转 PST** 标签页
2. 选择 EML 文件所在输入目录
3. 选择 PST 输出目录
4. 勾选归类方式（可选）：
   - 按年份 - 按邮件年份放入对应文件夹
   - 按月份 - 按年份-月份分类
5. 点击 **扫描** 预估邮件数量
6. 点击 **开始转换**

### IMAP 收取邮件

1. 切换到 **IMAP 收取** 标签页
2. 选择或添加邮箱配置
3. 输入邮箱地址和密码（或授权码）
4. 点击 **刷新** 获取文件夹列表
5. 按住 Ctrl 多选需要收取的文件夹
6. 勾选归类方式（可选）
7. 选择输出目录
8. 点击 **收取到 PST** 或 **收取到 EML**

### 联系人格式转换

1. 切换到 **联系人格式转换** 标签页
2. 选择源文件（CSV/VCF）
3. 选择目标格式
4. 配置字段映射
5. 点击 **开始转换**

### O365 同步

1. 切换到 **O365 同步** 标签页
2. 输入租户 ID、客户端 ID 和用户 ID
3. 点击 **连接**
4. 选择同步类型（邮件/联系人/日历）
5. 按照向导完成同步

## 归类说明

归类功能会在指定目录下创建类似以下结构：

```
输出目录/
├── Inbox/
│   ├── 2026/
│   │   ├── 04/
│   │   │   └── 邮件1.eml
│   │   └── 03/
│   │       └── 邮件2.eml
│   └── 2025/
│       └── 邮件3.eml
├── mail1/
│   ├── 2026/
│   │   └── 04/
│   │       └── 邮件4.eml
│   └── 2025/
│       └── 邮件5.eml
└── mail2/
    └── ...
```

## 技术架构

### 核心技术栈

| 组件 | 技术 | 说明 |
|------|------|------|
| 框架 | .NET 8.0 + WPF | 现代化桌面应用 |
| 邮件解析 | MimeKit | 强大的 MIME 邮件解析库 |
| IMAP 客户端 | MailKit | 支持 IMAP/SMTP 协议 |
| PST 操作 | Outlook COM Interop | 通过 Outlook 操作 PST 文件 |
| O365 | Microsoft Graph SDK | Azure AD 和 Office 365 API |
| 配置管理 | Microsoft.Extensions.Configuration | JSON 配置文件支持 |
| 日志 | Serilog | 结构化日志记录 |
| 本地化 | LocalizationManager + LocExtension | 中英文界面切换 |

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
│   │   ├── ClassificationService.cs # 分类服务
│   │   ├── Office365SyncService.cs  # O365 同步
│   │   └── LocalizationManager.cs   # i18n 词条查找
│   ├── Markup/              # XAML 标记扩展
│   │   └── LocExtension.cs          # {loc:Loc Key} 标记
│   ├── Language/            # 翻译文件
│   │   ├── zh-cn.xml                # 简体中文
│   │   └── en-us.xml                # 英文
│   ├── MainWindow.xaml      # 主窗口
│   └── App.xaml             # 应用入口
├── tests/                   # 单元测试
├── docs/superpowers/translations/keys.csv  # 翻译词条主表
└── README.md
```

## 注意事项

1. **Outlook 依赖**：EML 转 PST 和 OST 提取功能需要本地安装 Microsoft Outlook
2. **IMAP 限制**：部分邮件服务器默认限制单次收取数量，建议分批收取
3. **授权码**：使用 QQ邮箱、网易邮箱等需要开启 IMAP 并获取授权码
4. **O365 权限**：O365 同步需要正确的 API 权限配置

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

### Q: O365 登录失败？
A: 请检查：
- 租户 ID、客户端 ID 是否正确
- 应用是否在 Azure AD 中注册
- 是否有足够的 API 权限

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
