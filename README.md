# InvoiceRenamer 发票批量重命名工具

> 从一堆 PDF / 图片发票中自动提取关键字段，按你定义的规则批量重命名，支持撤销与 Excel 导出。

## 功能特性

- **智能识别**：支持增值税专票/普票、电子发票、通行费发票、铁路电子客票等多种票面格式，自动提取开票日期、发票类型、销售方/购买方、金额等字段
- **云端 OCR**：集成腾讯云 OCR，支持图片发票（JPG/PNG/BMP/TIFF）和扫描件 PDF 识别，零新增依赖（标准库实现 TC3-HMAC-SHA256 签名）
- **WebView UI**：基于 PyWebview + HTML/CSS/JS 的新版界面，渲染效果优于传统 GUI 框架
- **非发票检测**：自动跳过非发票文件（如汇总单、明细单），避免误处理
- **自定义命名规则**：拖拽排列字段顺序，勾选启用字段，实时预览生成文件名
- **批量重命名 + 一键撤销**：确认后批量执行，操作有误可一键回滚
- **表头排序**：点击表头按任意列升序/降序排列
- **Excel 导出**：识别结果一键导出为 .xlsx，方便存档核对
- **密钥安全**：API 密钥 XOR+Base64 混淆存储，不明文暴露
- **单文件 exe**：37 MB，无需安装 Python 环境直接运行

## 界面预览

![主界面](docs/软件初始化.png)
![云端识别设置](docs/配置云端设置.png)

## 快速开始

### 方式一：下载 exe（推荐）

1. 前往 [Releases](https://github.com/simple-ty/InvoiceRenamer/releases) 下载最新版 `InvoiceRenamer_vX.X.X.exe`
2. 双击运行（首次运行如被 Windows SmartScreen 拦截，点击「仍要运行」）
3. 选择发票所在文件夹 → 扫描 → 确认 → 重命名

### 方式二：源码运行

```bash
# 克隆仓库
git clone https://github.com/simple-ty/InvoiceRenamer.git
cd InvoiceRenamer

# 安装依赖
pip install -r requirements.txt

# 运行
python webapp.py
```

**依赖清单**（仅 3 个第三方库，追求轻量）：

| 库 | 用途 |
|---|---|
| pdfplumber | PDF 文本提取与发票字段解析 |
| pywebview | WebView 桌面 UI 框架 |
| openpyxl | Excel 导出 |

## 云端 OCR 配置

1. 注册 [腾讯云](https://cloud.tencent.com/) 账号
2. 开通 [增值税发票识别](https://cloud.tencent.com/product/ocr) 服务（每月免费 1000 次）
3. 在 [API 密钥管理](https://console.cloud.tencent.com/cam/capi) 获取 SecretId 和 SecretKey
4. 在软件左下角点击「云端识别设置」→ 输入密钥 → 保存

密钥使用 XOR + Base64 混淆存储在 `~/.invoice_renamer/config.json`，不明文保存。

> **提示**：选择文件夹时，工具会**递归搜索**该文件夹下所有子目录中的 PDF 和图片文件。建议将发票统一整理到一个文件夹后再选择，避免扫描范围过大。

## 项目结构

```
InvoiceRenamer/
├── webapp.py                # 主程序入口
├── webapp.spec              # PyInstaller 打包配置
├── build_webview.bat        # 一键打包脚本
├── webview/                 # 前端资源
│   ├── main.html            # 主界面 HTML
│   ├── css/main.css         # 样式
│   └── js/main.js           # 交互逻辑
├── invoice_parser.py        # PDF 解析 + 发票字段提取
├── cloud_ocr.py             # 腾讯云 OCR 模块（签名/识别/密钥管理）
├── name_builder.py          # 命名规则引擎
├── excel_exporter.py        # Excel 导出
├── config.py                # 常量与配置
├── rules.json               # 正则规则（可配置）
├── requirements.txt         # 依赖清单
└── icon.ico                 # 应用图标
```

## 配置说明

### rules.json — 正则规则

发票字段提取的正则表达式集中存放在 `rules.json`，无需改代码即可调整识别规则：

```json
{
  "date": [...],           // 开票日期
  "number": [...],         // 发票号码
  "type_keywords": [...],  // 发票类型关键词
  "buyer_patterns": [...], // 购买方
  "seller_patterns": [...],// 销售方
  "amount_patterns": [...] // 金额
}
```

### 公司名智能截断

销售方/购买方名称自动截断到第一个主后缀，降低文件名长度：

| 原始名称 | 截断结果 |
|---|---|
| 中国石化销售股份有限公司湖南石油分公司 | 中国石化销售股份有限公司 |
| 湖南投资集团股份有限公司绕城公路西南段分公司 | 湖南投资集团股份有限公司 |

主后缀：公司 / 有限公司 / 股份有限公司 / 有限责任公司 / 股份公司

## 开发说明

### 打包 exe

```bash
# 确保 PyInstaller 已安装
pip install pyinstaller

# 打包
build_webview.bat
```

输出：`dist/InvoiceRenamer_v{版本号}.exe`（约 37 MB）

打包配置说明：
- onefile 模式，单文件分发
- 无控制台窗口（windowed）
- 内嵌 rules.json + icon.ico + webview/ 前端资源
- 未启用 UPX 压缩（避免杀软误报）

## 版本历史

| 版本 | 主要更新 |
|---|---|
| **v0.5.16** | 设置对话框修复：密钥验证 **** 兜底、眼按钮 flex 布局修复点击失效、提示文案精简；清理测试文件 |
| **v0.5.11** | 实时逐行扫描 + 增量渲染；手动编辑文件名；单实例锁；6种状态精细化 + 行着色；Toast 通知 |
| **v0.5.0** | WebView 新版 UI（HTML/CSS/JS）；腾讯云 OCR 图片发票识别；扫描件 PDF 识别；API 密钥混淆存储；云端用量统计；HTTP Server 通信架构 |
| v0.4.5 | 修复 buyer/seller 后缀丢失与方向反转；长公司名智能截断；选择按钮合并为浮层下拉菜单；扫描结果逐行实时显示；表头点击排序；表头改名「销售方/行程信息」 |
| v0.4.0 | 非发票文件检测与跳过；is_invoice 三道关卡判定；列宽自适应 + 水平滚动条；拖拽排序修复；状态提示精简 |
| v0.3.0 | 配置模块化（config.py）；正则规则外置（rules.json）；一键撤销功能；命名模板拖拽排序 |

## 技术特点

- **轻量依赖**：仅 3 个第三方库，exe 37 MB（曾尝试引入 PyMuPDF/PaddleOCR，因体积过大移除）
- **云端 OCR 零依赖**：腾讯云 API 签名用 Python 标准库实现（hashlib + hmac + urllib），不引入 SDK
- **通信架构**：WebView + 本地 HTTP Server（ThreadingHTTPServer），前端 fetch 调用，事件轮询推送
- **保守识别策略**：非发票文件宁可跳过不误处理
- **数据安全**：原始文件不动，改名前可预览，执行后可撤销

## License

个人项目，保留所有权利。仅供学习交流使用。
