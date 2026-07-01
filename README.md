# MailBatch 批量邮件工具

基于 Python + tkinter 的 Windows 桌面应用：从 Excel 读取名单，用 `{{列名}}` 模板批量发送带附件邮件。

## 功能

- 选择 Excel（单 sheet）、附件、SMTP 配置（常用邮箱预设 + 自定义）、邮件主题
- 内嵌模板编辑器：打开、编辑、保存；关闭时随配置一并保存
- 字段列表双击插入 `{{column}}` 占位符
- **邮件预览**：按行号预览收件人、抄送、主题、正文和附件
- **抄送**：从 Excel 抄送列解析（支持 `抄送`、`抄送人`、`CC` 等列名，多个邮箱用逗号分隔）
- **配置记忆**：关闭时自动保存上次配置；支持手动保存、导入、导出 JSON 配置；支持一键清空配置
- 发送前可校验必填项
- 后台线程发送，界面显示进度与日志
- 支持将旧模板 `【列名】` 手动改为 `{{列名}}`

## 环境要求

- Windows 10/11
- Python 3.10+（必须真实安装，不能只有 Microsoft Store 占位符）

## 先安装 Python

如果运行脚本时出现 `Python was not found`，说明本机尚未安装 Python，或 Windows 的「应用执行别名」拦截了命令。

1. 安装 Python：
   - 官网：https://www.python.org/downloads/windows/
   - 安装时勾选 **Add python.exe to PATH**
   - 或使用命令：`winget install Python.Python.3.12`
2. 关闭 Microsoft Store 别名：
   - **设置 → 应用 → 高级应用设置 → 应用执行别名**
   - 关闭 `python.exe` 和 `python3.exe` 两项
3. 重新打开终端，确认可用：

```powershell
python --version
```

应显示类似 `Python 3.12.x`。若仍失败，可尝试：

```powershell
py -3 --version
```

## 安装与运行

```powershell
cd C:\Users\abc\Projects\mailBatch

py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt

# 复制并编辑 SMTP 配置
copy .env.example .env

# 启动桌面应用
python app.py
```

## 使用说明

1. **Excel**：选择名单文件，应用读取第一个 sheet 的表头和数据行。收件邮箱列名支持：`邮箱`、`邮件`、`Email`、`email`。抄送列支持：`抄送`、`抄送人`、`抄送邮箱`、`CC` 等；多个抄送地址可用中英文逗号或分号分隔。
2. **模板**：在编辑器中直接编辑，占位符格式为 `{{列名}}`。可用编辑器上方按钮打开/保存文件，或将 `【列名】` 转为 `{{列名}}`。
3. **附件**：添加需要随每封邮件发送的文件（如 PNG 图片）。
4. **SMTP**：选择 **邮箱预设**（自动识别 / QQ / 163 / Gmail / Outlook / Apple iCloud 等），或选 **自定义** 手动填写。在 **自动识别** 模式下，输入用户名或发件人邮箱后会按域名自动填充 SMTP 主机、端口和加密方式。
5. **预览**：选择预览行号，点击「预览」查看该行的收件人、抄送、主题、正文和附件。
6. **校验**：检查主题、模板、邮箱列、附件等必填项。
7. **发送**：确认后逐行发送，日志区显示成功/失败详情。
8. **配置**：顶部菜单栏可直接 **保存配置 / 导出配置 / 导入配置 / 清空配置**；关闭应用时自动保存到 `last_config.json`。

### 素材文件

项目目录可放置：

- `module1.txt` — 邮件正文模板（旧格式可在编辑器中点击「转换 【】→{{}}」）
- `新员工入职指引-邮件名单假模板.xlsx` — 名单
- `-1556627939.png`、`300420580.png`、`881290838.png` — 附件

启动时会自动尝试加载同目录下的 xlsx 与默认附件。

## 打包 exe

先确保 `python --version` 或 `py -3 --version` 可用，再执行：

```powershell
.\build_exe.ps1
```

若打包失败，脚本会明确报错，不会再误报 `Build complete`。

或：

```cmd
build_exe.bat
```

产物为单文件 `dist\MailBatch.exe`（基于 tkinter，体积远小于 PySide6 版本），可直接通过微信等方式分发。

将 `.env`、模板 txt、Excel、附件放在 exe 同目录即可使用。目标电脑无需安装 Python；若启动时报 DLL 相关错误，请安装 [Visual C++ Redistributable x64](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist)。

## 项目结构

| 文件 | 说明 |
|------|------|
| `app.py` | tkinter 桌面应用入口 |
| `mail_core.py` | Excel 读取、模板渲染、SMTP 批量发送 |
| `requirements.txt` | Python 依赖 |
| `.env.example` | SMTP 配置示例 |
| `build_exe.ps1` / `build_exe.bat` | PyInstaller 打包脚本 |

## mail_core 公开 API（供二次开发）

```python
from mail_core import (
    load_excel,           # (path) -> ExcelData
    render_template,      # (text, row_dict) -> str
    convert_legacy_placeholders,
    load_smtp_config,
    validate_send_request,
    send_batch,
)
```
