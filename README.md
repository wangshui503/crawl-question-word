# 极光题库导出脚本

`crawl_question.py` 按学科和章节获取当前登录账号可访问的题目，并导出为 Word（`.docx`）。默认筛选标签为“三年级”，每个章节保存一个文件。

## 文件与环境

将以下三个文件放在同一目录：

```text
crawl_question.py
requirements.txt
README.md
```

验证环境为 Windows、Python 3.12.14，依赖版本见 `requirements.txt`。建议在终端中运行，以便输入 token 和查看错误信息。

## 安装依赖

在 PowerShell 中进入文件所在目录，再执行：

```powershell
python -m pip install -r requirements.txt
```

依赖用途：`requests` 请求接口，`python-docx` 写入 `.docx` 文件。其余模块属于 Python 标准库，无需安装。

## 获取登录 token

1. 打开 <https://app.ji-guang.top/>，重新登录账号。
2. 按 **F12**，进入 **网络（Network）**，选择 **Fetch/XHR**。
3. 在网页中打开一个学科或章节，让浏览器发出新的请求。
4. 找到地址包含 `api.ji-guang.top/api/v1/user/` 的请求。
5. 打开 **标头（Headers）→ 请求标头（Request Headers）**，找到 `Authorization`。
6. 复制 `Bearer` 后面的完整 token。脚本也接受带 `Bearer ` 前缀的内容。

Token 是登录凭证，请勿公开分享或提交到代码仓库。过期后需要重新登录获取。

## 运行

```powershell
python .\crawl_question.py
```

按提示粘贴 token 并按回车。输入不会显示字符，这是正常行为。

默认输出目录为 `Path.home() / "Desktop" / "题库"`，在当前用户电脑上通常是 `C:\Users\ROG\Desktop\题库`。如果桌面已被重定向到 OneDrive，请通过 `--save-dir` 明确指定目录。

指定标签、输出位置和分页大小：

```powershell
python .\crawl_question.py --tag "三年级" --save-dir "D:\题库" --page-size 100
```

查看帮助：

```powershell
python .\crawl_question.py --help
```

| 参数 | 默认值 | 含义 |
| --- | --- | --- |
| `--tag` | `三年级` | 网站中的学科标签，需要与网站一致 |
| `--save-dir` | 当前用户的 `Desktop\题库` | Word 保存目录，不存在时自动创建 |
| `--page-size` | `100` | 每页请求数量，允许范围为 `1–10000` |
| `--no-explanation` | 默认不启用 | 不导出解析，只保留题目、选项和答案 |

### 通过环境变量提供 token

脚本读取 token 的顺序为：当前进程的 `JIGUANG_TOKEN` 环境变量 → 脚本顶部的 `TOKEN` → 交互式输入。

无需把 token 写入 PowerShell 历史记录，可以在当前终端中隐藏输入后设置环境变量：

```powershell
$secureToken = Read-Host "请输入 token" -AsSecureString
$env:JIGUANG_TOKEN = [System.Net.NetworkCredential]::new("", $secureToken).Password
try {
    python .\crawl_question.py
} finally {
    Remove-Item Env:\JIGUANG_TOKEN -ErrorAction SilentlyContinue
    Remove-Variable secureToken -ErrorAction SilentlyContinue
}
```

环境变量中的 token 以明文提供给 Python 进程，上面的命令会在运行结束后移除它。已保存到其他程序的安全存储或内存中的 token 不会自动传给本脚本；本脚本未实现安全存储读取功能。非交互式运行时，需要提前提供 `JIGUANG_TOKEN` 或配置 `TOKEN`。

## 导出内容与结果

每个章节生成一份 A4 Word 文档，包含学科与章节标题、题目数量、连续题号、题干、各项选项和答案。默认附上已有解析，优先使用 `ai_explanation`，为空时使用 `explanation`；没有解析时省略该段，不会自动生成解析。答案为空时标明“未提供”。

选项按键排序，支持 A–E 之外的选项；空选项省略。HTML 标签转为纯文本，保留段落和换行。嵌套列表逐项换行，对象按键值显示。正文使用宋体，标题使用黑体，并附页码。

只导出题目、选项和答案：

```powershell
python .\crawl_question.py --no-explanation
```

图片和附件不会下载或嵌入；题干 HTML 中的图片以文字占位提示。HTML 富文本样式和公式排版不会转为原生 Word 公式，需要时可在 Word 中进一步编辑。

文件名格式为 `学科-章节-章节ID.docx`。文件名中的非法字符会被替换；章节ID用于区分同名章节。重新运行会覆盖相同路径的结果。

脚本对列表分页并去重。网络异常、HTTP 429 和服务器 5xx 最多尝试三次；401 登录失效会立即停止。题目请求之间默认等待 0.15 秒，该间隔目前只能在代码中调整。

普通题目错误或分页错误发生时，已获取的章节题目会保存为 `学科-章节-章节ID-部分结果.docx`。没有成功获取的题目时，不创建空 Word 文档。若发生登录失效或手动中断，当前尚未写入的章节不会保存，之前已写入的文件仍保留。

退出码 `0` 表示至少生成一个文件且没有记录失败；`1` 表示爬取失败、结果不完整或没有生成文件；参数使用错误返回 `2`；用户中断返回 `130`。脚本不支持断点续传，重试会重新获取。

## 常见问题

| 提示或现象 | 处理方法 |
| --- | --- |
| `ModuleNotFoundError` / 缺少依赖 | 用运行脚本的同一个 Python 执行 `python -m pip install -r requirements.txt` |
| token 已过期 / HTTP 401 | 重新登录获取 token；如已设置 `JIGUANG_TOKEN`，也要同步更新，因为它优先于代码中的 `TOKEN` |
| HTTP 403 | 检查当前账号是否拥有该学科或题目的访问权限 |
| 返回非 JSON 内容 | 检查接口根地址，正确值为 `https://api.ji-guang.top/api/v1`，不能使用带 `#/home` 的网页地址 |
| 没有学科 | 检查 `--tag` 是否与网站标签一致，以及当前账号可访问的内容 |
| 缺少字段 / options 格式异常 / 分页重复 | 接口响应可能发生变化，需要结合新的响应调整解析，脚本会报告错误 |
| HTTPS 证书校验失败 | 检查系统时间、证书配置和网络环境；脚本保持证书校验开启 |
| 连接超时 / 网络请求失败 | 检查网络连接；脚本设置了 `session.trust_env = False`，不会读取系统代理和相关环境配置 |
| Word 保存失败 | 关闭正在 Word 中打开的同名文件，检查目录权限或通过 `--save-dir` 更换目录 |
| 中文显示乱码 | 可尝试在 PowerShell 中设置 `$env:PYTHONIOENCODING = "utf-8"` 后再运行 |

## 验证范围

Word 导出版已通过 11 项本地测试，覆盖过期 token、错误 API 地址、401、不合法 JSON、分页重复、Word 导出、失败章节不生成空文件、HTML 清理、嵌套选项、关闭解析和部分结果标识。已用构造题目验证文档生成和内容读取；这些样例不是在线获取的真实题目。当前验证环境缺少 LibreOffice，尚未完成页面渲染和视觉排版检查。已在线核实 API 根地址，并观察到旧 token 在学科详情接口返回 401；使用有效新 token 的完整在线爬取仍需验证。
