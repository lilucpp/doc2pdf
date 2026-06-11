# wps2pdf

> 命令行工具，将 WPS / Microsoft Office 文档（.doc/.docx）转换为 PDF。

[![Platform](https://img.shields.io/badge/Platform-Windows-blue)](https://github.com/lilucpp/doc2pdf)
[![Language](https://img.shields.io/badge/Language-C++11-00599C)](https://isocpp.org/)
[![Build](https://img.shields.io/badge/Build-CMake_%2F_MSVC-6DB33F)](https://cmake.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## 特性

- **无需打印机驱动** — 通过 COM 自动化调用 WPS / Office 原生导出 API
- **自动检测** — 优先使用 WPS，回退到 Microsoft Word
- **支持页面范围** — 可指定起始/结束页码
- **轻量** — 单 exe，无外部运行时依赖（静态链接 /MT）
- **兼容 32/64 位系统** — 推荐 32 位构建以兼容大多数 WPS/Office 安装

## 实现原理

通过 **COM 自动化** 调用应用程序的 `ExportAsFixedFormat` API，无需任何中间件或虚拟打印机。

## 兼容性

| 应用程序 | ProgID | 要求 |
|----------|--------|------|
| **WPS Office** (推荐) | `KWps.Application` | WPS Office 已安装 |
| **Microsoft Office Word 2007** | `Word.Application` | 需 **SP2** 或安装 [Save as PDF 插件](http://www.microsoft.com/downloads/details.aspx?familyid=4d951911-3e7e-4ae6-b059-a2e79ed87041) |
| **Microsoft Office Word 2010+** | `Word.Application` | 原生支持，无需额外组件 |
| **Microsoft Office 365** | `Word.Application` | 原生支持 |

> 程序启动顺序：先尝试 WPS → 失败则尝试 Office Word → 都失败则报错退出。

## 使用

```
wps2pdf <输入文件> [-o 输出文件] [-f 起始页] [-t 结束页]
```

| 参数 | 说明 |
|------|------|
| `-o` / `--output` | 输出 PDF 路径（默认：输入文件同名.pdf） |
| `-f` / `--from` | 起始页码（默认：第 1 页） |
| `-t` / `--to` | 结束页码（默认：最后一页，需同时指定 `-f`） |

### 示例

```powershell
wps2pdf 1.docx                      # 全部页面
wps2pdf 1.docx -f 2 -t 5           # 第 2-5 页
wps2pdf 1.docx -f 3                # 从第 3 页到末尾
wps2pdf 1.docx -o D:\out.pdf       # 指定输出路径
```

## 编译

### 要求

- Visual Studio 2022 Build Tools（或完整版 VS）
- CMake 3.10+

### 构建（32 位，静态链接 /MT）

```powershell
cmake -G "Visual Studio 17 2022" -A Win32 -B build
cmake --build build --config Release
```

产物在 `build\Release\wps2pdf.exe`，不依赖 VC 运行时库。

### 架构说明

工具编译为 32 位，通过 COM 调用 WPS/Office。32 位工具只能调用 32 位的 WPS/Office（64 位系统上这是最常见的安装方式）。

若需支持纯 64 位 Office，请用 `-A x64` 额外编译 64 位版本：

```powershell
cmake -G "Visual Studio 17 2022" -A x64 -B build64
cmake --build build64 --config Release
```

## 相关关键词

`doc2pdf` `docx2pdf` `wps2pdf` `word转pdf` `wps转pdf` `office转pdf` `命令行pdf转换`

## 许可证

[MIT](LICENSE)
