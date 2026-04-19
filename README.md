# 泽PPT备份助手 / Ze PPT Backup Assistant

泽PPT备份助手是一款用于课堂、会议、培训等场景的本地文档自动备份工具。

程序启动后会在后台运行，自动检测当前电脑打开的 PowerPoint、WPS 演示、Word/WPS 文字以及常见 PDF 文件，并将文件复制到本地备份目录。自动备份过程默认不弹窗，适合需要减少手动复制文件的电脑环境。

Ze PPT Backup Assistant is a lightweight Windows tool for automatically backing up PPT, Word, and PDF files opened during classes, meetings, or training sessions.

> 使用提醒：请在本人或已授权的电脑上使用本软件。因未获授权使用、误操作、数据丢失、隐私纠纷等造成的后果，由使用者自行承担。

## 当前版本 / Version

v5.0

## 主要功能 / Features

- 自动检测并备份 PPT / Word / PDF
- 支持 Microsoft PowerPoint、Microsoft Word、WPS 演示、WPS 文字
- 支持常见 PDF 阅读器检测，兼容部分浏览器 PDF 标题识别
- 默认备份到 D/E 盘，D/E 不可用时自动回退到 C 盘或文档目录
- 按日期分类保存，重名文件自动编号
- 使用路径、修改时间、内容哈希等规则减少重复备份
- 右下角托盘后台运行，双击打开主界面
- 主界面查看最近备份、打开备份目录、打开日志
- 支持删除选中记录、清理失效记录、清空记录、清空日志
- 支持一键体检，便于售后排查
- 支持磁盘空间保护，可限制每个备份位置最多占用空间
- 支持自定义备份格式：PPT / Word / PDF
- 支持自定义备份位置
- 首次运行显示免责声明
- 关于/反馈窗口支持复制反馈邮箱

## 发布包 / Release Package

正式给普通客户使用时，建议发送：

`发布包-releases/v5.0/泽PPT标准备份版.zip`

标准备份版用于本地自动备份和记录查看。

For normal customer delivery, use the standard backup package above. It is focused on local file backup and backup record management.

## 默认备份目录 / Default Backup Folder

默认优先保存到：

```text
D:\泽宁PPPPPPPPTTTT备份\
E:\泽宁PPPPPPPPTTTT备份\
```

如果 D/E 盘不可用，会自动保存到：

```text
C:\泽宁PPPPPPPPTTTT备份\
```

如果 C 盘根目录也不可写，会保存到当前用户文档目录。

## 仓库结构 / Repository Structure

```text
源码-src/
  ZeBackupAssistant.cs          源码 / Source code

文档-docs/
  详细功能完整说明.txt
  客户交付话术.txt
  售后排查说明.txt

发布包-releases/v5.0/
  泽PPT标准备份版.zip

CHANGELOG.txt
README.md
```

## 编译说明 / Build

本项目为 C# WinForms 单文件程序，可使用 .NET Framework 4.x 自带的 `csc.exe` 编译。

参考命令：

```powershell
$src = "源码-src\ZeBackupAssistant.cs"
$out = "泽PPT备份助手.exe"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
& $csc /nologo /codepage:65001 /target:winexe /platform:anycpu "/out:$out" `
  /reference:Microsoft.CSharp.dll `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.Management.dll `
  $src
```

## 注意事项 / Notes

- 不要提交客户备份文件、日志、去重索引或个人隐私文件。
- 正式对外销售时，优先分发标准备份版。

## 反馈 / Feedback

有建议或者问题，欢迎反馈到：

`zeningshuyi@gmail.com`
