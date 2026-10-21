# 可编辑 Visio 图重建技能

配合支持 SKILL.md 的编程代理，把截图、科研图、热图和架构图重建成可编辑的 Microsoft Visio 图元，再分别检查文件结构、实际渲染和编辑行为。

[English](README.md) · [技能指令](SKILL.md) · [验证说明](references/verification.md) · [示例](examples/README.md)

## 能力范围

这个项目提供重建工作流、实现参考、VSDX 静态校验器和 Visio 导出脚本。图中的文字、框、线和色块应尽量成为独立对象。

它需要代理观察源图并编写重建逻辑，目前不是通用的一键图片转换软件。示例脚本用于展示可编辑输出，不执行任意图片识别；采用直接 XML 生成路线时仍需已知有效的 VSDX 模板。

## 安装与使用

在仓库根目录运行：

```powershell
python scripts/install_skill.py
```

安装到 CODEX_HOME 指定目录下的 skills，未设置时使用 ~/.codex/skills。已有同名技能时会停止，不会覆盖。安装后若未发现技能，可重新打开代理会话。

示例请求：

> 使用 $editable-visio-recreation 将这张图重建成可编辑 .vsdx。保留原图；分别交付 Visio 文件、实际导出的 PNG、静态校验结果和代表性缩放/连线测试结果。

也可以只运行静态校验：

```powershell
python scripts/validate_vsdx.py diagram.vsdx --profile source --report check.json
```

Windows 上使用桌面 Visio 导出真实预览：

```powershell
.\scripts\render_visio.ps1 -InputPath .\diagram.vsdx -OutputDirectory .\new-preview-folder -Dpi 300
```

输出目录必须是新目录。生成报告可能含本机绝对路径，公开前应检查。

## 环境与预设

- 静态校验和测试：Python 3.10+，仅标准库，可在 Windows、Linux、macOS 运行。
- 真实预览和示例生成：Windows、PowerShell、已安装的桌面 Microsoft Visio、交互式登录会话。
- 字体须在运行机器上可用。Pillow 只用于可选的取色或布局草图，不是静态校验器依赖。
- 默认 source 预设不限制页面大小和字体。
- 可选 journal-small 预设：最大 80×60 mm、6 pt、中文 SimSun、西文 Times New Roman；也可传入自定义 JSON。

## 如何理解验证结果

三个层次分别记录：

1. 静态检查：ZIP/XML、关系目标、页面和图元、部分文字格式与固定几何警告。
2. 应用渲染：Visio 实际打开最终文件并导出，再检查文字、连线和布局。
3. 编辑行为：在可丢弃副本上测试代表性图元缩放和连接线跟随。

static_ok=true 只代表已实现的静态检查通过。它不保证 Visio 能打开、不保证版式一致，也不保证所有图元可正确编辑。校验器不完整解析主控形状、样式继承、公式或字体替换；--allow-media 也不能证明混合文件具有有用的可编辑内容。

## 测试与示例

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
```

测试夹具是验证逻辑用的简化 OPC 包，不是完整的 Visio 兼容性证据。实际文件示例及复现方式见 [examples/README.md](examples/README.md)。

欢迎贡献你有权公开的示例、可复现问题和有意义的验证测试。请说明问题属于静态检查、实际渲染还是交互编辑。

## 许可证

[MIT](LICENSE)。仓库自建示例同样适用；Visio 软件和第三方字体遵循各自许可证。

