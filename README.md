# Visio SVG Stencil Importer

将本地 SVG 图标批量转换为可复用的 Microsoft Visio 形状库（`.vssx`）。每个 SVG 会成为一个可命名、可拖拽的主控形状，适用于绘制架构图、流程图和科研示意图时建立自己的图标库。

## 本 Skill 的作用

- 扫描一个本地 SVG 文件夹。
- 为每个图标创建一个 Visio 主控形状。
- 输出一个可在 Visio 中通过“更多形状 > 打开模具”加载的 `.vssx` 文件。
- 使用 PowerPoint 将 SVG 转为矢量 EMF，以兼容无法通过自动化接口直接导入 SVG 的 Visio 安装。

最终模具中的图标保持矢量缩放能力，但保存的是 EMF 表示，而不是可编辑的原始 SVG XML 路径。

## 使用方式

运行环境：Windows 桌面版 Microsoft Visio 和 PowerPoint。

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Import-SvgIconsToVisioStencil.ps1 `
  -IconDirectory "C:\path\to\icons" `
  -StencilPath "C:\path\to\output\svg-icons.vssx"
```

若输出文件已存在，脚本会停止。只有明确要替换已有模具时才添加 `-Force`。

## 版权与合规

本项目不提供、检索或下载任何第三方图标素材，也不会对素材网站进行抓包、接口逆向、绕过访问控制或反爬机制。此类操作可能违反网站服务条款、著作权规定或其他适用规则。

请自行前往图标作者官网、授权素材库或其他合法渠道，手动下载所需图标，并在使用前确认：

- 许可是否允许当前用途，特别是商业使用、修改和再分发。
- 是否需要署名、保留版权声明或购买授权。
- 不要将未经授权的素材随生成的 `.vssx` 或本项目再次分发。

脚本仅对用户已经合法取得并保存在本地的 SVG 文件进行格式转换与模具封装；素材的授权责任由使用者承担。
