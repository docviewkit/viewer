# DocViewKit Viewer

[English](https://github.com/docviewkit/viewer/blob/main/README.md) · [简体中文](https://github.com/docviewkit/viewer/blob/main/README.zh-CN.md)

**轻量预览，快速集成，融入每一种业务。**

让用户直接在你的应用中阅读 Word、Excel、PowerPoint、PDF、OFD 等文档。DocViewKit 将文档预览、搜索和原文定位融入 OA 附件、审批流程、CRM/ERP、网盘与企业应用。文档在前端本地解析和渲染，无需部署文档转换服务器。

[官网](https://docviewkit.com/zh-cn/) · [用自己的文件体验](https://docviewkit.com/zh-cn/demo/) · [使用文档](https://docviewkit.com/docs/quickstart/) · [npm](https://www.npmjs.com/package/@docviewkit/viewer)

## 为什么选择 DocViewKit

- **快速集成。** 使用现成的 Viewer 即可开始，无需从头搭建预览界面。页面导航、搜索、缩放和打印已准备好。
- **文件在本地处理。** 解析、渲染和搜索都在你的应用前端运行。DocViewKit 不会将文档上传到转换服务，附件访问与权限仍由你的应用掌控。
- **轻量与高性能。** 精简核心与按需加载的格式包，让未使用的解析器不占用首屏加载。页面、幻灯片和表格视口按需渲染。
- **多种格式，一致体验。** 现代与旧版 Office、WPS、PDF、OFD、OpenDocument、iWork 等支持格式共用同一个预览组件。
- **从结果回到原文。** 搜索并跳转到相关内容，将应用中的搜索结果或 AI 引用关联到文档页面、对象和区域，方便用户核查来源。
- **融入你的产品。** 按产品需要调整界面、命令、主题和文字水印；需要自定义渲染或访问原文对象时，使用 Engine API。
- **开源，免费使用。** Viewer、Engine 和全部已实现的格式采用 Apache-2.0。付费服务围绕支持、定制和企业交付展开。

## 适用场景

- **OA、审批与内部工具：** 在处理任务的同时阅读相关附件。
- **CRM、ERP 与客户门户：** 在现有业务流程中查看合同、记录和业务文档。
- **网盘、企业 SaaS 与 AI 知识库：** 预览文件，将搜索结果或引用定位到原文。
- **法务审阅、审计与政务系统：** 使用自己的访问控制查看文档、定位依据。
- **教育与培训：** 在应用内阅读课程材料、演示文稿与作业。

## 用户可以做什么

通过导航、缩略图、缩放与全屏浏览页面、幻灯片和工作表。搜索文档文字，选择并复制内容，跳转到相关位置。为页面、缩略图和打印输出添加文字水印。可以使用完整 Viewer 界面，也可以选择适合产品的精简文档视图。

Viewer 提供只读预览。文档存储、授权与业务流程由你的应用负责，DocViewKit 提供阅读与原文定位能力。

## 支持格式

| 文档类型 | 示例 |
| --- | --- |
| 现代 Microsoft Office | Word `.docx`、Excel `.xlsx`、PowerPoint `.pptx`，以及已支持的宏启用、模板和放映变体 |
| 旧版 Microsoft Office | Word `.doc`、Excel `.xls`、PowerPoint `.ppt` |
| WPS Office | `.wps`、`.et`、`.dps` |
| 固定版式文档 | PDF、OFD 1.0/1.1、XPS 和 OpenXPS |
| OpenDocument | `.odt`、`.ods`、`.odp`、`.odg`，以及已支持的模板与扁平 XML 变体 |
| Apple iWork | `.pages`、`.numbers`、`.key` 单文件包 |
| 文本与表格数据 | CSV 和 RTF |

不同格式和文档的兼容性和渲染深度存在差异。请用真实文件体验[在线 Demo](https://docviewkit.com/zh-cn/demo/)，并查看[支持格式与限制](https://docviewkit.com/docs/supported-formats/)。部分旧版 iWork 路径仍返回内嵌预览，这是尚待补齐的实现缺口，不能据此判断原生内容渲染或保真度。

## 快速开始

安装 Viewer：

```sh
npm install @docviewkit/viewer
```

将组件放入页面：

```html
<docviewkit-viewer></docviewkit-viewer>
```

在客户端代码中导入组件，按需启用可选格式，然后打开由你的应用提供的 `File`：

```js
import "@docviewkit/viewer";

const viewer = document.querySelector("docviewkit-viewer");
viewer.config = {
  engine: {
    formatPack: () => import("@docviewkit/viewer/extended-formats")
      .then(({ extendedFormatPack }) => extendedFormatPack),
  },
};
await viewer.open(file);
```

组件只加载当前文档需要的可选格式模块。仅使用 OFD 的应用可选择 `ofd-formats`。部署时保留配套的 Worker、Wasm、字体和编解码资源；资源托管与私有附件接入方式见[集成指南](https://docviewkit.com/docs/quickstart/)。

## 适配你的前端

JavaScript、React、Vue 和 Angular 应用可以使用同一个 Viewer，支持 Chrome、Edge、Firefox 和 Safari。兼容的 WebView 也可用于 Electron、Tauri、Ionic 和 Capacitor。React Native 与 Flutter 应用可以通过具备所需浏览器能力的 WebView 嵌入 Viewer，具体要求见[运行环境兼容性](https://docviewkit.com/docs/browser-compatibility/)。

[React 与 Next.js](https://docviewkit.com/react-office-viewer/) · [Vue](https://docviewkit.com/vue-office-viewer/) · [Angular](https://docviewkit.com/angular-office-viewer/)

## 开源与专业服务

DocViewKit Viewer 和 Engine 采用 [Apache-2.0](LICENSE)。全部已实现的格式、Engine API、界面定制、精简模式与文字水印均可使用，无需运行时许可证、域名登记或强制品牌展示。第三方组件和字体保留各自的许可证；再分发相关材料时，请保留 [LICENSE](LICENSE)、[NOTICE](NOTICE) 和 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。这些条款适用于从当前源码构建的版本，旧版已发布包继续适用其随附许可证。商标权另行约定。

如需集成协助、文档兼容性优化、性能调优、自定义界面或企业部署，请联系 [novalag778@gmail.com](mailto:novalag778@gmail.com)。支持、定制和企业交付的范围、验收案例、交付时间与 SLA 按服务约定确定。

## 了解更多

- [Viewer API](https://docviewkit.com/docs/viewer-api/) 与[原文定位](https://docviewkit.com/docs/source-location/)
- [性能指南](https://docviewkit.com/docs/performance/) 与[精度和保真度](https://docviewkit.com/docs/accuracy-fidelity/)
- [集成与架构](https://github.com/docviewkit/viewer/blob/main/docs/ARCHITECTURE.md#integration-reference)
- [构建与贡献](https://github.com/docviewkit/viewer/blob/main/docs/RELEASING.md#local-development)
- [源码与版本](https://github.com/docviewkit/viewer) · [反馈问题](https://github.com/docviewkit/viewer/issues) · [安全政策](https://github.com/docviewkit/viewer/blob/main/docs/SECURITY.md)

采用 Viewer 前，请验证自己的真实文档。需要支持时，机密文件仅通过双方约定的私密渠道提供。
