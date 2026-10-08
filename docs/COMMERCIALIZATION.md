# DocViewKit 开源与服务交付

> 执行基线：自有源码采用 Apache-2.0；商业收入来自支持、定制和企业交付。

DocViewKit 是浏览器本地、只读的文档预览组件和 Engine SDK。
全部已实现格式、Engine API、UI 插槽、CSS parts、极简模式和文字水印
均按 Apache-2.0 提供，不要求运行时许可证、域名登记或强制品牌展示。
第三方代码、字体和测试材料保留原版权与许可证；DocViewKit 商标单独管理。

## 付费服务

- 支持：接入指导、问题复现、兼容性诊断、升级协助与约定响应时间。
- 定制：现有格式保真修复、专项性能优化、按格式裁剪和业务集成。
- 企业交付：部署验证、可复核构建、验收材料、培训与 SLA。

服务按实际需求、交付范围、验收 Case、周期和 SLA 报价，不预设软件使用费。
公开问题通过 GitHub Issues；私密文件通过双方约定的渠道提供。
服务需求直接通过邮件 novalag778@gmail.com 联系；官网无需注册、登录或应用登记。
旧账户与授权申请数据保留归档，官网不再读取或修改。
现有客户的服务义务按原合同履行，新版本的软件权利遵循 Apache-2.0。

## 工程与发布

核心源码位于 [docviewkit/viewer](https://github.com/docviewkit/viewer)，
产生 Viewer、Engine、格式包和 Pages 演示；保持现有公开 Viewer 入口兼容，
并通过 @docviewkit/viewer/engine 提供 Engine API。
官网、在线文档和 Demo 位于 [docviewkit/website](https://github.com/docviewkit/website)，
使用固定版本的 npm 核心库，独立测试和部署。
根工程 @docviewkit/sdk 的 private 标记仅防止误向 npm 发布构建工程，
不限制源码使用权。第三方许可、固定字体版本和哈希门禁继续生效。
公开 Git 历史及外来测试文件前，须完成凭据、隐私、来源和再分发权核验。

每个变更以修改前失败、修改后通过的真实 Case 为准；解析和运行时修改需
验证真实渲染及 Chrome、Firefox、Safari 内核行为。保留近似保真与原生视觉
认证边界，测试通过不自动等于已发布或已部署。
