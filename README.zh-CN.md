<div align="center">

# AI 采购协调助手

从零散需求到可追溯的采购工作流。

[English](README.md) · **简体中文**

![AI Agents](https://img.shields.io/badge/AI%20Agents-334155?style=flat-square) ![Workflow Automation](https://img.shields.io/badge/Workflow%20Automation-334155?style=flat-square) ![HTML](https://img.shields.io/badge/HTML-E34F26?style=flat-square) [![MIT License](https://img.shields.io/badge/License-MIT-2563EB?style=flat-square)](LICENSE)

</div>

<p align="center">
<a href="#项目概览">项目概览</a> · <a href="#方案流程">方案流程</a> · <a href="#作者与许可">作者与许可</a>
</p>

---

> **项目定位** · 方案与可行性研究


## 项目概览

围绕家装、酒店与工程采购场景，研究如何把分散的消息、PDF、图片和 Excel 整理为可追溯、可审核的订单信息。

> **当前阶段**：需求拆解与可行性研究。仓库尚未包含可运行的 Agent、后端服务或生产集成。

## 方案流程

| 阶段 | 设计目标 |
| --- | --- |
| 资料归集 | 将邮件、聊天记录、图片和表格归入同一项目 |
| 信息抽取 | 使用 OCR 与 LLM 提取品名、数量、尺寸、材质和供应商字段 |
| 版本与冲突检查 | 识别变更、字段缺失和需求冲突，保留来源 |
| 分类与审核 | 按项目、房间、品类和供应商组织信息，交由人工确认 |
| 单据输出 | 生成便于客户确认和工厂执行的订单文件 |

以上为拟议流程，不表示功能已经实现。

## 阅读入口

| 材料 | 内容 |
| --- | --- |
| [可行性报告](feasibility-report.html) | 浏览器可打开的方案与可行性分析 |
| [工作流任务拆解](家装下单工作流任务拆解文字稿.md) | 用户角色、字段、阶段、人工确认与实施计划 |
| `project background.txt` | 原始背景记录，不是运行依赖 |

## 本地阅读

```bash
git clone https://github.com/OOOHGOD/AI-Procurement-Coordinator.git
cd AI-Procurement-Coordinator
```

用浏览器打开 `feasibility-report.html`，或用编辑器阅读 Markdown。阅读材料不需要安装依赖，也不需要 API Key。

## 后续实现方向

先用样本文档验证需求，定义订单 Schema，搭建最小的信息抽取与审核流程；评估字段准确性、变更追踪和冲突识别后，再逐步增加外部集成。

## 作者与许可

作者 / 维护者：[OOOHGOD](https://github.com/OOOHGOD) · 原创贡献：[MIT License](LICENSE)

允许商用和修改；分发软件副本或实质部分时，须保留版权与许可声明。第三方代码、数据、模型和媒体素材保留各自权利，详见 [许可适用范围](LICENSE_SCOPE.md)。
