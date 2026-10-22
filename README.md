<div align="center">

# AI Procurement Coordinator

From scattered requirements to traceable procurement workflows.

**English** · [简体中文](README.zh-CN.md)

![AI Agents](https://img.shields.io/badge/AI%20Agents-334155?style=flat-square) ![Workflow Automation](https://img.shields.io/badge/Workflow%20Automation-334155?style=flat-square) ![HTML](https://img.shields.io/badge/HTML-E34F26?style=flat-square) [![MIT License](https://img.shields.io/badge/License-MIT-2563EB?style=flat-square)](LICENSE)

</div>

<p align="center">
<a href="#overview">Overview</a> · <a href="#proposed-workflow">Proposed workflow</a> · <a href="#repository-guide">Repository guide</a> · <a href="#author--license">Author &amp; License</a>
</p>

---

> **Project focus** · Workflow Design


## Overview

A product and workflow design study for turning scattered procurement requirements into structured, traceable order information for suppliers and factories.

**Status:** Planning and feasibility study. This checkout contains documents and an HTML report; an operational agent, backend, and production integrations are not included.

## Proposed workflow

1. Collect source messages, images, PDFs, and spreadsheets.
2. Extract product, quantity, dimension, material, and supplier fields.
3. Track changes and flag conflicting or missing requirements.
4. Group confirmed items by project, room, category, or supplier.
5. Generate order documents for human review before dispatch.

These steps describe the intended system, rather than implemented integrations.

## Repository guide

| Material | Purpose |
| --- | --- |
| [Feasibility report](feasibility-report.html) | Browser-readable analysis of the proposed system |
| [Workflow breakdown](家装下单工作流任务拆解文字稿.md) | Chinese requirements, roles, stages, and implementation considerations |
| `project background.txt` | Original project context, not a runtime dependency |

## Read locally

```bash
git clone https://github.com/OOOHGOD/AI-Procurement-Coordinator.git
cd AI-Procurement-Coordinator
```

Open `feasibility-report.html` in a browser or read the Markdown document in an editor. No dependency installation or API key is required. The report and detailed requirements are in Chinese; this README is the English entry point.

## Next steps

Validate requirements with sample documents, define an order schema, build an extraction-and-review prototype, and evaluate field accuracy and change tracking before adding dispatch integrations.

## Author & License

Author / maintainer: [OOOHGOD](https://github.com/OOOHGOD) · Original contributions: [MIT License](LICENSE)

Commercial use and modification are permitted with preservation of copyright and license notices in copies or substantial portions. Third-party code, datasets, models, and media retain their own terms; see [license scope](LICENSE_SCOPE.md).
