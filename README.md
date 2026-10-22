**English** | [简体中文](README.zh-CN.md)

# AI Procurement Coordinator

A product and workflow design study for turning scattered procurement requirements into structured, traceable order information for suppliers and factories.

**Category:** Artificial Intelligence · Workflow Automation  
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

## Author and license

Author / maintainer: [https://github.com/OOOHGOD](https://github.com/OOOHGOD).

Eligible original contributions are under the [MIT License](LICENSE), permitting commercial use and modification with preservation of copyright and license notices in copies or substantial portions. [License scope](LICENSE_SCOPE.md) explains third-party exclusions and upstream terms.
