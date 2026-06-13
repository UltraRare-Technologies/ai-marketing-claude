# AI Marketing Suite

This workspace contains a comprehensive marketing analysis and automation
skill system for Claude Code. It provides 14 marketing skills and an
orchestrator that routes `/market` commands.

## Purpose

Used by the CMO role to run marketing audits, generate copy, build email
sequences, create content calendars, analyze competitors, and produce
client-ready reports — all from the terminal.

Designed for sales prospecting (audit a prospect's site before a call),
client deliverables (PDF reports), and agency workflow automation.

## Architecture

```
ai-marketing-claude/
├── market/SKILL.md              # Main orchestrator — routes all /market commands
├── skills/                      # 14 sub-skills (one per command)
├── agents/                      # 5 parallel subagents for full audit
├── scripts/                     # Python utility scripts
├── templates/                   # Marketing templates (email, proposal, calendar)
├── install.sh / uninstall.sh    # Skill installer/uninstaller
└── requirements.txt             # Python dependencies (reportlab for PDF)
```

## Skills (14)

Located at `skills/`:

| Skill | Command | What It Does |
|-------|---------|-------------|
| market-audit | `/market audit <url>` | Full 6-dimension audit with 5 parallel agents |
| market-seo | `/market seo <url>` | SEO content audit |
| market-copy | `/market copy <url>` | Optimized copy with before/after examples |
| market-brand | `/market brand <url>` | Brand voice analysis and guidelines |
| market-competitors | `/market competitors <url>` | Competitive intelligence report |
| market-funnel | `/market funnel <url>` | Sales funnel analysis and optimization |
| market-landing | `/market landing <url>` | Landing page CRO analysis |
| market-emails | `/market emails <topic>` | Complete email sequences |
| market-social | `/market social <topic>` | 30-day social media content calendar |
| market-ads | `/market ads <url>` | Ad creative and copy for all platforms |
| market-launch | `/market launch <product>` | Product launch playbook |
| market-proposal | `/market proposal <client>` | Client proposal generator |
| market-report | `/market report <url>` | Full marketing report (Markdown) |
| market-report-pdf | `/market report-pdf <url>` | Professional marketing report (PDF) |

## Orchestrator

`market/SKILL.md` is the entry point. It reads the command after `/market`
and routes to the appropriate sub-skill. Do not invoke sub-skills directly —
always go through the orchestrator.

## Agents

5 parallel subagents run during a full audit (`/market audit`):
- `market-content.md` — Content & messaging
- `market-conversion.md` — CRO & funnel
- `market-competitive.md` — Competitive positioning
- `market-technical.md` — Technical SEO & tracking
- `market-strategy.md` — Brand, pricing & growth

## Installation

Skills are installed to `~/.claude/skills/market*/` via `./install.sh`.
PDF support requires `pip install reportlab`.

## Used By

CMO role — for prospect audits, client deliverables, and marketing collateral.
