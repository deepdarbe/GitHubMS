# connect-apps-plugin

A [Claude Code](https://code.claude.com/docs/en/plugins) plugin that helps you
connect to and troubleshoot the cloud apps used in this repository: **Azure**,
**Microsoft 365** (Graph / Exchange / SharePoint / Teams), and **GitHub**.

## Usage

Load the plugin by pointing Claude Code at this directory:

```bash
claude --plugin-dir ./connect-apps-plugin
```

## Commands

| Command | Description |
| --- | --- |
| `/connect-apps-plugin:azure [subscription]` | Connect to Azure and show the active context/subscriptions. Prefers the modern `Az` module / Azure CLI over the retired `AzureRM`. |
| `/connect-apps-plugin:m365 [service]` | Connect to a Microsoft 365 service (`graph`, `exchange`, `sharepoint`, `teams`). |
| `/connect-apps-plugin:github [org/repo]` | Authenticate to GitHub and verify repo access. |
| `/connect-apps-plugin:status` | Show connection status across all supported apps. |

## Agent

- **connection-helper** — diagnoses and fixes authentication/connection
  problems (expired tokens, wrong tenant/subscription, missing modules,
  insufficient scopes). Invoked automatically when a connect command fails, or
  explicitly via the Task tool.

## Layout

```
connect-apps-plugin/
├── .claude-plugin/
│   └── plugin.json          # plugin manifest
├── commands/                # slash commands
│   ├── azure.md
│   ├── m365.md
│   ├── github.md
│   └── status.md
├── agents/
│   └── connection-helper.md # troubleshooting subagent
└── README.md
```

## Notes

This repo's `Azure_Connect_Script.ps1` uses the deprecated `AzureRM` module
(`Connect-AzureRmAccount`), which is retired. The commands here recommend the
supported `Az` PowerShell module and Azure CLI equivalents. No credentials or
secrets are stored by this plugin — all sign-in is interactive.
