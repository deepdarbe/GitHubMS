---
description: Connect to an Azure account and show the current context and available subscriptions
argument-hint: "[subscription name or id]"
allowed-tools: Bash, Read
---

Help the user connect to Azure and establish a working context.

Steps:

1. Detect whether the Azure PowerShell module (`Az`) or the older `AzureRM`
   module is available, and whether the Azure CLI (`az`) is installed.
   - PowerShell: `Get-Module -ListAvailable Az.Accounts, AzureRM.profile`
   - CLI: `az version`
2. Prefer the modern `Az` module / Azure CLI. The legacy
   `Azure_Connect_Script.ps1` in this repo uses the deprecated
   `Connect-AzureRmAccount` cmdlet — mention that `AzureRM` is retired and
   the equivalent modern flow is `Connect-AzAccount`.
3. Connect:
   - PowerShell: `Connect-AzAccount`
   - CLI: `az login`
4. Show the active context and subscriptions:
   - PowerShell: `Get-AzContext`, then `Get-AzSubscription`
   - CLI: `az account show`, then `az account list -o table`
5. If the user supplied "$ARGUMENTS", select that subscription:
   - PowerShell: `Set-AzContext -Subscription "$ARGUMENTS"`
   - CLI: `az account set --subscription "$ARGUMENTS"`

Report the resulting active subscription and tenant clearly. Do not run
interactive login commands silently — explain what each command does before
running it, since browser-based authentication may be required.
