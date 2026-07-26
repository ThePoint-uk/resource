<#
.SYNOPSIS
  Creates the "Promoted Results" list used by the Unified Search web part.

.DESCRIPTION
  Provisions a standard SharePoint list with the columns the web part expects,
  plus two example entries. Content owners then manage promoted results by
  editing list items — no admin centre access or redeployment needed.

  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser

.EXAMPLE
  ./Create-PromotedResultsList.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/intranet"
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [string]$ListTitle = "Promoted Results"
)

Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Creating list '$ListTitle' on $SiteUrl…"
New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch:$false | Out-Null

# Title (built-in) = card headline. Additional columns:
Add-PnPField -List $ListTitle -DisplayName "Keywords"    -InternalName "Keywords"    -Type Note     -AddToDefaultView | Out-Null
Add-PnPField -List $ListTitle -DisplayName "Description" -InternalName "Description" -Type Note     -AddToDefaultView | Out-Null
Add-PnPField -List $ListTitle -DisplayName "LinkUrl"     -InternalName "LinkUrl"     -Type URL      -AddToDefaultView | Out-Null
Add-PnPField -List $ListTitle -DisplayName "ButtonText"  -InternalName "ButtonText"  -Type Text                        | Out-Null
Add-PnPField -List $ListTitle -DisplayName "Enabled"     -InternalName "Enabled"     -Type Boolean  -AddToDefaultView | Out-Null
Add-PnPField -List $ListTitle -DisplayName "StartDate"   -InternalName "StartDate"   -Type DateTime                    | Out-Null
Add-PnPField -List $ListTitle -DisplayName "EndDate"     -InternalName "EndDate"     -Type DateTime                    | Out-Null

# Default Enabled to Yes for new items.
Set-PnPField -List $ListTitle -Identity "Enabled" -Values @{ DefaultValue = "1" } | Out-Null

Write-Host "Adding example promoted results…"
Add-PnPListItem -List $ListTitle -Values @{
  Title       = "Book annual leave in MyHR"
  Keywords    = "annual leave; holiday; leave; time off; book leave"
  Description = "Colleagues book annual leave through MyHR > My Time > Request leave. Requests need line-manager approval and at least two weeks' notice."
  LinkUrl     = "$SiteUrl/SitePages/MyHR.aspx, Open MyHR"
  ButtonText  = "Open MyHR"
  Enabled     = $true
} | Out-Null

Add-PnPListItem -List $ListTitle -Values @{
  Title       = "Need IT help?"
  Keywords    = "it support; password; ticket; helpdesk; service desk"
  Description = "Raise a ticket in the IT portal or call the service desk. Password resets are self-service."
  LinkUrl     = "$SiteUrl/SitePages/IT-Support.aspx, IT Self-Service Portal"
  ButtonText  = "IT Self-Service Portal"
  Enabled     = $true
} | Out-Null

Write-Host "Done. Point the web part's 'Promoted results list' setting at '$ListTitle'."
