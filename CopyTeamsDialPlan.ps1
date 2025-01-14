<# 
Copy Teams Dial Plan Normalisation Rules to another Dial Plan
    Version: v1.0
    Date: 14/01/2025
    Author: Rob Watts - Cloud Solution Architect - Microsoft
    Description: This script will copy Teams Dial Plan Normalisation Rules to another Dial Plan.

DISCLAIMER
   THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
   MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES
   OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR
   PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR
   ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS
   INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR
   INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
   BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR
   INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
#>

# Pop out disclaimer
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[Windows.Forms.MessageBox]::Show("
THIS CODE IS SAMPLE CODE. 

THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.

MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. 

THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU.

IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.", "***DISCLAIMER***", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)

#Check to see if Microsoft Teams PowerShell module is installed
Write-Host "Checking if MicrosoftTeams module exists..."
If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "404: MicrosoftTeams module not found. Please install Teams PowerShell module." -ForegroundColor DarkRed
  }
  Else {
    Write-Host "MicrosoftTeams module exists... Continue we shall... " -ForegroundColor Green
  }

# Import Teams Module
Import-Module MicrosoftTeams

# Connects to Microsoft Teams
Write-Host "Connecting to the Mother Ship (Microsoft Teams)..."
Connect-MicrosoftTeams

# Shows Gridview of Dial Plans, user selects Dial Plan to copy from
$FromDialPlan = Get-CsTenantDialPlan | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Dial Plan to copy from."

# Shows Normalisation Rules of selected Dial Plan
(Get-CsTenantDialPlan -Identity $FromDialPlan.Identity).NormalizationRules

# Pause for user to check Normalisation Rules
Write-Host "Please check dial plan normalisation rules above. If happy to proceed, press any key to continue"
Pause

# Shows Gridview of Dial Plans, user selects Dial Plan to copy to
$ToDialPlan = Get-CsTenantDialPlan | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Dial Plan to copy to."
$NR=(Get-CsTenantDialPlan -Identity $FromDialPlan.Identity).NormalizationRules

# Copy Normalisation Rules to selected Dial Plan
Set-CsTenantDialPlan -Identity $ToDialPlan.Identity -NormalizationRules @{add=$nr}
Write-Host "Normalisation Rules copied from $($FromDialPlan.Identity) to $($ToDialPlan.Identity)"

# Shows updated normalisation Rules of selected Dial Plan
(Get-CsTenantDialPlan -Identity $ToDialPlan.Identity).NormalizationRules

Write-Host "Script completed"