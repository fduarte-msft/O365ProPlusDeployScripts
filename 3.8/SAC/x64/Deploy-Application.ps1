<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false,
	[Parameter(Mandatory=$false)]
	[ValidateSet('O365ProPlusRetail','VisioProRetail','ProjectProRetail','ProjectProXVolume','VisioProXVolume','ProjectStdXVolume','VisioStdXVolume')]
	[string]$Products = 'O365ProPlusRetail'
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Microsoft'
	[string]$appName = 'Office 365 ProPlus'
	[string]$appVersion = '16.0'
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '23/09/2019'
	[string]$appScriptAuthor = 'Fred Duarte'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.0'
	[string]$deployAppScriptDate = '23/09/2019'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Office applications if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		Show-InstallationWelcome -CloseApps 'excel,groove,onenote,onenotem,infopath,onenote,outlook,mspub,powerpnt,lync,communicator,winword,winproj,visio' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## Create Products list
		[Collections.Generic.List[String]]$newProducts = $Products

		## Get existing configuration
		$c2rConfigRegPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

		## If Office 365 ProPlus products are installed, inventory which ones
		if (Test-Path -Path $c2rConfigRegPath -ErrorAction SilentlyContinue) {
			# Get C2R architecture
			$c2rArch = Get-ItemProperty -Path $c2rConfigRegPath -Name 'Platform'
			if ($c2rArch.Platform -match 'x86') {
				Write-Log -Message "Office 365 ProPlus will be migrated to 64-bit" -Severity 1 -Source $deployAppScriptFriendlyName
				$migrateArch = $true
			}
		
			# Get C2R channel
			$c2rChannel = Get-ItemProperty -Path $c2rConfigRegPath -Name 'CDNBaseURL'
			if ($c2rChannel.CDNBaseUrl -notmatch 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114') {
				Write-Log -Message "Office 365 ProPlus will be migrate to Semi-Annual Channel." -Severity 1 -Source $deployAppScriptFriendlyName
				$migrateChannel = $true
			}
		
			# Get C2R installed products
			$c2rInstalledProducts = Get-ItemProperty -Path $c2rConfigRegPath -Name 'ProductReleaseIds'
		}
		
		# Build a list of products to install
		if ($c2rInstalledProducts) {
			# Multiple C2R products installed
			if ($c2rInstalledProducts.ProductReleaseIds -match ',') {
				$c2rProducts = $c2rInstalledProducts.ProductReleaseIds.Split(',')
				foreach ($c2rProduct in $c2rProducts) {
					if ($c2rProduct -notmatch $Products) {
						$newProducts.Add($c2rProduct)
					}
				}
			# Single C2R product installed
			} else {
				if ($c2rInstalledProducts.ProductReleaseIds -notmatch $Products) {
					$newProducts.Add($c2rInstalledProducts.ProductReleaseIds)
		
				}
			}
		} 
		
		# Convert Visio Standard 2016 to Visio Pro for Office 365
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioStdXVolume') {
			if ($Products -match 'VisioProRetail') {
				if ($newProducts.Contains('VisioStdXVolume')) {
					$newProducts.Remove('VisioStdXVolume')
					Write-Log -Message "Microsoft Visio Standard 2016 (VL) will be migrated to Microsoft Visio Pro for Office 365 (Subscription)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Visio Professional 2016 to Visio Pro for Office 365
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioProXVolume') {
			if ($Products -match 'VisioProRetail') {
				if ($newProducts.Contains('VisioProXVolume')) {
					$newProducts.Remove('VisioProXVolume')
					Write-Log -Message "Microsoft Visio Proffesional 2016 (VL) will be migrated to Microsoft Visio Pro for Office 365 (Subscription)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Project Standard 2016 to Project Pro for Office 365
		if ($c2rInstalledProducts.ProductReleaseIds -match 'ProjectStdXVolume') {
			if ($Products -match 'ProjectProRetail') {
				if ($newProducts.Contains('ProjectStdXVolume')) {
					$newProducts.Remove('ProjectStdXVolume')
					Write-Log -Message "Microsoft Project Standard 2016 (VL) will be migrated to Microsoft Project Pro for Office 365 (Subscription)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Project Professional 2016 to Project Pro for Office 365
		if ($c2rInstalledProducts.ProductReleaseIds -match 'ProjectProXVolume') {
			if ($Products -match 'ProjectProRetail') {
				if ($newProducts.Contains('ProjectProXVolume')) {
					$newProducts.Remove('ProjectProXVolume')
					Write-Log -Message "Microsoft Project Proffesional 2016 (VL) will be migrated to Microsoft Project Pro for Office 365 (Subscription)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Visio Pro for Office 365 to Visio Standard 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioProRetail') {
			if ($Products -match 'VisioStdXVolume') {
				if ($newProducts.Contains('VisioProRetail')) {
					$newProducts.Remove('VisioProRetail')
					Write-Log -Message "Microsoft Visio Pro for Office 365 (Subscription) will be migrated to Microsoft Visio Standard 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Visio Pro for Office 365 to Visio Professional 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioProRetail') {
			if ($Products -match 'VisioProXVolume') {
				if ($newProducts.Contains('VisioProRetail')) {
					$newProducts.Remove('VisioProRetail')
					Write-Log -Message "Microsoft Visio Pro for Office 365 (Subscription) will be migrated to Microsoft Visio Professional 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Project Pro for Office 365 to Project Standard 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'ProjectProRetail') {
			if ($Products -match 'ProjectStdXVolume') {
				if ($newProducts.Contains('ProjectProRetail')) {
					$newProducts.Remove('ProjectProRetail')
					Write-Log -Message "Microsoft Project Pro for Office 365 (Subscription) will be migrated to Microsoft Project Standard 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Project Pro for Office 365 to Project Professional 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'ProjectProRetail') {
			if ($Products -match 'ProjectProXVolume') {
				if ($newProducts.Contains('ProjectProRetail')) {
					$newProducts.Remove('ProjectProRetail')
					Write-Log -Message "Microsoft Project Pro for Office 365 (Subscription) will be migrated to Microsoft Project Professional 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Visio Standard 2016 to Visio Professional 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioStdXVolume') {
			if ($Products -match 'VisioProXVolume') {
				if ($newProducts.Contains('VisioStdXVolume')) {
					$newProducts.Remove('VisioStdXVolume')
					Write-Log -Message "Microsoft Visio Standard 2016 (VL) will be migrated to Microsoft Visio Professional 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}
		
		# Convert Visio Professional 2016 to Visio Standard 2016
		if ($c2rInstalledProducts.ProductReleaseIds -match 'VisioProXVolume') {
			if ($Products -match 'VisioStdXVolume') {
				if ($newProducts.Contains('VisioProXVolume')) {
					$newProducts.Remove('VisioProXVolume')
					Write-Log -Message "Microsoft Visio Professional 2016 (VL) will be migrated to Microsoft Visio Standard 2016 (VL)" -Severity 1 -Source $deployAppScriptFriendlyName
					$office365 = $true
				}
			}
		}

		## Log Office 365 ProPlus inventory data
		$oldConfig = $c2rInstalledProducts.ProductReleaseIds -join ','
		$newConfig = $newProducts -join ','

		## Log old configuration and new configuration
		Write-Log -Message "Old Configuraiton: $oldConfig" -Severity 1 -Source $deployAppScriptFriendlyName
		Write-Log -Message "New Configuration: $newConfig" -Severity 1 -Source $deployAppScriptFriendlyName

		## Get all Microsoft installations
		$officeProducts = Get-InstalledApplication -Name 'Microsoft'

		## Determine which Office products we will remove 
		foreach ($officeProduct in $officeProducts) {
			if ($officeProduct.DisplayName -match 'office professional plus 2007') {
				Write-Log -Message "$($officeProduct.DisplayName) is installed. Will be removed to install Office 365 ProPlus (Semi-Annual Channel)." -Severity 1 -Source $deployAppScriptFriendlyName
				$office2k7 = $true
			}

			if ($officeProduct.DisplayName -match 'office professional plus 2010') {
				Write-Log -Message "$($officeProduct.DisplayName) is installed. Will be removed to install Office 365 ProPlus (Semi-Annual Channel)." -Severity 1 -Source $deployAppScriptFriendlyName
				$office2k10 = $true
			}

			if ($officeProduct.DisplayName -match 'office professional plus 2013') {
				Write-Log -Message "$($officeProduct.DisplayName) is installed. Will be removed to install Office 365 ProPlus (Semi-Annual Channel)." -Severity 1 -Source $deployAppScriptFriendlyName
				$office2k13 = $true
			}

			if ($officeProduct.DisplayName -match 'office professional plus 2016') {
				Write-Log -Message "$($officeProduct.DisplayName) is installed. Will be removed to install Office 365 ProPlus (Semi-Annual Channel)." -Severity 1 -Source $deployAppScriptFriendlyName
				$office2k16 = $true
			}

			if ($officeProduct.DisplayName -match 'office professional plus 2019') {
				Write-Log -Message "$($officeProduct.DisplayName) is installed. Will be removed to install Office 365 ProPlus (Semi-Annual Channel)." -Severity 1 -Source $deployAppScriptFriendlyName
				$office2k19 = $true
			}
		}

		## Scrub previous versions of Office
		
		## Office 2007
		if ($office2k7) {
			try {
				Write-Log -Message "Uninstalling Office Professional Plus 2007" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrub07.vbs ALL /FR /QUIET /LOG $configToolkitLogDir\Office_2007_Uninstall" -WindowStyle 'Hidden' -PassThru
                Write-Log -Message "Uninstalled Office Professional PLus 2007. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally { 
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Office 2010
		if ($office2k10) {
			try {
				Write-Log -Message "Uninstalling Office Professional Plus 2010" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrub10.vbs ALL /FR /QUIET /LOG $configToolkitLogDir\Office_2010_Uninstall" -WindowStyle 'Hidden' -PassThru
				Write-Log -Message "Uninstalled Office Professional PLus 2010. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally {
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Office 2013
		if ($office2k13) {
			try {
				Write-Log -Message "Uninstalling Office Professional Plus 2013" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrub_O15msi.vbs ALL /FR /QUIET /LOG $configToolkitLogDir\Office_2013_Uninstall" -WindowStyle 'Hidden' -PassThru
				Write-Log -Message "Uninstalled Office Professional Plus 2013. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally {
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Office 2016
		if ($office2k16) {
			try {
				Write-Log -Message "Uninstalling Office Professional Plus 2016" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrub_O16msi.vbs ALL /FR /QUIET /LOG $configToolkitLogDir\Office_2016_Uninstall" -WindowStyle 'Hidden' -PassThru
				Write-Log -Message "Uninstalled Office Professional Plus 2016. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally {
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Office 2019
		if ($office2k19) {
			try {
				Write-Log -Message "Uninstalling Office Professional Plus 2019" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrubC2R.vbs /QUIET /RETERRORSUCCESS /LOG $configToolkitLogDir\Office_proplus_2019_Uninstall" -WindowStyle 'Hidden' -PassThru
				Write-Log -Message "Uninstalled Office Professional Plus 2019. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally {
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Office 365 ProPlus
		if ($office365 -or $migrateChannel -or $migrateArch) {
			try {
				Write-Log -Message "Uninstalling Office 365 ProPlus" -Severity 1 -Source $deployAppScriptFriendlyName
				$scrub = Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "$dirSupportFiles\OffScrubC2R.vbs /QUIET /RETERRORSUCCESS /LOG $configToolkitLogDir\Office_365_proplus_Uninstall" -WindowStyle 'Hidden' -PassThru
				Write-Log -Message "Uninstalled Office 365 ProPlus. Exit code: $($scrub.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
			} catch {
				Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			} finally {
                $mainExitCode = $scrub.ExitCode
            }
		}

		# Build configuration file
		$config = ""

		If ($newProducts.Contains('O365ProPlusRetail')) {
    		$config += @"
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('VisioStdXVolume')) {
    		$config += @"
    <Product ID="VisioStdXVolume" PIDKEY="NY48V-PPYYH-3F4PX-XJRKJ-W4423">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('VisioProXVolume')) {
			$config += @"
    <Product ID="VisioProXVolume" PIDKEY="69WXN-MBYV6-22PQG-3WGHK-RM6XC">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('ProjectStdXVolume')) {
			$config += @"
    <Product ID="ProjectStdXVolume" PIDKEY="D8NRQ-JTYM3-7J2DX-646CT-6836M">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('ProjectProXVolume')) {
			$config += @"
    <Product ID="ProjectProXVolume" PIDKEY="WGT24-HCNMF-FQ7XH-6M8K7-DRTW9">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('VisioProRetail')) {
			$config += @"
    <Product ID="VisioProRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

		if ($newProducts.Contains('ProjectProRetail')) {
			$config += @"
    <Product ID="ProjectProRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>

"@
}

$configXML = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Broad" >

$($config)

  </Add>

  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="$($configToolkitLogDir)" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>

</Configuration>
"@

		# Log location of configuration.xml file
		Write-Log -Message "The final Office 365 ProPlus 64-bit (Semi-Annual Channel) confirugarion will be written at `"$configXmlFilePath`"" -Severity 1 -Source $deployAppScriptFriendlyName

		# if configuration XML file exists, let's remove it
        $configXmlFilePath = Join-Path $envTemp 'configuration.xml'
		if (Test-Path -Path "$configXmlFilePath") {
			Write-Log -Message "Removing existing configuration.xml file from: `"$configXmlFilePath`"" -Severity 1 -Source $deployAppScriptFriendlyName
			Remove-Item -Path "$configXmlFilePath" -Force -ErrorAction SilentlyContinue
		}

		# Write Configuration XML file
		$configXML | Out-File -FilePath "$configXmlFilePath" -Encoding default -Force

		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## Install Office 365 ProPlus

		try {
			Write-Log -Message "Starting: Office 365 ProPlus configuration" -Severity 1 -Source $deployAppScriptFriendlyName
			$install = Execute-Process -Path "$dirFiles\Setup.exe" -Parameters "/configure `"$configXmlFilePath`"" -PassThru
			Write-Log -Message "Exit Code: $($install.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
		} catch {
			Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			Write-Log -Message "Failure: $($install.ExitCode)" -Severity 3 -Source $deployAppScriptFriendlyName
		} finally {
            $mainExitCode = $install.ExitCode
        }
		
		if ($mainExitCode -ne 0) {
			Write-Log -Message "Failure: Please check logs at $configToolkitLogDir" -Severity 3 -Source $deployAppScriptFriendlyName
			Write-Log -Message "Exit code: $($mainExitCode)" -Severity 3 -Source $deployAppScriptFriendlyName
		}


		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>
		
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Welcome Message, close Office applications if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		Show-InstallationWelcome -CloseApps 'excel,groove,onenote,onenotem,infopath,onenote,outlook,mspub,powerpnt,lync,communicator,winword,winproj,visio' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		Show-InstallationProgress

        $config = "" 

        if ($Products.Contains('O365ProPlusRetail')) {
            $config += @"
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('VisioStdXVolume')) {
            $config += @"
    <Product ID="VisioStdXVolume">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('VisioProXVolume')) {
            $config += @"
    <Product ID="VisioProXVolume">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('ProjectStdXVolume')) {
            $config += @"
    <Product ID="ProjectStdXVolume">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('ProjectProXVolume')) {
            $config += @"
    <Product ID="ProjectProXVolume">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('VisioProRetail')) {
            $config += @"
    <Product ID="VisioProRetail">
      <Language ID="en-us" />
    </Product>

"@
}

        if ($Products.Contains('ProjectProRetail')) {
            $config += @"
    <Product ID="ProjectProRetail">
      <Language ID="en-us" />
    </Product>

"@
}

$configXML = @"
<Configuration>
  <Remove ALL="FALSE">

$($config)

  </Remove>

  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="$($configToolkitLogDir)" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>

</Configuration>
"@

        # if configuration XML file exists, let's remove it
        $configXmlFilePath = Join-Path $envTemp 'configuration.xml'
        if (Test-Path -Path $configXmlFilePath) {
            Write-Log -Message "Removing existing configuration.xml file from: `"$configXmlFilePath`"" -Severity 1 -Source $deployAppScriptFriendlyName
            Remove-Item -Path $configXmlFilePath -Force -ErrorAction SilentlyContinue
        }

        # Write Configuration XML file
        $configXML | Out-File -FilePath $configXmlFilePath -Encoding default -Force
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		## Uninstall Office 365 ProPlus

        try {
			Write-Log -Message "Starting: Office 365 ProPlus configuration" -Severity 1 -Source $deployAppScriptFriendlyName
			$uninstall = Execute-Process -Path "$dirFiles\Setup.exe" -Parameters "/configure `"$configXmlFilePath`"" -PassThru
			Write-Log -Message "Exit Code: $($uninstall.ExitCode)" -Severity 1 -Source $deployAppScriptFriendlyName
		} catch {
			Write-Log -Message "Failure: $($Error[0].Exception.Message)" -Severity 3 -Source $deployAppScriptFriendlyName
			Write-Log -Message "Failure: $($uninstall.ExitCode)" -Severity 3 -Source $deployAppScriptFriendlyName
		} finally {
            $mainExitCode = $uninstall.ExitCode
        }
		
		if ($mainExitCode -ne 0) {
			Write-Log -Message "Failure: Please check logs at $configToolkitLogDir" -Severity 3 -Source $deployAppScriptFriendlyName
			Write-Log -Message "Exit code: $($mainExitCode)" -Severity 3 -Source $deployAppScriptFriendlyName
		}


		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'


	}

	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
