<#
.SYNOPSIS
	This script will download missing images used by OfflineList.

.DESCRIPTION
	This script has been written to download all missing images used by OfflineList. It will read/update the existing XML dat files in the OfflineList installatoi folder. OfflineList should be installed before this script can be used.

.NOTES
	Author    : Sander Siemonsma
	Date      : 11-10-2017
	Version   : 1.0
	PSVersion : 5.1

.PARAMETER WhatIf
	The WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes. You don't need to specify a value with this switch.

.EXAMPLE
	.\Start-OfflineListDownload.ps1
	Start script normally.

.EXAMPLE
	.\Start-OfflineListDownload.ps1 -Verbose
	Start script with aditional information.

.LINK
	http://offlinelist.free.fr/
#>


[CmdletBinding()]
param
(
	[parameter(Mandatory=$false)][switch]$WhatIf
)

#Variables
$OfflineListProgramFolder=(Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OfflineList" -ErrorAction SilentlyContinue).DisplayIcon | Split-Path -ErrorAction SilentlyContinue
if(!($OfflineListProgramFolder))
{
	Write-Error -Message "OfflineList is not installed. Cannot determine installation path."
	exit
}
$OfflineListDataFolder=Join-Path -Path $OfflineListProgramFolder -ChildPath "datas"
$OfflineListDataItems=Get-Item -Path $OfflineListDataFolder\* -Include *.zip
$OfflineListImageFolder=Join-Path -Path $OfflineListProgramFolder -ChildPath "imgs"

#New-Objects
$wc=New-Object -TypeName System.Net.WebClient
$Shell=New-Object -ComObject Shell.Application

function Get-DataFolderSuffix
{
	<#
	.SYNOPSIS
		Get Foldername based on image number.
	
	.PARAMETER ImageNumber
		Image number where foldername needs to be checked for.

	.INPUTS
		Int32

	.OUTPUTS
		System.String
	#>

	param
	(
		[parameter(Mandatory=$true)][int]$ImageNumber
	)
	$lower=[math]::floor(($ImageNumber-1)/500)*500+1
	$upper=[math]::floor(($ImageNumber-1)/500)*500+500
	return "$lower-$upper"
}

function Get-XMLData
{
	<#
	.SYNOPSIS
		Load Zip file with XML Data.
	
	.PARAMETER ZipFile
		Zip file containing the XML file.

	.INPUTS
		String

	.OUTPUTS
		System.Xml.XmlDocument
	#>

	param
	(
		[parameter(Mandatory=$true)][string]$ZipFile
	)

	$Zip=$Shell.NameSpace($ZipFile)
	if($Zip.Items().Count -ne 1)
	{
		Write-Error -Message "Wrong zip file specified. Zip file must only contain one xml file."
		exit
	}
	$XMLFile=$Zip.Items() | Select-Object -ExpandProperty Path | Split-Path -Leaf
	$Shell.NameSpace($PWD.Path).CopyHere($Zip.Items())
	$XML=[xml](Get-Content -Path $XMLFile)
	Remove-Item -Path $XMLFile
	return $XML
}

workflow Start-DownloadItems
{
	<#
	.SYNOPSIS
		Start parallel download of images
	
	.PARAMETER DownloadItems
		Hash table with download images. Has to contain member names ImageFolder, ImageItem and ImageURL.
	
	.INPUTS
		System.Management.Automation.PSCustomObject
	#>

	param
	(
		[parameter(Mandatory=$true)]$DownloadItems
	)

	foreach -parallel($DownloadItem in $DownloadItems)
	{
		InlineScript
		{
			if(!(Test-Path -Path $using:DownloadItem.ImageFolder))
			{
				Write-Verbose -Message "Creating folder $($using:DownloadItem.ImageFolder)"
				New-Item -ItemType Directory -Path $using:DownloadItem.ImageFolder
			}
			$wc=New-Object -TypeName System.Net.WebClient
			Write-Verbose -Message "Downloading: $($using:DownloadItem.ImageURL)"
			Write-Verbose -Message "Target location: $($using:DownloadItem.ImageItem)"
			$wc.DownloadFile($using:DownloadItem.ImageURL,$using:DownloadItem.ImageItem)
		}
	}
}

foreach($Item in $OfflineListDataItems)
{
	Write-Output -InputObject "Retreiving XML Data for: $($Item.Name)"
	$OfflineListDataItem=Join-Path -Path $OfflineListDataFolder -ChildPath $Item.Name
	$XML=Get-XMLData -ZipFile $OfflineListDataItem
	Write-Output -InputObject "Data loaded: $($XML.dat.configuration.datName)"

	#Check for new version of dat file.
	Write-Output -InputObject "Checking XML version"
	$datVersionURL=$XML.dat.configuration.newDat.datVersionURL
	$datVersionItem=Join-Path -Path $PWD -ChildPath $XMLfile".version"
	$wc.DownloadFile($datVersionURL, $datVersionItem)
	$datVersion=Get-Content -Path $datVersionItem
	Remove-Item -Path $datVersionItem
	if(($XML.dat.configuration.datVersion -lt $datVersion) -and !($WhatIf.IsPresent))
	{
		Write-Output -InputObject "Update XML file needed. Download version $datVersion"
		$datURL=$XML.dat.configuration.newDat.datURL."#text"
		$wc.DownloadFile($datURL,$OfflineListDataItem)
		$XML=Get-XMLData -ZipFile $OfflineListDataItem
	}
	elseif(($XML.dat.configuration.datVersion -lt $datVersion) -and ($WhatIf.IsPresent))
	{
		Write-Output -InputObject "What if: Update XML file needed. Download version $datVersion"
	}

	#Variables from XML file
	$XMLImageFolder=Join-Path -Path $OfflineListImageFolder -ChildPath $XML.dat.configuration.datName
	$XMLImageURL=$XML.dat.configuration.newDat.imURL

	#Populating missing images
	$DownloadItems=@()
	Write-Output -InputObject "Populating missing images for $($XML.dat.configuration.datName)"
	$progresscounter=0
	foreach($game in $XML.dat.games.game)
	{
		$progresscounter++
		Write-Progress -Activity "Populating missing images" -Status $game.title -PercentComplete (($progresscounter/$XML.dat.games.game.count)*100)
		$ImageNumber=$game.imageNumber
		$ImageFolder=Join-Path -Path $XMLImageFolder -ChildPath (Get-DataFolderSuffix($ImageNumber))
		$ImageItem=Join-Path -Path $ImageFolder -ChildPath ($ImageNumber + "a.png")
		$ImageURL=$XMLImageURL + (Get-DataFolderSuffix -ImageNumber ($ImageNumber)) + "/" + $ImageNumber + "a.png"
		if(!(Test-Path -Path $ImageItem))
		{
			Write-Verbose -Message "$($XML.dat.configuration.datName): Missing $($ImageNumber)a.png. Adding to Downloadlist."
			$DownloadItems+=New-Object -TypeName PSObject -Property ([ordered]@{
				ImageItem=$ImageItem
				ImageFolder=$ImageFolder
				ImageURL=$ImageURL
			})
		}
		$ImageItem=Join-Path -Path $ImageFolder -ChildPath ($ImageNumber + "b.png")
		$ImageURL=$XMLImageURL + (Get-DataFolderSuffix -ImageNumber ($ImageNumber)) + "/" + $ImageNumber + "b.png"
		if(!(Test-Path -Path $ImageItem))
		{
			Write-Verbose -Message "$($XML.dat.configuration.datName): Missing $($ImageNumber)b.png. Adding to Downloadlist."
			$DownloadItems+=New-Object -TypeName PSObject -Property ([ordered]@{
				ImageItem=$ImageItem
				ImageFolder=$ImageFolder
				ImageURL=$ImageURL
			})
		}
	}
	Write-Progress -Activity "Populating missing images" -Completed
	if(($DownloadItems) -and !($WhatIf.IsPresent))
	{
		Write-Output -InputObject "Starting downloading missing images for $($XML.dat.configuration.datName)"
		Start-DownloadItems -DownloadItems $DownloadItems
		Write-Output -InputObject "Completed downloading missing images for $($XML.dat.configuration.datName)"
	}
	elseif(($DownloadItems) -and ($WhatIf.IsPresent))
	{
		Write-Output -InputObject "What if: Starting downloading missing images for $($XML.dat.configuration.datName)"
	}
	else
	{
		Write-Output -InputObject "No missing images found for $($XML.dat.configuration.datName)"
	}
}