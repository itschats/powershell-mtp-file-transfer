#this is an enhanced version of https://github.com/nosalan/powershell-mtp-file-transfer/blob/master/phone_backup.ps1
#Used to copy files from PC to Phone
#modified: 1-Jan-2025 - itschats
#####################################################################################

$ErrorActionPreference = [string]"Stop"
###Update this to your PC folder
$SourceDirForWhatsApp = [string]"C:\Backups\Folder1\Android\media\com.whatsapp\WhatsApp"
$Summary = [Hashtable]@{NewFilesCount=0; ExistingFilesCount=0}

function Get-SubFolder($parentDir, $subPath)
{
  $result = $parentDir
  foreach($pathSegment in ($subPath -split "\\"))
  {
    $result = $result.GetFolder.Items() | Where-Object {$_.Name -eq $pathSegment} | select -First 1
    if($result -eq $null)
    {
	  $parentDir.GetFolder.NewFolder($pathSegment)
	  $result = $parentDir.GetFolder.Items() | Where-Object {$_.Name -eq $pathSegment} | select -First 1
    }
  }
  return $result;
}


function Get-PhoneMainDir($phoneName)
{
  $o = New-Object -com Shell.Application
  $rootComputerDirectory = $o.NameSpace(0x11)
  $phoneDirectory = $rootComputerDirectory.Items() | Where-Object {$_.Name -eq $phoneName} | select -First 1
    
  if($phoneDirectory -eq $null)
  {
    throw "Not found '$phoneName' folder in This computer. Connect your phone."
  }
  
  return $phoneDirectory;
}


function Copy-FromBackupSource-ToPhone($sourceDirPath, $destMtpDir)
{
  $sourceDirShell = (new-object -com Shell.Application).NameSpace($sourceDirPath)
  $destDirShell = (new-object -com Shell.Application).NameSpace($fulldestDirPath)
  $destExistingItems = $destMtpDir.GetFolder.Items()

  $copiedCount, $existingCount = 0
  
  foreach($item in (Get-ChildItem $sourceDirPath | Sort-Object -descending))
  {
   $itemName = ($item.Name)

   if(Test-Path $item.FullName -PathType Container)
   {
      Write-Host $item.Name " is folder, stepping into"
      Copy-FromBackupSource-ToPhone  $item.FullName (Get-SubFolder $destMtpDir $item.Name)
   }
   elseif($destMtpDir.GetFolder.ParseName($item.Name))
   {
      Write-Host "Element '$itemName' already exists"
      $existingCount++;
   }
   else
   {
	 $copiedCount++;
	 Write-Host ("Copying #{0}: {1}\{2}" -f $copiedCount, $sourceDirPath, $item.Name)
	 $destMtpDir.GetFolder.CopyHere($item.FullName, 0004)
	 Do {
		 Sleep -Milliseconds 100
		 Write-Host "Waiting to copy....."
	 }While (!($destMtpDir.GetFolder.ParseName($item.Name)))
   }
   $script:Summary.NewFilesCount += $copiedCount
  }
  $script:Summary.NewFilesCount += $copiedCount
  $script:Summary.ExistingFilesCount += $existingCount 
  Write-Host "Copied '$copiedCount' elements from '$sourceDirPath'"
}


#####################################################################################
##Entry point
#####################################################################################

$phoneName = "Phone Name" #Phone name as it appears in This PC
$phoneRootDir = Get-PhoneMainDir $phoneName

try {
	#Can also copy to second account by copying to Android\media\com.whatsapp\WhatsApp\accounts\1001
	Copy-FromBackupSource-ToPhone $SourceDirForWhatsApp (Get-SubFolder $phoneRootDir "Internal shared storage\Android\media\com.whatsapp\WhatsApp\")
	write-host ($Summary | out-string)
} catch {
	$_
} finally {
	write-host ($Summary | out-string)
}