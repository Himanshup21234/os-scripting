#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

Function Check-FilesPath()
{
param
    (
        [Parameter(Mandatory=$true)] [string] $LocalFolder
    )
If (!(Test-Path -Path $LocalFolder)) {
                New-Item -ItemType Directory -Path $LocalFolder | Out-Null
        }
}

Function Download-AllFilesFromLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.Folder] $SourceFolder,
        [Parameter(Mandatory=$true)] [string] $TargetFolder
    )
    Try {
         
        #Create Local Folder, if it doesn't exist
        $FolderName = ($SourceFolder.ServerRelativeURL) -replace "/","\"
        #$FolderName=$FolderName.split("\")[-1]
        $LocalFolder = $TargetFolder + $FolderName
        Check-FilesPath -LocalFolder $LocalFolder
         
        #Get all Files from the folder
        $FilesColl = $SourceFolder.Files
        $Ctx.Load($FilesColl)
        $Ctx.ExecuteQuery()
 
        #Iterate through each file and download
        Foreach($File in $FilesColl)
        {
            $TargetFile = $LocalFolder+"\"+$File.Name
            #Download the file
            $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$File.ServerRelativeURL)
            $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
            $FileInfo.Stream.CopyTo($WriteStream)
            $WriteStream.Close()
            write-host -f Green "Downloaded File:"$TargetFile
        }
         
        #Process Sub Folders
        $SubFolders = $SourceFolder.Folders
        $Ctx.Load($SubFolders)
        $Ctx.ExecuteQuery()
        Foreach($Folder in $SubFolders)
        {
            If($Folder.Name -ne "Forms")
            {
                #Call the function recursively
                Download-AllFilesFromLibrary -SiteURL $SiteURL -SourceFolder $Folder -TargetFolder $TargetFolder
            }
        }
  }
    Catch {
        write-host -f Red "Error Downloading Files from Library!" $_.Exception.Message
    }
}

#Set parameter values
$SiteURL="https://site.sharepoint.com/sites/123ASDFAF/ABC/"
$LibraryName="HP Documents"
$TargetFolder="E:\EXTRACTED_DATA\"

#Setup Credentials to connect
$username = "uname@domain.company.com"
$password = "encrypted_password"
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $(convertto-securestring $password -asplaintext -force)
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

#Setup the context
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Ctx.Credentials = $Credentials
      
#Get the Library
$List = $Ctx.Web.Lists.GetByTitle($LibraryName)
$Ctx.Load($List)
$Ctx.Load($List.RootFolder)
$Ctx.ExecuteQuery()

#Calling of Above functions
Download-AllFilesFromLibrary -SiteURL $SiteURL -SourceFolder $List.RootFolder -TargetFolder $TargetFolder
