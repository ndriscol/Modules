

function Compress-ItemsToZip
{
<#
.Synopsis
   Compress files in a directory to a zip. Superceeded in powershell v5 
.DESCRIPTION
   Long description
.EXAMPLE
   Compress-ItemsToZip -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
.EXAMPLE
   Compress-ItemsToZip -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   Compress-ItemsToZip -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
   
   Compress-ItemsToZip -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
#>

    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        $DestinationZipFile,
        [Parameter()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        $SourceDirectory,
        [Parameter()]
        [Validateset('True','False')]
        $RemoveSourceFiles

        )

if( ! ( Test-Path -Path $DestinationZipFile ) ){
   
    try{

            Add-Type -Assembly System.IO.Compression.FileSystem -ErrorAction Stop
            $CompressionLevel = [System.IO.Compression.CompressionLevel]::Optimal 
            
        }Catch [System.exception]{

                $_ | fl * -Force

        }

    Try{

            [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectory,$DestinationZipFile, $CompressionLevel, $false)

        }Catch [System.exception]{

                $_ | fl * -Force

        }

        if( $RemoveSourceFiles -match $true ){

            try{

                    Remove-Item -Path $SourceDirectory -Force -Recurse -ErrorAction Stop 

                }catch [System.Exception]{

                     $_ | fl * -Force

                }
        }

    }else{


        Write-Output "Error: $DestinationZipFile already exists choose another filename for your ZIP file."

    }

}


Extract-CompressItems{
<#
.Synopsis
   Compress files in a directory to a zip. Superceeded in powershell v5 
.DESCRIPTION
   Long description
.EXAMPLE
   Compress-ItemsToZip -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
.EXAMPLE
   Compress-ItemsToZip -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   Compress-ItemsToZip -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
   
   Compress-ItemsToZip -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
#>

    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$SourceZip,
        [Parameter()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        $DestinationDirectory

        )

    try{

            Add-Type -Assembly System.IO.Compression.FileSystem -ErrorAction Stop
             
            
        }Catch [System.exception]{

                $_ | fl * -Force

        }


Foreach( $Zip in $SourceZip ){

     if( (Test-Path -Path $Zip) -and ( Test-Path -Path $DestinationDirectory ) ){

          try{

                   [System.IO.Compression.ZipFile]::ExtractToDirectory($Zip, $DestinationDirectory, $False)
         
              }Catch [System.exception]{

                    $_ | fl * -Force

            }

        }else{

        
            Write-Output "Error: $Zip or $DestinationDirectory. Doesnt Exist"
        


        }

    }

}


function Set-ScomMaintenanceMode
{
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
 
    Param(       
            [Parameter()]
            $Inputlist,
            [Parameter()]
            [int]$maintenanceModeMinutes,
            [Parameter()]
            [String]$maintenanceModeComment,
            [ValidateSet("Start", "Stop")]
            $MaintenanceStopOrStart,
            [Parameter()]
            [String]$ManagementServer = 'SCOM.domain.local'
    )
    

if ( Test-Path -Path $Inputlist )  {     

New-SCOMManagementGroupConnection –ComputerName $ManagementServer

    Write-Output ("Stopping or starting Maintenance mode? {0}" -f $maintenanceStoporstart )

    if( $maintenanceStoporstart -match 'Start' ){

       ForEach ( $Computer in ( Import-csv -Path $inputlist -ErrorAction Stop ).Name ){
       
        Try{
             
               
                $NewEndTime = (Get-Date).addMinutes( $maintenanceModeMinutes )
                $ScomObj = Get-SCOMClassInstance -Name $Computer 
                foreach( $obj in $ScomObj ){

                   if( ! ( $obj.InMaintenanceMode -eq $true ) ){

                        Start-SCOMMaintenanceMode -Instance $obj -EndTime $NewEndTime -Comment $maintenanceModeComment -ErrorAction Stop

                    }
                }

                Write-Output ("Placeing {0} into maintenance mode until {1}!" -f $Computer,$NewEndTime )
        
            }Catch [System.Exception]{


                $_ | fl * -Force

            }
        }

}elseif( $maintenanceStoporstart -eq 'Stop' ){

New-SCOMManagementGroupConnection –ComputerName $ManagementServer -Verbose

     ForEach ( $Computer in ( Import-csv -Path $inputlist -ErrorAction Stop ).Name ){
       
       Try{

               $ObjectInstance = Get-SCOMClassInstance -Name $Computer
               Get-SCOMMaintenanceMode -Instance $ObjectInstance | Set-SCOMMaintenanceMode -EndTime (Get-Date) -ErrorAction Stop -Verbose
               Write-Output ("stopping maintenance on {0}!" -f $ObjectInstance[0])

           }catch [System.exception]{

               $_ | fl * -Force

          }
    }
}

}else{



    Write-Output ( "CSV Doest exist on path {0}!" -f  $Inputlist )


}
}

[System.Net.WebRequest]::DefaultWebProxy.Credentials =  [System.Net.CredentialCache]::DefaultCredentials

Function Get-VaultItem{
[cmdletbinding()]
Param(


)

    $null = [Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime] 
    $windowsCredentialStore = [Windows.Security.Credentials.PasswordVault]::new()
    $windowsCredentialStore.RetrieveAll() | Select-Object -Property UserName,Resource

}

Function Add-VaultItem{
[cmdletbinding()]
Param(
	[Parameter(Mandatory=$true)]
	[String]$ResourceName,

	[Parameter(Mandatory=$false)]
	[PSCredential]$PSCredential = (Get-Credential)
)

$Null = [Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]
$windowsCredentialStore = [Windows.Security.Credentials.PasswordVault]::new()

    try{
	    $windowsCredentialStore.FindAllByUserName($ResourceName) | Out-Null
	    throw $ResourceName + "Already exists! Choose another resource name to store the credential under!"
    }catch [system.exception]{	
	    $StoreCreds = [Windows.Security.Credentials.PasswordCredential]::new($ResourceName, $PSCredential.getnetworkcredential().Username,$PSCredential.getnetworkcredential().Password) 
	    $windowsCredentialStore.Add($StoreCreds)
    }

}

Function Remove-VaultItem{
[cmdletbinding()]
Param(
	[Parameter(Mandatory=$true)]
	[String]$ResourceName,

	[Parameter(Mandatory=$true)]
	[String]$UserName
)

$null =[Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]
$windowsCredentialStore = [Windows.Security.Credentials.PasswordVault]::new()

    try{
        $Resource = $windowsCredentialStore.Retrieve($ResourceName,$UserName)
	    $windowsCredentialStore.Remove($Resource)
    }catch [system.exception]{	
        throw $ResourceName + " Record doesnt exist!"
    }

}

function Retrieve-VaultItem{
Param(
	[Parameter(Mandatory=$true)]
	[String]$ResourceName,

	[Parameter(Mandatory=$true)]
	[String]$UserName
)

$null = [Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime] 
$windowsCredentialStore = [Windows.Security.Credentials.PasswordVault]::new()


    try{
	Set-Clipboard -value ($windowsCredentialStore.Retrieve($ResourceName,$UserName)).Password
	Write-Host ('Credential for {0} has been added to the clipboard' -f $UserName)
    }catch [system.exception]{	
        throw 'Record didnt exist to retrieve'
    }

}

Export-ModuleMember -Function Add-VaultItem, Get-VaultItem, Remove-VaultItem, Retrieve-VaultItem



 

