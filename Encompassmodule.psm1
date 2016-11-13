

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


function Start-MaintenanceMode
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
    Param
    (
        
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String[]]$ComputerName,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [int]$maintenanceModeMinutes,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$maintenanceModeComment
    )


   ForEach ( $Computer in $ComputerName ){


        $ObjectInstance = Get-SCOMClassInstance -Name $Computer
        $MaintenanceEntry = Get-SCOMMaintenanceMode -Instance $ObjectInstance
        $NewEndTime = (Get-Date).addMinutes($maintenanceModeMinutes)

    Try{

            Set-SCOMMaintenanceMode -MaintenanceModeEntry $MMEntry -EndTime $NewEndTime -Comment $maintenanceModeComment -ErrorAction
        
        }Catch [System.Exception]{


            $_ | fl * -Force

        }
    }

}