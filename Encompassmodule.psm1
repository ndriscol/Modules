

function Add-ToZipFiles
{
<#
.Synopsis
   Compress files in a directory to a zip file.
.DESCRIPTION
   Long description
.EXAMPLE
   Add-ToZipFiles -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
.EXAMPLE
   Add-ToZipFiles -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
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
   Add-ToZipFiles -$DestinationZipFile C:\ZippedFiles\TempZip.zip -$SourceDirectory C:\temp\ 
   
   Add-ToZipFiles -DestinationZipFile C:\Scripts\ProdTestRemove.zip -SourceDirectory C:\Tempdir\ -RemoveSourceFiles True
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