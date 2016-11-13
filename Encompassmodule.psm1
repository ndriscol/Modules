

function Add-ToZipFiles
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
   The functionality that best describes this cmdlet
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
        $SourceDirectory

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

    try{

            

        }catch [System.Exception]{


        }


    }

}