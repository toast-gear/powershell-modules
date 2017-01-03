<#
.Synopsis
.EXAMPLE
   Compress-Zip -SourceDirectoryPath 'C:\Users\Steve\TEMP\input\basedir' -DestinationArchiveFileName 'C:\Users\Steve\TEMP\output2.zip' -CompressionLevel Fastest -IncludeBaseDirectory True
.NOTES
   
#>

function Compress-Zip
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $SourceDirectoryPath,
        [Parameter(Mandatory=$true)]
        [ValidateScript({
            if([System.IO.Path]::HasExtension($_))
            {
                $false
            }
            else
            {
                $true
            }
        })]
        $DestinationArchivePath,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Fastest', 'NoCompression', 'Optimal')]
        $CompressionLevel,
        [Parameter(Mandatory=$true)]
        [ValidateSet($true, $false)]
        $IncludeBaseDirectory
    )

    try
    {
        Add-Type -Assembly System.IO.Compression.FileSystems -ErrorAction Stop -ErrorVariable Assembly
    }
    catch
    {
        foreach ($Err in $Assembly | Select-Object -Unique)
        {
            Write-Warning $Err.Message
        }
        return
    }

    switch($CompressionLevel.ToLower())
    {
        fastest
        {
            [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectoryPath, "$DestinationArchivePath.zip", [System.IO.Compression.CompressionLevel]::Fastest, [System.Convert]::ToBoolean($IncludeBaseDirectory))
            break
        }
        nocompression
        {
            [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectoryPath, "$DestinationArchivePath.zip", [System.IO.Compression.CompressionLevel]::NoCompression, [System.Convert]::ToBoolean($IncludeBaseDirectory))
            break
        }
        optimal
        {
            [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectoryPath, "$DestinationArchivePath.zip", [System.IO.Compression.CompressionLevel]::Optimal, [System.Convert]::ToBoolean($IncludeBaseDirectory))
            break
        }
    }
}
