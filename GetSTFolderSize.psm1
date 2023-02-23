function Get-STFolderSize {
    <#
    .SYNOPSIS
        Gets folder sizes using COM and by default with a fallback to robocopy.exe, with the
        logging only option, which makes it not actually copy or move files, but just list them, and
        the end summary result is parsed to extract the relevant data.

    .DESCRIPTION
        There is a -ComOnly parameter for using only COM, and a -RoboOnly parameter for using only
        robocopy.exe with the logging only option.

        The robocopy output also gives a count of files and folders, unlike the COM method output.
        The default number of threads used by robocopy is 8, but I set it to 16 since this cut the
        run time down to almost half in some cases during my testing. You can specify a number of
        threads between 1-128 with the parameter -RoboThreadCount.

        Both of these approaches are apparently much faster than .NET and Get-ChildItem in PowerShell.

        The properties of the objects will be different based on which method is used, but
        the "TotalBytes" property is always populated if the directory size was successfully
        retrieved. Otherwise you should get a warning (and the sizes will be zero).

        Online documentation: https://www.powershelladmin.com/wiki/Get_Folder_Size_with_PowerShell,_Blazingly_Fast

        MIT license. http://www.opensource.org/licenses/MIT

        Copyright (C) 2015-present, Joakim Borger Svendsen
        All rights reserved.
        Svendsen Tech.

    .PARAMETER Path
        Path or paths to measure size of.

    .PARAMETER LiteralPath
        Path or paths to measure size of, supporting wildcard characters
        in the names, as with Get-ChildItem.

    .PARAMETER Precision
        Number of digits after decimal point in rounded numbers.

    .PARAMETER RoboOnly
        Do not use COM, only robocopy, for always getting full details.

    .PARAMETER ComOnly
        Never fall back to robocopy, only use COM.

    .PARAMETER RoboThreadCount
        Number of threads used when falling back to robocopy, or with -RoboOnly.
        Default: 16 (gave the fastest results during my testing).

    .PARAMETER ExcludeDirectory
        Names and paths for directories to exclude, as supported by the RoboCopy.exe /XD syntax.
        To guarantee its use, you need to use the parameter -RoboOnly. If COM is used, nothing
        is filtered out.
        
    .PARAMETER ExcludeFile
        Names/paths/wildcards for files to exclude, as supported by the RoboCopy.exe /XF syntax.
        To guarantee its use, you need to use the parameter -RoboOnly. If COM is used, nothing
        is filtered out.

    .EXAMPLE
        Import-Module .\Get-FolderSize.psm1
        PS C:\> 'C:\Windows', 'E:\temp' | Get-STFolderSize

    .EXAMPLE
        Get-STFolderSize -Path Z:\Database -Precision 2

    .EXAMPLE
        Get-STFolderSize -Path Z:\Users\*\AppData\*\Something -RoboOnly -RoboThreadCount 64

    .EXAMPLE
        Get-STFolderSize -Path "Z:\Database[0-9][0-9]" -RoboOnly

    .EXAMPLE
        Get-STFolderSize -LiteralPath 'A:\Full[HD]FloppyMovies' -ComOnly

        Supports wildcard characters in the path name with -LiteralPath.

    .EXAMPLE
        PS C:\temp> dir -dir | Get-STFolderSize | select Path, TotalMBytes | 
            Sort-Object -Descending totalmbytes

        Path                              TotalMBytes
        ----                              -----------
        C:\temp\PowerShellGetOld               1.2047
        C:\temp\git                            0.6562
        C:\temp\Docker                         0.3583
        C:\temp\MergeCsv                       0.1574
        C:\temp\testdir                        0.0476
        C:\temp\DotNetVersionLister            0.0474
        C:\temp\temp2                          0.0420
        C:\temp\WriteAscii                     0.0328
        C:\temp\tempbenchmark                  0.0257
        C:\temp\From PS Gallery Benchmark      0.0253
        C:\temp\RemoveOldFiles                 0.0238
        C:\temp\modules                        0.0234
        C:\temp\STDockerPs                     0.0216
        C:\temp\RandomData                     0.0205
        C:\temp\GetSTFolderSize                0.0198
        C:\temp\Benchmark                      0.0151

    .LINK 
        https://www.powershelladmin.com/wiki/Get_Folder_Size_with_PowerShell,_Blazingly_Fast

    #>
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param(
        [Parameter(ParameterSetName = "Path",
                   Mandatory = $True,
                   ValueFromPipeline = $True,
                   ValueFromPipelineByPropertyName = $True,
                   Position = 0)]
            [Alias('Name', 'FullName')]
            [String[]] $Path,
        [System.Int32] $Precision = 4,
        [Switch] $RoboOnly,
        [Switch] $ComOnly,
        [Parameter(ParameterSetName = "LiteralPath",
                   Mandatory = $true,
                   Position = 0)] [String[]] $LiteralPath,
        [ValidateRange(1, 128)] [Byte] $RoboThreadCount = 16,
        [String[]] $ExcludeDirectory = @(),
        [String[]] $ExcludeFile = @())
    
    Begin {
    
        $DefaultProperties = 'Path', 'TotalBytes', 'TotalMBytes', 'TotalGBytes', 'TotalTBytes', 'DirCount', 'FileCount',
            'DirFailed', 'FileFailed', 'TimeElapsed', 'StartedTime', 'EndedTime'
        if (($ExcludeDirectory.Count -gt 0 -or $ExcludeFile.Count -gt 0) -and -not $ComOnly) {
            $DefaultProperties += @("CopiedDirCount", "CopiedFileCount", "CopiedBytes", "SkippedDirCount", "SkippedFileCount", "SkippedBytes")
        }
        if ($RoboOnly -and $ComOnly) {
            Write-Error -Message "You can't use both -ComOnly and -RoboOnly. Default is COM with a fallback to robocopy." -ErrorAction Stop
        }
        if (-not $RoboOnly) {
            $FSO = New-Object -ComObject Scripting.FileSystemObject -ErrorAction Stop
        }
        function Get-RoboFolderSizeInternal {
            [CmdletBinding()]
            param(
                # Paths to report size, file count, dir count, etc. for.
                [String[]] $Path,
                [System.Int32] $Precision = 4)
            begin {
                if (-not (Get-Command -Name robocopy -ErrorAction SilentlyContinue)) {
                    Write-Warning -Message "Fallback to robocopy failed because robocopy.exe could not be found. Path '$p'. $([datetime]::Now)."
                    return
                }
            }
            process {
                foreach ($p in $Path) {
                    Write-Verbose -Message "Processing path '$p' with Get-RoboFolderSizeInternal. $([datetime]::Now)."
                    $RoboCopyArgs = @("/L","/S","/NJH","/BYTES","/FP","/NC","/NDL","/TS","/XJ","/R:0","/W:0","/MT:$RoboThreadCount")
                    if ($ExcludeDirectory.Count -gt 0) {
                        $RoboCopyArgs += @(@("/XD") + @($ExcludeDirectory))
                    }
                    if ($ExcludeFile.Count -gt 0) {
                        $RoboCopyArgs += @(@("/XF") + @($ExcludeFile))
                    }
                    [DateTime] $StartedTime = [DateTime]::Now
                    [String] $Summary = robocopy $p NULL $RoboCopyArgs | Select-Object -Last 8
                    [DateTime] $EndedTime = [DateTime]::Now
                    #[String] $DefaultIgnored = '(?:\s+[\-\d]+){3}'
                    <# only used for systems installed in german
                    [regex] $HeaderRegex = '\s+Insgesamt\s*Kopiert.*?(?:bersprungen).*?(?:bereinstimmung)\s+FEHLER\s+Extras'
                    [regex] $DirLineRegex = 'Verzeich.\s*:\s*(?<DirCount>\d+)(?:\s+\d+){3}\s+(?<DirFailed>\d+)\s+\d+'
                    [regex] $FileLineRegex = 'Dateien\s*:\s*(?<FileCount>\d+)(?:\s+\d+){3}\s+(?<FileFailed>\d+)\s+\d+'
                    [regex] $BytesLineRegex = 'Bytes\s*:\s*(?<ByteCount>\d+)(?:\s+\d+){3}\s+(?<BytesFailed>\d+)\s+\d+'
                    [regex] $TimeLineRegex = 'Zeiten\s*:\s*(?<TimeElapsed>\d+).*'
                    [regex] $EndedLineRegex = 'Beendet\s*:\s*(?<EndedTime>.+)'
                    #>
                    
                    [Regex] $HeaderRegex = '\s+Total\s*Copied\s+Skipped\s+Mismatch\s+FAILED\s+Extras'
                    [Regex] $DirLineRegex = "Dirs\s*:\s*(?<DirCount>\d+)(?<CopiedDirCount>\s+[\-\d]+)(?<SkippedDirCount>\s+[\-\d]+)(?:\s+[\-\d]+)\s+(?<DirFailed>\d+)\s+[\-\d]+"
                    [Regex] $FileLineRegex = "Files\s*:\s*(?<FileCount>\d+)(?<CopiedFileCount>\s+[\-\d]+)(?<SkippedFileCount>\s+[\-\d]+)(?:\s+[\-\d]+)\s+(?<FileFailed>\d+)\s+[\-\d]+"
                    [Regex] $BytesLineRegex = "Bytes\s*:\s*(?<ByteCount>\d+)(?<CopiedBytes>\s+[\-\d]+)(?<SkippedBytes>\s+[\-\d]+)(?:\s+[\-\d]+)\s+(?<BytesFailed>\d+)\s+[\-\d]+"
                    [Regex] $TimeLineRegex = "Times\s*:\s*.*"
                    [Regex] $EndedLineRegex = "Ended\s*:\s*(?<EndedTime>.+)"
                    if ($Summary -match "$HeaderRegex\s+$DirLineRegex\s+$FileLineRegex\s+$BytesLineRegex\s+$TimeLineRegex\s+$EndedLineRegex") {
                        New-Object PSObject -Property @{
                            Path = $p
                            TotalBytes = [Decimal] $Matches['ByteCount']
                            TotalMBytes = [Math]::Round(([Decimal] $Matches['ByteCount'] / 1MB), $Precision)
                            TotalGBytes = [Math]::Round(([Decimal] $Matches['ByteCount'] / 1GB), $Precision)
                            TotalTBytes = [Math]::Round(([Decimal] $Matches['ByteCount'] / 1TB), $Precision)
                            BytesFailed = [Decimal] $Matches['BytesFailed']
                            DirCount = [Decimal] $Matches['DirCount']
                            FileCount = [Decimal] $Matches['FileCount']
                            DirFailed = [Decimal] $Matches['DirFailed']
                            FileFailed  = [Decimal] $Matches['FileFailed']
                            TimeElapsed = [Math]::Round([Decimal] ($EndedTime - $StartedTime).TotalSeconds, $Precision)
                            StartedTime = $StartedTime
                            EndedTime   = $EndedTime
                            CopiedDirCount = [Decimal] $Matches['CopiedDirCount']
                            CopiedFileCount = [Decimal] $Matches['CopiedFileCount']
                            CopiedBytes = [Decimal] $Matches['CopiedBytes']
                            SkippedDirCount = [Decimal] $Matches['SkippedDirCount']
                            SkippedFileCount = [Decimal] $Matches['SkippedFileCount']
                            SkippedBytes = [Decimal] $Matches['SkippedBytes']

                        } | Select-Object -Property $DefaultProperties
                    }
                    else {
                        Write-Warning -Message "Path '$p' output from robocopy was not in an expected format."
                    }
                }
            }
        }
    
    }
    
    Process {
        if ($PSCmdlet.ParameterSetName -eq "Path") {
            $Paths = @(Resolve-Path -Path $Path | Select-Object -ExpandProperty ProviderPath -ErrorAction SilentlyContinue)
        }
        else {
            $Paths = @(Get-Item -LiteralPath $LiteralPath | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue)
        }
        foreach ($p in $Paths) {
            Write-Verbose -Message "Processing path '$p'. $([DateTime]::Now)."
            if (-not (Test-Path -LiteralPath $p -PathType Container)) {
                Write-Warning -Message "$p does not exist or is a file and not a directory. Skipping."
                continue
            }
            # We know we can't have -ComOnly here if we have -RoboOnly.
            if ($RoboOnly) {
                Get-RoboFolderSizeInternal -Path $p -Precision $Precision
                continue
            }
            $ErrorActionPreference = 'Stop'
            try {
                $StartFSOTime = [DateTime]::Now
                $TotalBytes = $FSO.GetFolder($p).Size
                $EndFSOTime = [DateTime]::Now
                if ($null -eq $TotalBytes) {
                    if (-not $ComOnly) {
                        Get-RoboFolderSizeInternal -Path $p -Precision $Precision
                        continue
                    }
                    else {
                        Write-Warning -Message "Failed to retrieve folder size for path '$p': $($Error[0].Exception.Message)."
                    }
                }
            }
            catch {
                if ($_.Exception.Message -like '*PERMISSION*DENIED*') {
                    if (-not $ComOnly) {
                        Write-Verbose "Caught a permission denied. Trying robocopy."
                        Get-RoboFolderSizeInternal -Path $p -Precision $Precision
                        continue
                    }
                    else {
                        Write-Warning "Failed to process path '$p' due to a permission denied error: $($_.Exception.Message)"
                    }
                }
                Write-Warning -Message "Encountered an error while processing path '$p': $($_.Exception.Message)"
                continue
            }
            $ErrorActionPreference = 'Continue'
            New-Object PSObject -Property @{
                Path = $p
                TotalBytes = [Decimal] $TotalBytes
                TotalMBytes = [Math]::Round(([Decimal] $TotalBytes / 1MB), $Precision)
                TotalGBytes = [Math]::Round(([Decimal] $TotalBytes / 1GB), $Precision)
                TotalTBytes = [Math]::Round(([Decimal] $TotalBytes / 1TB), $Precision)
                BytesFailed = $null
                DirCount = $null
                FileCount = $null
                DirFailed = $null
                FileFailed  = $null
                TimeElapsed = [Math]::Round(([Decimal] ($EndFSOTime - $StartFSOTime).TotalSeconds), $Precision)
                StartedTime = $StartFSOTime
                EndedTime = $EndFSOTime
            } | Select-Object -Property $DefaultProperties
        }
    }
    
    End {
        
        if (-not $RoboOnly) {
            [Void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($FSO)
        }
        
        [System.GC]::Collect()
        #[System.GC]::WaitForPendingFinalizers()
    
    }

}
