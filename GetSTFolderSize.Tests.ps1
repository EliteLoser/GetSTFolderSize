#requires -version 3


# Pester 4.x tests for Get-STFolderSize in the GetSTFolderSize module.
# Joakim Borger Svendsen. Started 2018-12-17. Svendsen Tech.

Import-Module -Name Pester -ErrorAction Stop
#Import-Module -Name GetSTFolderSize -ErrorAction Stop
ipmo ..\GetSTFolderSize -Force

# This is mine also, in the PowerShell Gallery. Just crudely making it available for the tests.
Save-Module -Name RandomData -Path $Env:Temp -Force
Import-Module "$Env:Temp\RandomData" -Force

Describe Get-STFolderSize {
    
    # Set up a test directory structure.
    $STBaseDir = "$Env:Temp\GetSTFolderSize.Tests.Svendsen.Tech"
    # Clean things up in advance, just in case.
    Remove-Item -LiteralPath $STBaseDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path "$STBaseDir\sub1\sub2\sub3" `
        -ItemType Directory -Force | Out-Null

    New-RandomData -Path $STBaseDir -Count 5 -Size 1024 -LineLength 128
    
    It "Calculates the size of a directory with 5x 1024-byte files correctly using -ComOnly." {
        ($ComResult = (Get-STFolderSize -Path $STBaseDir -ComOnly)).TotalBytes | Should -Be (5*1024)
    }
    
    It "Calculates the size of a directory with 5x 1024-byte files correctly using -RoboOnly." {
        ($RoboResult = (Get-STFolderSize -Path $STBaseDir -RoboOnly)).TotalBytes | Should -Be (5*1024)
    }

    It "Gives the same result with COM and RoboCopy" {
        [Int] $ComResult -eq [Int] $RoboResult | Should -Be $True
    }

    # Add subdirectories with files and test recursively from now on.
    New-RandomData -Path "$STBaseDir\sub1" -Count 5 -Size 1024 -LineLength 128
    New-RandomData -Path "$STBaseDir\sub1\sub2" -Count 5 -Size 1024 -LineLength 128
    New-RandomData -Path "$STBaseDir\sub1\sub2\sub3" -Count 5 -Size 1024 -LineLength 128
    
    It "Calculates the size of a directory tree with 5x4x 1024-byte files correctly using -ComOnly." {
        ($ComResult = (Get-STFolderSize -Path $STBaseDir -ComOnly)).TotalBytes | Should -Be (20*1024)
    }
    
    It "Calculates the size of a directory tree with 5x4x 1024-byte files correctly using -RoboOnly." {
        ($RoboResult = (Get-STFolderSize -Path $STBaseDir -RoboOnly)).TotalBytes | Should -Be (20*1024)
    }

    It "Gives the same result with COM and RoboCopy with a directory tree structure" {
        [Int] $ComResult -eq [Int] $RoboResult | Should -Be $True
    }

    # Add files with a different extension for exclusion tests.
    New-RandomData -Path $STBaseDir -Count 5 -BaseName ignored -Extension .ignored -Size 1024 -LineLength 128

    It "Ignores *.ignored files that should be 5 kB in size in the test directory tree." {
        $RoboResult = Get-STFolderSize -Path $STBaseDir -RoboOnly -ExcludeFile *.ignored
        $RoboResult.SkippedBytes | Should -Be (5*1024)
        $RoboResult.SkippedFileCount | Should -Be 5
    }

    It "Ignores a specified directory name (with files containing 5 kB worth of files)." {
        $RoboResult = Get-STFolderSize -Path $STBaseDir -RoboOnly -ExcludeDirectory sub3
        $RoboResult.SkippedDirCount | Should -Be 1
    }

}
