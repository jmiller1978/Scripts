#requires -version 2

<#
    .SYNOPSIS
        Organizes movies based on date they were taken.
    .DESCRIPTION
        This script will take any number of movies and organize them into folders that will be
        created based on the years and months they were filmed.
        
        Author: James E. Miller
        Version: 1.0.20130821 
    .PARAMETER Source
        Source folder with original movie files.
    .PARAMETER Target
        Target folder where the movies should be organized.
    .PARAMETER Recurse
        Include the sub-folders of the source folder as well.
    .EXAMPLE
        .\Sort-Movies.ps1 -Source "C:\Videos"

        In this example, only the source folder is specified which will have the script create/use
        the dated folders under the source.
    .EXAMPLE
        .\Sort-Movies.ps1 -Source "C:\Videos" -Target "D:\Videos"

        For this example, both the source and target folders are specified which will have the script
        create/use the dated folders under the target and move the movies there.
    .EXAMPLE
        .\Sort-Movies.ps1 -Source "C:\Videos" -Target "D:\Videos" -Recurse

        This does the same as Example #2 except it'll grab all the movies under the source and all
        sub-folders.
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Source,
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Target,
    [switch]$Recurse
)

#$ErrorActionPreference = "SilentlyContinue"

if ($Recurse) {
    $movies = Get-ChildItem -Recurse | ? { $_.Extension -like ".mov" -or $_.Extension -like ".mp4" }
}
else {
    $movies = Get-ChildItem | ? { $_.Extension -like ".mov" -or $_.Extension -like ".mp4" }
}

$total = $movies.Count
$i = 1

if ($movies) {
    $movies | % {
        # Display progress bar.
        Write-Progress -Activity "Sorting $i out of $total movies" -Status $_.Name -PercentComplete (($i/$total) * 100)
        # Attempt to read date movie was taken and move to respective folder.
        try {
            $filename = $_.FullName
            $oldfilename = $_.BaseName
            $extension = $_.Extension
            $datetaken = [datetime]$_.LastWriteTime

            $folder = (Get-Date $datetaken -Format "yyyy-MM (MMM)").ToString() # Ex: 2013-08 (Aug)

            
            $newfilename = (Get-Date $datetaken -Format MMddyyyy-HHmmss).ToString()
            $newfilename = $newfilename + "($oldfilename)$extension"

            if ($Target) {
                if (!(Test-Path $Target\$folder -PathType Container)) {
                    New-Item $Target\$folder -ItemType Directory | Out-Null
                }
                Move-Item $filename -Destination "$Target\$folder\$newfilename"    
            }
            else {
                if (!(Test-Path $Source\$folder -PathType Container)) {
                    New-Item $Source\$folder -ItemType Directory | Out-Null
                }
                Move-Item $filename -Destination "$Source\$folder\$newfilename" 
            }
        }
        # If any errors occured, do nothing. Errors are likely caused by duplicates.
        catch {
            #do nothing
        }
        # Regardless of success, increment progress counter.
        finally {
            $i++
        }
    }
}

