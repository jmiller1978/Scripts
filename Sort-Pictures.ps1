#requires -version 2

<#
    .SYNOPSIS
        Organizes photos based on date they were taken.
    .DESCRIPTION
        This script will takeany number of pictures and organize them into folders that will be
        created based on the years and months the pictures were taken.
        
        Author: James E. Miller
        Version: 1.0.20130821 
    .PARAMETER Source
        Source folder with original picture files.
    .PARAMETER Target
        Target folder where the pictures should be organized.
    .PARAMETER Recurse
        Include the sub-folders of the source folder as well.
    .EXAMPLE
        .\Sort-Pictures.ps1 -Source "C:\Pictures"

        In this example, only the source folder is specified which will have the script create/use
        the dated folders under the source.
    .EXAMPLE
        .\Sort-Pictures.ps1 -Source "C:\Pictures" -Target "D:\Photos"

        For this example, both the source and target folders are specified which will have the script
        create/use the dated folders under the target and move the pictures there.
    .EXAMPLE
        .\Sort-Pictures.ps1 -Source "C:\Pictures" -Target "D:\Photos" -Recurse

        This does the same as Example #2 except it'll grab all the pictures under the source and all
        sub-folders.
    .LINK
        http://http://www.exiv2.org/tags.html
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Source,
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Target,
    [switch]$Recurse
)

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Text")
$ErrorActionPreference = "SilentlyContinue"

if ($Recurse) {
    $pictures = Get-ChildItem -Path $Source -Filter *.jpg -Recurse
}
else {
    $pictures = Get-ChildItem -Path $Source -Filter *.jpg
}

$total = $pictures.Count
$i = 1

if ($pictures) {
    $pictures | % {
        # Display progress bar.
        Write-Progress -Activity "Sorting $i out of $total pictures" -Status $_.Name -PercentComplete (($i/$total) * 100)
        # Attempt to read date picture was taken and move to respective folder.
        try {
            $filename = $_.FullName
            $photo = [System.Drawing.Image]::FromFile($filename)
            
            $datevalue = $photo.GetPropertyItem(36867) #date picture was taken
            $datevalue = (New-Object System.Text.UTF8Encoding).GetString($datevalue.Value)
            $datetaken = [datetime]::ParseExact($datevalue,"yyyy:MM:dd HH:mm:ss`0",$null)

            $folder = (Get-Date $datetaken -Format "yyyy-MM (MMM)").ToString() # Ex: 2013-08 (Aug)

            $oldfilename = $_.BaseName
            $newfilename = (Get-Date $datetaken -Format MMddyyyy-HHmmss).ToString()
            $newfilename = $newfilename + "($oldfilename).jpg"

            $photo.Dispose()

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
        # If any errors occured, release photo. Errors are likely caused by EXIF data missing and/or duplicate existed.
        catch {
            if ($photo) { $photo.Dispose() }
        }
        # Regardless of success, increment progress counter.
        finally {
            $i++
        }
    }
}

