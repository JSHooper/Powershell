#Requires -Version 4.0
#Requires -Modules ActiveDirectory

<#
MIT License

Copyright (c) 2020 James Hooper

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

#Global Imports
Import-Module ActiveDirectory

# Global Variables used throughout Script
# TenantName to be set to your tenant name: e.g. contoso
$TenantName = "" #Please fill in with your tenant name

#OneDriveSite is the generic OneDrive location for the tenant (This is used in the Get-OneDriveLink function)
# $OneDriveSite = "https://<TenantName>-my.sharepoint.com/personal/"
$OneDriveSite = "https://$TenantName-my.sharepoint.com/personal/"

# Functions Section

Function Get-OneDriveLink()
{
<#
    Example: Get-OneDriveLink -UserPrincipalName "Joe.Bloggs@contoso.com"
    Output: https://<TenantName>-my.sharepoint.com/personal/Joe_Bloggs_contoso_com
#>
    [OutputType([String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [String]$UserPrincipalName
    )

    $ODUPN = $UserPrincipalName.Replace("@","_").Replace(".","_")
    
    return "$OneDriveSite$ODUPN"
    
}

Function New-SPMTCSV()
{
<#
    Example: New-SPMTCSV -Source "HomeDirectoryLocation" -Target "OnedriveURL" -CSVOutput "CSVFile"

    This function is to help with the creation of the SPMT CSV file to allow multiple users to be migrated using the SPMT tool
#>
    param
    (
        [Parameter(Mandatory = $true)]
        [String]$Source,
        [Parameter(Mandatory = $true)]
        [String]$Target,
        [Parameter(Mandatory = $true)]
        [String]$CSVOutput
    )
    
    "$Source,,,$Target,Documents," | Out-file $CSVOutput -Append -Encoding ascii
}

Function Get-HomeDirectory()
{
<#
    Example: Get-HomeDirectory -UserPrincipalName Joe.Bloggs@domain.com
#>
    param
    (
        [Parameter(Mandatory = $true)]
        [String]$UserPrincipalName
    )

    Try{
        $HomeDir = (Get-AdUser -Filter {UserPrincipalName -eq $UserPrincipalName} -Properties HomeDirectory).HomeDirectory
    }catch{
        Write-Error $_
    }

    if(!([string]::IsNullOrEmpty($HomeDir))){
        return $HomeDir
    }else{
        return $null
    }

    
}

Function Get-FileName()
{
<#
    Creates an open dialog box to pick the CSv file for the users
#>

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = ".\"
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
 
}

# Function Section End

#Script Start

$CSV = Get-FileName
$CSVOutput = "$($CSV.SubString(0,$CSV.Length-4))-SPMT.csv"
$Users = Import-Csv $CSV
ForEach($User in $Users.UserPrincipalName)
{
    New-SPMTCSV -Source (Get-HomeDirectory -UserPrincipalName $User) -Target (Get-OneDriveLink -UserPrincipalName $User) -CSVOutput $CSVOutput
}


#Script End