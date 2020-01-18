# Copyright (c) Ryutaro Koma. All rights reserved.
# Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
# https://github.com/rykoma/Format-MarkdownTable

# Version 1.0

function Format-MarkdownTable {
    [CmdletBinding()]
    [Alias("fm")]

    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]
        $InputObject,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [Alias("ft")]
        [switch]
        $FormatTableStyle,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [switch]
        $HideStandardOutput,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [switch]
        $ShowMarkdown,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [switch]
        $DoNotCopyToClipboard,

        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $false)]
        [string[]]
        $Property = @("*")
    )
    
    Begin {
        ## Internal Function
        function EscapeMarkdown([object]$InputObject) {
            $Temp = ""

            if ($null -eq $InputObject) {
                return ""
            }
            elseif ($InputObject.GetType().BaseType -eq [System.Array]) {
                $Temp = "{" + [System.String]::Join(", ", $InputObject) + "}"
            }
            elseif ($InputObject.GetType() -eq [System.Collections.ArrayList]) {
                $Temp = "{" + [System.String]::Join(", ", $InputObject.ToArray()) + "}"
            }
            else {
                $Temp = $InputObject.ToString()
            }

            return $Temp.Replace("*", "\*")
        }

        if ($null -ne $InputObject -and $InputObject.GetType().BaseType -eq [System.Array]) {
            Write-Error "InputObject must not be System.Array. Don't use InputObject, but use the pipeline to pass the array object."
            $NeedToReturn = $true
            return
        }

        $LastCommandLine = (Get-PSCallStack)[1].Position.Text

        $Result = ""

        $HeadersForFormatTableStlye = New-Object System.Collections.Generic.List[string]
        $ContentsForFormatTableStyle = New-Object System.Collections.Generic.List[object]

        $TempOutputList = New-Object System.Collections.Generic.List[object]
    }

    Process {
        if ($NeedToReturn) { return }

        $CurrentObject = $null

        if ($_ -eq $null) {
            $CurrentObject = $InputObject | Select-Object -Property $Property
        }
        else {
            $CurrentObject = $_ | Select-Object -Property $Property
        }

        $Props = $CurrentObject | Get-Member -Name $Property -MemberType Property, NoteProperty

        if ($FormatTableStyle.IsPresent) {
            
            foreach ($Prop in $Props) {
                if ($HeadersForFormatTableStlye.Contains($Prop.Name) -eq $false) {
                    $HeadersForFormatTableStlye.Add($Prop.Name)
                }
            }

            $ContentsForFormatTableStyle.Add($CurrentObject)
        }
        else {
            $Output = "|Property|Value|`r`n"
            $Output += "|:--|:--|`r`n"

            $TempOutput = New-Object PSCustomObject

            foreach ($Prop in $Props) {
                $EscapedPropName = EscapeMarkdown($Prop.Name)
                $EscapedPropValue = EscapeMarkdown($CurrentObject.($($Prop.Name)))
                $Output += "|$EscapedPropName|$EscapedPropValue`r`n"
                $TempOutput | Add-Member -MemberType NoteProperty $Prop.Name -Value $CurrentObject.($($Prop.Name))
            }

            $Output += "`r`n"

            $Result += $Output

            $TempOutputList.Add($TempOutput)
        }
    }
    
    End {
        if ($NeedToReturn) { return }

        if ($FormatTableStyle.IsPresent) {
            $HeaderRow = "|"
            $SeparatorRow = "|"
            $ContentRow = "|"

            foreach ($Prop in $HeadersForFormatTableStlye) {
                $HeaderRow += "$(EscapeMarkdown($Prop))|"
                $SeparatorRow += ":--|"
                
            }

            foreach ($Content in $ContentsForFormatTableStyle) {
                $TempOutput = New-Object PSCustomObject

                foreach ($Prop in $HeadersForFormatTableStlye) {
                    $ContentRow += "$(EscapeMarkdown($Content.($($Prop))))|"

                    $TempOutput | Add-Member -MemberType NoteProperty $Prop -Value $Content.($($Prop))
                }
                
                $ContentRow += "`r`n"

                $TempOutputList.Add($TempOutput)
            }

            $Result = $HeaderRow + "`r`n" + $SeparatorRow + "`r`n" + $ContentRow
        }

        $ResultForConsole = $Result
        $Result = "**" + $LastCommandLine.Replace("*", "\*") + "**`r`n`r`n" + $Result

        if ($HideStandardOutput.IsPresent -eq $false) {
            if ($FormatTableStyle.IsPresent) {
                $TempOutputList | Format-Table * -AutoSize
            }
            else {
                $TempOutputList | Format-List *
            }
        }

        if ($ShowMarkdown.IsPresent) {
            Write-Output $ResultForConsole
        }

        if ($DoNotCopyToClipboard.IsPresent -eq $false) {
            Set-Clipboard $Result
            Write-Warning "Markdown text has been copied to the clipboard."
        }
    }
}

## one object

# Get-Mailbox rykoma | Format-MarkdownTable

# Get-Mailbox rykoma | Format-MarkdownTable -FormatTableStyle

# Get-Mailbox rykoma | Format-MarkdownTable UserPrincipalName, PrimaryS*, Alias

# Get-Mailbox rykoma | Format-MarkdownTable UserPrincipalName, PrimaryS*, Alias -FormatTableStyle

# Get-Mailbox rykoma | Format-MarkdownTable -Property UserPrincipalName,PrimarySmtpAddress,DisplayName

# Get-Mailbox rykoma | Format-MarkdownTable -Property UserPrincipalName,PrimarySmtpAddress,DisplayName -FormatTableStyle

# Format-MarkdownTable -InputObject:(Get-Mailbox rykoma) UserPrincipalName,PrimarySmtpAddress,Database

# Format-MarkdownTable -InputObject:(Get-Mailbox rykoma) UserPrincipalName,PrimarySmtpAddress,Database -FormatTableStyle

# Format-MarkdownTable -InputObject:(Get-Mailbox rykoma) -Property UserPrincipalName,PrimarySmtpAddress,Alias

# Format-MarkdownTable -InputObject:(Get-Mailbox rykoma) -Property UserPrincipalName,PrimarySmtpAddress,Alias -FormatTableStyle

## multi object

# Get-Mailbox remoteuser* | Format-MarkdownTable

# Get-Mailbox remoteuser* | Format-MarkdownTable -FormatTableStyle

# Get-Mailbox remoteuser* | Format-MarkdownTable UserPrincipalName, PrimaryS*, Alias

# Get-Mailbox remoteuser* | Format-MarkdownTable UserPrincipalName, PrimaryS*, Alias -FormatTableStyle

# Get-Mailbox remoteuser* | Format-MarkdownTable -Property UserPrincipalName,PrimarySmtpAddress,DisplayName

# Get-Mailbox remoteuser* | Format-MarkdownTable -Property UserPrincipalName,PrimarySmtpAddress,DisplayName -FormatTableStyle

# Format-MarkdownTable -InputObject:(Get-Mailbox remoteuser*) UserPrincipalName, PrimarySmtpAddress, Database

# Format-MarkdownTable -InputObject:(Get-Mailbox remoteuser*) UserPrincipalName,PrimarySmtpAddress,Database -FormatTableStyle

# Format-MarkdownTable -InputObject:(Get-Mailbox remoteuser*) -Property UserPrincipalName,PrimarySmtpAddress,Alias

# Format-MarkdownTable -InputObject:(Get-Mailbox remoteuser*) -Property UserPrincipalName,PrimarySmtpAddress,Alias -FormatTableStyle

## result contains array

# get-mailbox rykoma | Format-MarkdownTable email*

# get-mailbox rykoma | Format-MarkdownTable email* -FormatTableStyle