# Copyright (c) Ryutaro Koma. All rights reserved.
# Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
# https://github.com/rykoma/Format-MarkdownTable

# Version 1.3.1

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
        $Property = @()
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
            elseif (Get-Member -InputObject $InputObject -Name ToString -MemberType Method) {
                $Temp = $InputObject.ToString()
            }
            else {
                $Temp = ""
            }

            return $Temp.Replace("*", "\*")
        }

        function GetDefaultDisplayProperty([object]$InputObject) {
            try {
                if ($null -eq $InputObject) {
                    return @("*")
                }
    
                $DataType = ($InputObject | Get-Member)[0].TypeName
    
                if ($DataType.StartsWith("Selected.")) {
                    return @("*")
                }            
                elseif ($DataType.StartsWith("Deserialized.")) {
                    $DataType = $DataType.Trim("Deserialized.")
                }
    
                $FormatData = Get-FormatData -TypeName $DataType -ErrorAction SilentlyContinue
    
                if ($null -eq $FormatData) {
                    return @("*")
                }
    
                return $FormatData.FormatViewDefinition.Control.Rows.Columns.DisplayEntry.Value
            }
            catch {
                return @("*")
            }
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
            if ($FormatTableStyle.IsPresent -and (($Property.Length -eq 0) -or ($Property.Length -eq 1 -and $Property[0] -eq ""))) {
                $Property = GetDefaultDisplayProperty($InputObject)
            }
            elseif (($Property.Length -eq 0) -or ($Property.Length -eq 1 -and $Property[0] -eq "")) {
                $Property = @("*")
            }

            $CurrentObject = $InputObject | Select-Object -Property $Property
        }
        else {
            if ($FormatTableStyle.IsPresent -and (($Property.Length -eq 0) -or ($Property.Length -eq 1 -and $Property[0] -eq ""))) {
                $Property = GetDefaultDisplayProperty($_)
            }
            elseif (($Property.Length -eq 0) -or ($Property.Length -eq 1 -and $Property[0] -eq "")) {
                $Property = @("*")
            }

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
            $ContentRow = ""

            foreach ($Prop in $HeadersForFormatTableStlye) {
                $HeaderRow += "$(EscapeMarkdown($Prop))|"
                $SeparatorRow += ":--|"
                
            }

            foreach ($Content in $ContentsForFormatTableStyle) {
                $TempOutput = New-Object PSCustomObject
                $ContentRow += "|"

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