# Format-MarkdownTable
Format PowerShell object to Markdown table

## Usage

1. Download Format-MarkdownTable.ps1 from the [releases](https://github.com/rykoma/Format-MarkdownTable/releases) page
2. Right-click the ps1 file and click [Property]
3. In the [General] tab, if you see "This file came from another computer and might be blocked to help protect this computer", check [Unblock]
4. Start Windows PowerShell
5. Dot source the Format-MarkdownTable.ps1

```powershell
. <path to Format-MarkdownTable.ps1>
```

e.g.

```
. C:\scripts\Format-MarkdownTable.ps1
```

## Example

```powershell
Get-ChildItem c:\ | Format-MarkdownTable Name, LastWriteTime, Mode
```

This example returns a summary of the child items in C drive, and markdown text will be copied to the clipboard. By default, the table will be formatted like the Format-List command style.

```powershell
Get-ChildItem c:\ | Format-MarkdownTable Name, LastWriteTime, Mode -FormatTableStyle
```

This example returns a summary of the child items in C drive, and markdown text will be copied to the clipboard. The table will be formatted like the Format-Table command style.

## Alias

```powershell
Get-ChildItem c:\ | fm Name, LastWriteTime, Mode -ft 
```

"fm" is an alias for the Format-MarkdownTable command. "ft" is an alias for the FormatTableStyle switch.

## Switch

```powershell
Get-ChildItem c:\ | fm Name, LastWriteTime, Mode -HideStandardOutput
```

Standard output will be hidden by using HideStandardOutput switch. Markdown text will be copied to the clipboard.

```powershell
Get-ChildItem c:\ | fm Name, LastWriteTime, Mode -HideStandardOutput -ShowMarkdown
```

Markdown text will be displayed in the console by using ShowMarkdown switch. It will also be copied to the clipboard.

```powershell
Get-ChildItem c:\ | fm Name, LastWriteTime, Mode -HideStandardOutput -ShowMarkdown -DoNotCopyToClipboard
```

By using HideStandardOutput switch, the markdown text will not be copied to the clipboard.