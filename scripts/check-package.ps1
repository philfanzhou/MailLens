[CmdletBinding()]
param([string]$Repository = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Test-MailLensPackage {
    param([string]$Root)
    $rootPath = [IO.Path]::GetFullPath($Root)
    $package = Join-Path $rootPath 'skills/mail-lens'
    $required = @('VERSION', 'LICENSE', 'scripts/install.ps1', 'skills/mail-lens/SKILL.md',
        'skills/mail-lens/scripts/outlook.ps1', 'skills/mail-lens/scripts/MailLens.Outlook.psm1',
        'skills/mail-lens/references/commands.md', 'skills/mail-lens/references/workflow.md')
    foreach ($relative in $required) {
        if (-not [IO.File]::Exists((Join-Path $rootPath $relative))) { throw "Required package file missing: $relative" }
    }
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    $texts = @{}
    $sourceCount = 0
    foreach ($file in @(Get-ChildItem -LiteralPath $rootPath -File -Recurse -Force | Where-Object { $_.FullName -notmatch '[\\/]\.git(?:[\\/]|$)' })) {
        $relative = $file.FullName.Substring($rootPath.Length + 1).Replace('\', '/')
        if ($file.Extension -notin @('.ps1', '.psm1', '.psd1', '.md', '.json', '.yml', '.yaml') -and $relative -notin @('VERSION', 'LICENSE', '.editorconfig', '.gitattributes', '.gitignore')) { continue }
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        $text = $utf8.GetString($bytes)
        if ($text.Length -gt 0 -and $text[0] -eq [char]0xfeff) { $text = $text.Substring(1) }
        if ($text.Length -eq 0 -or -not $text.EndsWith("`n") -or $text -match '(?m)[\t ]+\r?$') { throw "Empty file, missing final newline or trailing whitespace: $relative" }
        $texts[$relative] = $text
        if ($file.Extension -in @('.ps1', '.psm1', '.psd1')) {
            if ($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191) { throw "PowerShell UTF-8 BOM missing: $relative" }
            $tokens = $null; $errors = $null
            [void][Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            if ($errors.Count -gt 0) { throw "PowerShell parse error in ${relative}: $($errors[0].Message)" }
            foreach ($token in $tokens) {
                if ($token.Kind -eq 'Comment' -and $token.Text -match '[\u3400-\u9fff]') { throw "Non-English PowerShell comment: $relative" }
            }
            $sourceCount++
        }
        if ($file.Extension -eq '.md' -and $relative -notmatch '^skills/mail-lens/(?:SKILL\.md|references/)' -and $text -match '[\u3400-\u9fff]') { throw "Non-English project documentation: $relative" }
    }
    $version = $texts['VERSION'].Trim()
    if ($version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$') { throw 'Invalid root VERSION.' }
    $skill = $texts['skills/mail-lens/SKILL.md'].Replace("`r`n", "`n")
    $front = [regex]::Match($skill, '\A---\n(?<fields>(?:[^\n]+\n)+)---\n')
    if (-not $front.Success) { throw 'Missing or malformed skill frontmatter.' }
    $fields = @{}
    foreach ($line in $front.Groups['fields'].Value.TrimEnd("`n").Split("`n")) {
        $field = [regex]::Match($line, '^([a-z][a-z0-9-]*): ([^\r\n]+)$')
        if (-not $field.Success -or $fields.ContainsKey($field.Groups[1].Value)) { throw 'Malformed or duplicate flat frontmatter field.' }
        $value = $field.Groups[2].Value.Trim()
        if ($value -match '^[\[\]{|>]' -or [string]::IsNullOrWhiteSpace($value)) { throw 'Only nonempty flat frontmatter scalar values are supported.' }
        $fields[$field.Groups[1].Value] = $value
    }
    if ($fields['name'] -cne 'mail-lens' -or [string]::IsNullOrWhiteSpace($fields['description'])) { throw 'Invalid skill name or description.' }
    if ($fields.ContainsKey('version')) { throw 'Skill version must come only from root VERSION.' }
    foreach ($file in @(Get-ChildItem -LiteralPath $package -Filter '*.md' -Recurse)) {
        $text = $utf8.GetString([IO.File]::ReadAllBytes($file.FullName))
        foreach ($link in [regex]::Matches($text, '\[[^\]]*\]\(([^\s)]+)\)')) {
            $target = $link.Groups[1].Value
            if ($target -match '^(?:https?://|#)') { continue }
            if ($target -match '^[a-zA-Z][a-zA-Z0-9+.-]*:' -or [IO.Path]::IsPathRooted($target)) { throw 'Unsupported package resource link.' }
            $path = [IO.Path]::GetFullPath((Join-Path $file.DirectoryName ($target.Split('#')[0])))
            if (-not $path.StartsWith($package.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or -not [IO.File]::Exists($path)) { throw "Missing or escaping package resource: $target" }
        }
    }
    if ($sourceCount -lt 1) { throw 'No PowerShell source parsed.' }
    Write-Output "Package valid; VERSION=$version; parsed PowerShell files=$sourceCount."
}
if ($MyInvocation.InvocationName -ne '.') {
    try { Test-MailLensPackage $Repository; exit 0 }
    catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
}
