$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

foreach ($directory in @('Public', 'Private', '.')) {

    $path = Join-Path -Path $PSScriptRoot -ChildPath $directory
    if (Test-Path $path) {
        Get-ChildItem -Path $path -Filter "*.ps1" | ForEach-Object { . $_.FullName }
    }
}