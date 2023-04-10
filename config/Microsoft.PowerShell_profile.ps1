Write-Host 'Welcome to PowerShell my friend.'

$Env:rd_coding = "$Env:OneDrive\home\07-coding"
$Env:rd_jp = "$Env:rd_coding\jupyter-notebooks"


function c {
    # Make it easy to switch locations
    param ([String]$Loc)

    switch ($Loc) {
        "p" { Set-Location $Env:OneDrive\home\01-projects }
        "a" { Set-Location $Env:OneDrive\home\02-areas }
        "r" { Set-Location $Env:OneDrive\home\03-resources }
        "x" { Set-Location $Env:OneDrive\home\05-archive }
        "s" { Set-Location $Env:SNIPPETS }
        "j" { Set-Location $Env:rd_jp }
        Default { Set-Location $Env:REPOS}
    }
}

function pjp { jupyter lab $Env:rd_jp }
function jp { jupyter lab . }

function Backup-RConfig {
    $cur_loc = Get-Location
    Set-Location $env:REPOS\app-config
    python -m config_manager.main
    Set-Location $cur_loc
}

Set-Alias -Name bus -Value Backup-RConfig

function lgrep {
    param ([String]$Path, [String]$Needle)
    Get-ChildItem $Path -Recurse -File | Where-Object {$(Get-Content $_) -match $Needle}
    if ($PSBoundParameters.ContainsKey('Path')) {
        # Write-Host "MyParameter was passed: $Path"
    }
    else {
        # Write-Host "MyParameter was not passed."
    }
}

function stylize {
    param ([string]$Text)
    "$($PSStyle.Foreground.BrightCyan)$($PSStyle.Bold)" + $Text + "$($PSStyle.Reset)"
}

# Add entries here, this will be used by the prompt function for setting the 
# prefix. Add subdirectories before parent directories.
$_match_paths = [ordered]@{
    "ρ" = [regex]::Escape($(Resolve-Path $Env:REPOS).Path);
    "δ" = [regex]::Escape($(Resolve-Path $Env:OneDrive).Path);
    "~" = [regex]::Escape($(Resolve-Path ~).Path);
}

function prompt {
    $current_location = $($ExecutionContext.SessionState.Path.CurrentLocation).Path

    $result = $null
    foreach ($key in $_match_paths.Keys) {
        $current_path = $($_match_paths[$key])
        $result = $current_location -match "^($current_path)(.*)`$"
        if ($result) {
            $matching_key = $key
            $trailing_part = $Matches[2]
            break
        }
    }

    if ($result) {
        $prompt_path = "${matching_key}${trailing_part}"
    } else {
        $prompt_path = $current_location
    }
    $full_prompt = "PS $prompt_path$('>' * ($NestedPromptLevel + 1)) "
    stylize $full_prompt
}


function touch {
    param ([Parameter(Mandatory=$true)][String]$FileName)
    $null > $FileName
}