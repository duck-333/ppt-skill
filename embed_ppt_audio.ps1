param(
    [Parameter(Mandatory = $true)]
    [string]$InputPptx,

    [Parameter(Mandatory = $true)]
    [string]$AudioDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputPptx
)

$ErrorActionPreference = "Stop"

function Invoke-ComRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        try {
            return & $Action
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 9) {
                throw
            }
            Start-Sleep -Milliseconds (500 + 250 * $attempt)
        }
    }
}

function Get-OfficeRgb {
    param(
        [int]$Red,
        [int]$Green,
        [int]$Blue
    )

    return $Red + ($Green * 256) + ($Blue * 65536)
}

function Find-ShapeByName {
    param(
        $Slide,
        [string]$Name
    )

    foreach ($shape in $Slide.Shapes) {
        if ($shape.Name -eq $Name) {
            return $shape
        }
    }
    return $null
}

function Ensure-PronunciationButton {
    param(
        $Slide,
        [string]$Name,
        [string]$Label,
        [double]$Left,
        [int]$FillRgb
    )

    $shape = Find-ShapeByName -Slide $Slide -Name $Name
    if ($null -eq $shape) {
        # msoShapeRoundedRectangle = 5. PowerPoint COM geometry uses points.
        $shape = Invoke-ComRetry {
            $Slide.Shapes.AddShape(5, $Left, 135.45, 51.18, 51.18)
        }
        $shape.Name = $Name
    }

    $shape.Left = $Left
    $shape.Top = 135.45
    $shape.Width = 51.18
    $shape.Height = 51.18
    $shape.Fill.ForeColor.RGB = $FillRgb
    $shape.Line.ForeColor.RGB = Get-OfficeRgb -Red 255 -Green 255 -Blue 255
    $shape.Line.Weight = 2
    $shape.TextFrame.TextRange.Text = $Label
    $shape.TextFrame.TextRange.Font.Name = "Microsoft YaHei"
    $shape.TextFrame.TextRange.Font.Size = 34
    $shape.TextFrame.TextRange.Font.Bold = -1
    $shape.TextFrame.TextRange.Font.Color.RGB =
        Get-OfficeRgb -Red 255 -Green 255 -Blue 255
    # ppAlignCenter = 2, msoAnchorMiddle = 3.
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $shape.TextFrame.VerticalAnchor = 3
    return $shape
}

function Get-WordFromSlide {
    param($Slide)

    $wordShape = Find-ShapeByName -Slide $Slide -Name "WORD_NAME"
    if ($null -eq $wordShape -or $wordShape.HasTextFrame -ne -1) {
        return ""
    }

    $word = [string]$wordShape.TextFrame.TextRange.Text
    return $word.Trim()
}

$inputPath = (Resolve-Path -LiteralPath $InputPptx).Path
$audioPath = (Resolve-Path -LiteralPath $AudioDir).Path
$outputPath = [IO.Path]::GetFullPath($OutputPptx)

if (Test-Path -LiteralPath $outputPath) {
    throw "Output file already exists. Choose a new path: $outputPath"
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$presentation = $null
$embedded = 0
$wordsProcessed = 0
$missing = [System.Collections.Generic.List[string]]::new()

try {
    $powerPoint.Visible = -1
    Start-Sleep -Seconds 2
    $presentation = Invoke-ComRetry {
        $powerPoint.Presentations.Open($inputPath, 0, 0, 0)
    }

    foreach ($slide in $presentation.Slides) {
        $word = Get-WordFromSlide -Slide $slide
        if ([string]::IsNullOrWhiteSpace($word) -or $word -notmatch "^[A-Za-z'-]+$") {
            continue
        }

        $wordsProcessed++
        $stem = $word.ToLowerInvariant()
        $ukFile = Join-Path $audioPath "${stem}_uk.wav"
        $usFile = Join-Path $audioPath "${stem}_us.wav"

        foreach ($variant in @(
            @{
                Code = "UK"
                Label = [string][char]0x82F1
                File = $ukFile
                Left = 782.38
                Fill = Get-OfficeRgb -Red 21 -Green 94 -Blue 239
            },
            @{
                Code = "US"
                Label = [string][char]0x7F8E
                File = $usFile
                Left = 854.02
                Fill = Get-OfficeRgb -Red 255 -Green 0 -Blue 0
            }
        )) {
            if (-not (Test-Path -LiteralPath $variant.File)) {
                $missing.Add([IO.Path]::GetFileName($variant.File))
                continue
            }

            $buttonName = "PRON_$($variant.Code)_$stem"
            $button = Ensure-PronunciationButton `
                -Slide $slide `
                -Name $buttonName `
                -Label $variant.Label `
                -Left $variant.Left `
                -FillRgb $variant.Fill

            # PowerPoint must import the sound before Action is set to
            # ppActionNone. This produces relationships/audio and a:snd.
            $action = $button.ActionSettings.Item(1)
            $sound = $action.SoundEffect
            Invoke-ComRetry {
                $sound.ImportFromFile([IO.Path]::GetFullPath($variant.File))
            }
            $action.Action = 0
            $embedded++
        }
    }

    if ($missing.Count -gt 0) {
        $uniqueMissing = $missing | Sort-Object -Unique
        throw "Missing audio files ($($uniqueMissing.Count)): $($uniqueMissing -join ', ')"
    }

    Invoke-ComRetry {
        # ppSaveAsOpenXMLPresentation = 24.
        $presentation.SaveAs($outputPath, 24)
    }
    Start-Sleep -Seconds 3
}
finally {
    if ($null -ne $presentation) {
        Invoke-ComRetry { $presentation.Close() }
    }
    Invoke-ComRetry { $powerPoint.Quit() }
}

[pscustomobject]@{
    Output = $outputPath
    WordsProcessed = $wordsProcessed
    EmbeddedAudioFiles = $embedded
    ExpectedAudioFiles = $wordsProcessed * 2
} | Format-List
