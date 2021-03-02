param(
    [int]$year = $(throw "-year is required."),
    [Parameter(mandatory = $false)] [int]$layers = 0,
    [string]$sName = $(throw "-sName is required."),
    [int]$sPos = $(throw "-sPos = slide position"),
    [string]$type = $(throw "-type must be year, week or month"),
    [switch]$div = $false    
)

. ".\gen-functions.ps1"
. ".\style-functions.ps1"

$colors = @{"year" = (1,0,14); "month" = (5, 0.6, 1); "week" = (6, 0.6, 1); "day" = (7, 0.6, 1) }
$color, $brightness, $font = $colors[$type]
$labels = $divs[1], $divs[2], $divs[3], $divs[4]

$slide = $presentation.Slides.AddSlide($sPos, $layouts.Item(7))
$slide.Name = $sName

$holIdx, $currHols = Get-holIdx 1 1 $year $hols
for ($y = $year; $y -le ($year + $layers); $y++) {  
      
    $title = $slide.Shapes(1)
    $title.TextFrame.TextRange.Text = [string]$y + " Calendar"
    $title.Name = $sName

    $Top = $title.Top + $title.Height + 0.35 * $cm
    $Left = $title.Left
    $vDiv = ((26 - (5.6 * 4)) / 3) * $cm
    $hDiv = ((18 - (5.6 * 3)) / 2) * $cm

    for ($mIdx = 1; $mIdx -le $months.Count; $mIdx++) {
        $m = $months[$mIdx - 1]

        $nTop = $Top + [Math]::Floor(($mIdx - 1) / 3) * (5.6 * $cm + $vDiv)
        $nLeft = $Left + (($mIdx - 1) % 3) * ((5.6 * $cm) + $hDiv)
    
        $box = $slide.Shapes.AddTextbox(1, $nLeft, $nTop, 5.6 * $cm, 0.5 * $cm)
        use-miniCalTitle $box $m $y

        $templates.templates.Shapes("calendar-" + $type).Copy()
        [void] $slide.Shapes.Paste()
        $cal = $slide.Shapes($slide.Shapes.Count)
        use-miniCal $box $m $y $cal

        $row = 1
        for ($col = 1; $col -le 7; $col++) {
            $cal.Table.Cell($row, $col).Shape.TextFrame.TextRange.Text = [string]([char[]]($days[$col - 1])[0])
        }
        $row = 2
        $numDays = [DateTime]::DaysInMonth($y, $mIdx)
        $col = ( Get-Date -Year $y -Month $mIdx -Day 1 ).DayOfWeek.value__
        for ($i = 1; $i -le $numDays; $i++) {
            if ($col -eq 7) {
                $col = 0
                $row++
            }
            # Add holidays to queue
            while (($y -eq [int]$hols[$holIdx].Y) -and ($mIdx -eq [int]$hols[$holIdx].M) -and ($i -eq [int]$hols[$holIdx].D)) {
                $currHols.Add(($hols[$holIdx].Name, $hols[$holIdx].Days, 1, $hols[$holIdx].Repeat))
                $holIdx++
            }
            $cal.Table.Cell($row, $col + 1).Shape.TextFrame.TextRange.Text = [string] $i
            # Mark holidays
            foreach ($hol in $currHols) {
                # skip Sundays for Lent
                if (($col -eq 0) -and ($hol[0] -eq "Lent")) {
                    $hol[2] = [int]$hol[2] + 1
                }
                if ($hol[1] -ge 1 -and $hol[2] -eq 1) {
                    ## Mark holiday
                    $cell = $cal.Table.Cell($row, $col + 1)
                    fill-miniCal $cell $color $brightness $font
                    $hol[1] = [int]$hol[1] - 1
                    if ($hol[1] -eq 0) {
                        $del.Add($hol)
                    }
                    else { $hol[2] = $hol[3] } 
                }
                else { 
                    $hol[2] = $hol[2] - 1
                }
            }
            foreach ($de in $del) {
                [void] $currHols.Remove($de)
            }
            $del.Clear()
            $col++
        }
    }
    $currHols.Clear()

    ## Insert divider labels
    if ($div) {
        Add-divs $slide $title $labels
    }
}