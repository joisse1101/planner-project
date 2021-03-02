param(
    [int]$year = $(throw "-year is required."),
    [int]$mth = $(throw "-mth is required."),
    [int]$sPos = $(throw "-sPos = slide position")
)

. ".\gen-functions.ps1"
. ".\style-functions.ps1"

$labels = $divs[0], $divs[2], $divs[3], $divs[4]
$miniCal = $templates.month
$y = $year

$holIdx, $currHols = Get-holIdx 1 $mth $year $hols

for ($mIdx = $mth; $mIdx -le $months.Count; $mIdx++) {
    $m = $months[$mIdx - 1]
    ## Set layout number
    $numDays = [DateTime]::DaysInMonth($year, $mIdx)
    $day_idx = ( Get-Date -Year $year -Month $mIdx -Day 1 ).DayOfWeek.value__
    $cal = $day_idx + $numDays
    if ($cal -eq 36) { $layNum = 2 }
    ElseIf ($cal -eq 37) { $laynum = 3 }
    Else { $layNum = 1 }

    $slide = $presentation.Slides.AddSlide(($sPos + $mIdx - $mth), $layouts.Item($layNum))
    $slide.Name = "mth_" + $m
    $title = $slide.Shapes(1)
    $title.TextFrame.TextRange.Text = $m
    $title.Name = $slide.Name

    ## Insert day titles
    $nTop = $title.Top + $title.Height + $cm / 2
    $nLeft = $title.Left + 0.25 * $cm
    $lefts = New-Object System.Collections.Generic.List[int]
    foreach ($d in $days) {
        $lefts.Add($nLeft)
        $box = $slide.Shapes.AddTextbox(1, $nLeft, $nTop, 2.5 * $cm, 0.5 * $cm)
        use-mthDayLabel $box
        $box.TextFrame.TextRange.Text = $d
        $box.Name = "name_" + $d
    
        $nLeft = $nLeft + $box.Width
    }
    ## Insert calender numbers
    $nTop = $nTop + 0.5 * $cm
    for ($d = 1; $d -le $numDays; $d++) {
        if ($day_idx -eq 7) {
            $day_idx = 0
            $nTop = $nTop + 3 * $cm
        }

        $box = $slide.Shapes.AddTextbox(1, $lefts[$day_idx], $nTop, 2.5 * $cm, 3 * $cm)
        use-mthDateLabel $box
        $box.TextFrame.TextRange.Text = [string]$d
        $box.Name = "day_" + $d

        # Insert holiday
        # Add holidays to queue
        while (($y -eq [int]$hols[$holIdx].Y) -and ($mIdx -eq [int]$hols[$holIdx].M) -and ($d -eq [int]$hols[$holIdx].D)) {
            $currHols.Add(($hols[$holIdx].Name, $hols[$holIdx].Days, 1, $hols[$holIdx].Days, $hols[$holIdx].Repeat))
            $holIdx++
        }
        # Mark holidays
        $hTop = $nTop + 0.5 * $cm
        foreach ($hol in $currHols) {
            # skip Sundays for Lent
            if (($day_idx -eq 0) -and ($hol[0] -eq "Lent")) {
                $hol[2] = [int]$hol[2] + 1
            }
            if ($hol[1] -ge 1 -and $hol[2] -eq 1) {
                ## Mark holiday
                $box = $slide.Shapes.AddTextbox(1, $lefts[$day_idx], $hTop , 2.5 * $cm, 0.5 * $cm)
                $box.TextFrame.AutoSize = 0                
                use-mthHolLabel $box                
                $box.TextFrame.TextRange.Text = $hol[0].ToUpper()
                if ($hol[3] -gt 1) { [void] $box.TextFrame.TextRange.InsertAfter( " " + "D-" + [string]([int]$hol[3] - [int]$hol[1] + 1)) }
                $box.Name = "hol_" + $hol[0].ToUpper()
                $box.ZOrder(1)

                $box.TextFrame.AutoSize = 1
                $box.Left = [double]$lefts[$day_idx]
                $box.Top = $hTop
                $hol[1] = [int]$hol[1] - 1

                $hTop = $hTop + $box.Height + 0.1 * $cm
                if ($hol[1] -eq 0) {
                    $del.Add($hol)
                }
                else { $hol[2] = $hol[4] } 
            }
            else { 
                $hol[2] = $hol[2] - 1
            }
        }
        foreach ($de in $del) {
            [void] $currHols.Remove($de)
        }
        $del.Clear()
        $day_idx++
    }
    
    ## Insert calender minis
    $nTop = $nTop + 4 * $cm
    $nLeft = $title.Left
    $div = ($title.Width - 5.6 * $cm * 3) / 2
    $cals = 3
    if ($layNum -gt 1 ) {
        $cals = 2
        $nTop = $nTop - 3 * $cm
        $nLeft = $nLeft + 5.6 * $cm + $div
    }

    $y = $year        
    for ($n = ($mIdx + 1); $n -le ($mIdx + $cals); $n++) {
        if ($n -eq 13) { $y++ }
        $miniCal.Shapes("mini_" + $months[($n - 1) % 12] + "_" + $y).Copy()
        $box = $slide.Shapes.Paste()
        $box.Top = $nTop
        [double]$box.Left = [double]$nLeft

        $miniCal.Shapes("mini_" + $months[($n - 1) % 12] + "_" + $y + "_cal").Copy()
        $cal = $slide.Shapes.Paste()
        $cal.Top = $nTop + $box.Height
        [double]$cal.Left = [double]$nLeft

        $nLeft = $nLeft + $box.Width + $div
    }
    
    ## Insert divider labels 
    $nTop = 0.35 * $cm
    $nLeft = 20 * $cm
    foreach ($l in $labels) {
        $box = $slide.Shapes.AddTextbox(5, $nLeft, $nTop, 0.75 * $cm, 7.25 * $cm) # vertical textbox
        use-DivText $box $title
        $box.TextFrame.TextRange.Text = $l
        $box.Name = "div_" + $l
        $nTop = $nTop + $box.Height
    }
}