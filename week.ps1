param(
    [int]$year = $(throw "-year is required."),
    [int]$mth = $(throw "-mth is required."),
    [int]$day = $(throw "-day is required."),
    [int]$sPos = $(throw "-sPos = slide position")
)

. ".\gen-functions.ps1"
. ".\style-functions.ps1"

$d, $mIdx, $y = Get-fSunday $day $mth $year
$miniCal = $templates.week
$labels = $divs[0], $divs[1], $divs[3], $divs[4]

$num = 0
while ($y -le $year) {
    $numDays = [DateTime]::DaysInMonth($y, $mIdx)
    $m = $months[$mIdx-1]

    $slide = $presentation.Slides.AddSlide($sPos + $num, $layouts.Item(4))
    $slide.Name = [string]("wk_" + $d + $m)

    $title = $slide.Shapes(1)
    $title.TextFrame.TextRange.Text = [string]$d + " " + $m
    $title.Name = $slide.Name
    
    ## Insert mini calendar
    if (($d + 6) -gt $numDays) { 
        $mCal, $yCal = ($mIdx + 1), $y
        if ($mCal -eq 13) { $mCal, $yCal = 1, ($y + 1)}
    }
    else { $mCal, $yCal = $mIdx, $y }

    $nTop = $title.Top
    $nLeft = $title.Left + $title.Width + 0.05 * $cm
    $miniCal.Shapes("mini_" + $months[$mCal - 1] + "_" + $yCal).Copy()
    $box = $slide.Shapes.Paste()
    [double]$box.Top = [double]$nTop
    [double]$box.Left = [double]$nLeft

    $miniCal.Shapes("mini_" + $months[$mCal - 1] + "_" + $yCal + "_cal").Copy()
    $cal = $slide.Shapes.Paste()
    $cal.Top = $nTop + $box.Height
    [double]$cal.Left = [double]$nLeft

    for ($i = 0; $i -lt $days.Count; $i++) {
        $m = $months[$mIdx-1]
        $nTop = $title.Top + 6.75 * $cm + ($i % 4) * 5 * $cm
        $nLeft = $title.Left + [math]::Floor($i / 4) * 9.25 * $cm

        ## Insert day names
        $box = $slide.Shapes.AddTextbox(1, $nLeft, $nTop, 4 * $cm, 0.5 * $cm)
        use-titleText $box $title
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
        $box.TextFrame.VerticalAnchor = 3 # align middle
        $box.TextFrame.TextRange.Text = [string]$days[$i]
        $box.Name = "day_" + [string]$days[$i]

        ## Insert dates
        $box = $slide.Shapes.AddTextbox(1, $nLeft + 6.25 * $cm, $nTop, 2.5 * $cm, 0.5 * $cm)
        use-titleText $box $title
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
        $box.TextFrame.VerticalAnchor = 3 # align middle
        $box.TextFrame.TextRange.Text = [string]$d + $m.Substring(0,3)
        $box.Name = "date_" + [string]$d + $m

        ## Insert holiday
        # Add holidays to queue
        $holIdx = Get-holIdx $d $mIdx $y $hols
        while (($y -eq [int]$hols[$holIdx].Y) -and ($mIdx -eq [int]$hols[$holIdx].M) -and ($d -eq [int]$hols[$holIdx].D)) {
            $currHols.Add(($hols[$holIdx].Name, $hols[$holIdx].Days, 1, $hols[$holIdx].Days, $hols[$holIdx].Repeat))
            $holIdx++
        }
        # Mark holidays
        $hTop = $nTop + 0.25 * $cm
        foreach ($hol in $currHols) {
            # skip Sundays for Lent
            if (($day_idx -eq 0) -and ($hol[0] -eq "Lent")) {
                $hol[2] = [int]$hol[2] + 1
            }
            if ($hol[1] -ge 1 -and $hol[2] -eq 1) {
                ## Mark holiday
                $box = $slide.Shapes.AddTextbox(1, $nLeft, $hTop, 8.75 * $cm, 0.5 * $cm)
                use-BodyText $box
                $box.TextFrame.TextRange.Font.Bold = 0
                $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
                $box.TextFrame.VerticalAnchor = 3 # align middle
                $box.Fill.BackColor.ObjectThemeColor = 6
                $box.Fill.BackColor.Brightness = 0.6
                $box.TextFrame.TextRange.Text = $hol[0].ToUpper()
                $box.Name = "hol_" + $hol[0].ToUpper()
                if ($hol[3] -gt 1) { [void] $box.TextFrame.TextRange.InsertAfter( " " + "D-" + [string]([int]$hol[3] - [int]$hol[1] + 1)) }
                $box.Name = "hol_" + $hol[0].ToUpper()
                $box.ZOrder(1)
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

        ## Append end-date to title
        if ($i -eq 6) { [void] $title.TextFrame.TextRange.InsertAfter(" - " + $d + " " + $m) }
        
        ## Increment Y-M-D
        $d, $mIdx, $y = AddDMY $d $mIdx $y 1
    }

    ## Insert habit tracker
    $nTop = $title.Top + 21.5 * $cm
    $nLeft = $title.Left + 9.25 * $cm

    $box = $slide.Shapes.AddTextbox(1, $nLeft, $nTop, 8.75 * $cm, 0.75 * $cm)
    use-BodyText $box
    $box.Height = 0.75 * $cm
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.Fill.BackColor.ObjectThemeColor = 6
    $box.Fill.BackColor.Brightness = 0.6
    $box.Line.Visible = 1
    $box.Line.Weight = 1
    $box.Line.ForeColor.ObjectThemeColor = 6

    $box.TextFrame.TextRange.Text = "Habit Tracker"
    $box.Name = "habit_title"

    $templates.templates.Shapes("habit-week").Copy()
    [void] $slide.Shapes.Paste()
    $tracker = $slide.Shapes($slide.Shapes.Count)
    $tracker.Name = "habit_chart"
    $tracker.Top = $nTop + $box.Height
    [Double] $tracker.Left = $nLeft

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
    ## Increment slide position
    $num++
}