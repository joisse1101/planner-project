param(
    [int]$year = $(throw "-year is required."),
    [int]$mth = $(throw "-mth is required."),
    [int]$day = $(throw "-day is required."),
    [int]$sPos = $(throw "-sPos = slide position")
)

. ".\gen-functions.ps1"
. ".\style-functions.ps1"

[int]$d, [int]$mIdx, [int]$y = Get-fDay $day $mth $year
$day_idx = ( Get-Date -Year $y -Month $mIdx -Day $d ).DayOfWeek.value__
$num = 0
$labels = $divs[0], $divs[1], $divs[2], $divs[4]

while ($y -le $year) {
    $m = $months[$mIdx - 1]
    if (($day_idx % 2) -eq 1) { $layNum = 6 }
    else { $layNum = 5 }
    $slide = $presentation.Slides.AddSlide($sPos + $num, $layouts.Item($layNum))
    $slide.Name = "day_" + $d + $m

    $title = $slide.Shapes(1)
    $title.TextFrame.TextRange.Text = [string]$d + $m
    $title.Name = $slide.Name

    for ($i = 5; $i -le $layNum; $i++) {
        if ($i -eq 6) { [void] $title.TextFrame.TextRange.InsertAfter(" - " + [string]$d + $m) }
        ## Insert date
        $nTop = 1.35 * $cm
        $nLeft = $cm + ($i - 5) * 9 * $cm
        $box = $slide.Shapes.AddTextbox(1, $nLeft, $nTop, 1.5 * $cm, 0.5 * $cm)
        $box.TextFrame.AutoSize = 0
        $box.Height = 0.5 * $cm
        $box.TextFrame.TextRange.Font.Size = 9
        $box.TextFrame.TextRange.Font.Bold = 1
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
        $box.TextFrame.VerticalAnchor = 1 # align top
        $box.TextFrame.TextRange.Text = [string]$d + " " + $m.Substring(0, 3)
        $box.Name = "date_" + [string]$d + "_" + $m

        ## Insert day
        $box = $slide.Shapes.AddTextbox(1, $nLeft + $box.Width, $nTop, 2 * $cm, 0.5 * $cm)
        $box.TextFrame.AutoSize = 0
        $box.Height = 0.5 * $cm
        $box.TextFrame.MarginLeft = 0
        $box.TextFrame.TextRange.Font.Size = 9
        $box.TextFrame.TextRange.Font.Bold = 1
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
        $box.TextFrame.VerticalAnchor = 1 # align top
        $box.TextFrame.TextRange.Text = $days[$day_idx]
        $box.Name = "day_" + [string]$d + "_" + $m

        ## Insert type
        $box = $slide.Shapes.AddTextbox(1, $box.Left + $box.Width, $nTop, 5.5 * $cm, 0.75 * $cm)
        $box.TextFrame.AutoSize = 0
        $box.Height = 0.75 * $cm
        $box.TextFrame.MarginLeft = 0
        $box.TextFrame.TextRange.Font.Name = $title.TextFrame.TextRange.Font.Name
        $box.TextFrame.TextRange.Font.Size = 10
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # align right
        $box.TextFrame.VerticalAnchor = 1 # align top 
        if ($day_idx -eq 0) { $box.TextFrame.TextRange.Text = $dType[0] }
        else { $box.TextFrame.TextRange.Text = $dType[1] }
        $box.Name = "type_" + [string]$d + "_" + $m

        ## Insert holiday
        # Add holidays to queue
        $holIdx = Get-holIdx $d $mIdx $y $hols
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
                $box = $slide.Shapes.AddTextbox(1, $nLeft, $hTop, 9 * $cm, 0.5 * $cm)
                use-BodyText.$box
                $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
                $box.TextFrame.VerticalAnchor = 3 # align middle
                $box.Fill.BackColor.ObjectThemeColor = 7
                $box.Fill.BackColor.Brightness = 0.6
                $box.Line.Visible = 1
                $box.Line.Weight = 1
                $box.Line.ForeColor.ObjectThemeColor = 7
                $box.Line.ForeColor.Brightness = 0.6
                $box.ZOrder(1)
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

        ## Increment Y-M-D
        $d, $mIdx, $y = AddDMY $d $mIdx $y 1
        $day_idx = ($day_idx + 1) % 7
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

    $num++
}
