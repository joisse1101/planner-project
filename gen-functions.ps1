## [void] Link $targetSlide $linkSlide
function Link {
    param($target, $link)
    $target.ActionSettings(1).Action = 10
    $target.ActionSettings(1).Hyperlink.SubAddress = [string]$link.SlideID + ", " + [string]$link.SlideIndex + ", " + $link.Name
}

## [[int]Day, [int]Mth, [int]Year] AddDMY $d $m $y $numDays
function AddDMY {
    param([int]$d, [int]$m, [int]$y, [int]$num)
    $YMD = (Get-Date -Year $y -Month $m -Day $d).AddDays($num).ToString("yyyyMMdd")
    $y = [int]$YMD.Substring(0, 4)
    $m = [int]$YMD.Substring(4, 2)
    $d = [int]$YMD.Substring(6, 2)
    return [int]$d, [int]$m, [int]$y
}

## [[int]Day, [int]Mth, [int]Year] Get-fSunday $d $m $y
function Get-fSunday {
    param([int]$d, [int]$m, [int]$y)
    while ( (Get-Date -Year $y -Month $m -Day $d ).DayOfWeek.value__ -ne 0) {
        $d, $m, $y = AddDMY $d $m $y -1
    }
    return [int]$d, [int]$m, [int]$y
}

## [[int]Day, [int]Mth, [int]Year] Get-fDay $d $m $y
function Get-fDay {
    param([int]$d, [int]$m, [int]$y)
    $day_idx = ( Get-Date -Year $y -Month $m -Day $d ).DayOfWeek.value__
    if ((($day_idx % 2) -eq 0) -and ($day_idx -ne 0)) { 
        $d, $m, $y = AddDMY $d $m $y -1
    }
    return [int]$d, [int]$m, [int]$y
}

## [[int]holIdx] Get-holIdx $d $m $y $hols
function Get-holIdx {
    param([int]$d, [int]$m, [int]$y, $hols)
    $holIdx = 0
    while (($y -ge [int]$hols[$holIdx].Y)) { 
        if ($y -gt [int]$hols[$holIdx].Y) {
            $holIdx++
        }
        elseif ($y -eq [int]$hols[$holIdx].Y) {
            if ($m -gt [int]$hols[$holIdx].M) {
                $holIdx++
            }
            elseif ($m -eq [int]$hols[$holIdx].M) {
                if ($d -gt [int]$hols[$holIdx].D) {
                    $holIdx++
                }
                else { break }
            }
            else { break }
        }
        else { break }
    }
    return $holIdx
}