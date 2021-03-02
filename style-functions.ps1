function use-TitleText {
    param($box, $title)
    $box.TextFrame.AutoSize = 0
    $box.Height = 0.5 * $cm
    $box.TextFrame.TextRange.Font.Name = $title.TextFrame.TextRange.Font.Name
    $box.TextFrame.TextRange.Font.Size = 10
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.MarginTop = 0
}
function use-BodyText {
    param($box)
    $box.TextFrame.AutoSize = 0
    $box.Height = 0.5 * $cm
    $box.TextFrame.TextRange.Font.Size = 9
    $box.TextFrame.TextRange.Font.Bold = 1
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.MarginTop = 0
}

function use-DivText {
    param($box, $title)
    $box.TextFrame.AutoSize = 0
    $box.TextFrame.TextRange.Font.Name = $title.TextFrame.TextRange.Font.Name
    $box.TextFrame.TextRange.Font.Size = 10        
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.TextFrame.TextRange.Font.Color.ObjectThemeColor = 2



}

## [void] use-miniCalTitle $box $m $y        
function use-miniCalTitle {
    param($box, $m, $y)
    use-BodyText $box
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.TextFrame.TextRange.Text = [string]$m.ToUpper()
    $box.Name = "mini_" + $m + "_" + $y
}

## [void] use-miniCal $box $m $y $cal
function use-miniCal {
    param($box, $m, $y, $cal)
    $cal.Name = "mini_" + $m + "_" + $y + "_cal"
    $cal.Left = $box.Left
    $cal.Top = $box.Top + $box.Height
}

## [void] fill-miniCal $cell $color $brightness $font
function fill-miniCal {
    param ($cell, $color, $brightness, $font)
    $cell.Shape.Fill.BackColor.ObjectThemeColor = [int]$color
    $cell.Shape.Fill.BackColor.Brightness = [single]$brightness
    $cell.Shape.Fill.Solid()
    $cell.Shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = [int]$font
}