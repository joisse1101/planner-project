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
    $cell.Shape.Fill.Solid()
    $cell.Shape.Fill.ForeColor.ObjectThemeColor = [int]$color
    $cell.Shape.Fill.ForeColor.Brightness = [single]$brightness
    $cell.Shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = [int]$font
}

function use-mthDayLabel {
    param ($box)
    use-BodyText $box
    $box.TextFrame.AutoSize = 0
    $box.Height = 0.48 * $cm
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.Fill.Solid()
    $box.Fill.ForeColor.ObjectThemeColor = 5
    $box.Fill.ForeColor.Brightness = 0.6
    $box.Line.Visible = 1
    $box.Line.Weight = 0.75
    $box.Line.ForeColor.ObjectThemeColor = 2
}

function use-mthDateLabel {
    param ($box)
    $box.TextFrame.AutoSize = 0
    $box.TextFrame.TextRange.Font.Size = 9
    $box.TextFrame.TextRange.Font.Bold = 1
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
    $box.TextFrame.VerticalAnchor = 1 # align top
}

function use-mthHolLabel {
    param($box)
    use-mthDayLabel $box
    $box.Height = 0.5 * $cm
    $box.TextFrame.TextRange.Font.Size = 6
    $box.TextFrame.MarginBottom = 1
    $box.TextFrame.MarginTop = 1
    $box.Line.Weight = 1
    $box.Line.ForeColor.ObjectThemeColor = 5
    $box.Line.ForeColor.Brightness = 0.6
}