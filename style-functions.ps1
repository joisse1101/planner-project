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
    $box.Height = 0.475 * $cm
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

function use-wkDayLabel {
    param($box, $title)
    use-titleText $box $title
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
    $box.TextFrame.VerticalAnchor = 3 # align middle
}

function use-wkDateLabel {
    param($box, $title)
    use-titleText $box $title
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # align right
    $box.TextFrame.VerticalAnchor = 3 # align middle
}

function use-wkHolLabel {
    param($box)
    use-BodyText $box
    $box.TextFrame.TextRange.Font.Bold = 0
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.Fill.Solid()
    $box.Fill.ForeColor.ObjectThemeColor = 6
    $box.Fill.ForeColor.Brightness = 0.6
}

function use-wkTrackerLabel {
    param($box)
    use-BodyText $box
    $box.Height = 0.5 * $cm
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 1 # align top
    $box.TextFrame.MarginTop = 0.125 * $cm
    $box.Fill.Solid()
    $box.Fill.ForeColor.ObjectThemeColor = 6
    $box.Fill.ForeColor.Brightness = 0.6
    $box.Line.Visible = 1
    $box.Line.Weight = 1
    $box.Line.ForeColor.ObjectThemeColor = 6
}

function use-dayLabel {
    param ($box)
    use-BodyText $box
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # align left
    $box.TextFrame.VerticalAnchor = 3 # align middle
}

function use-dayTypeLabel {
    param($box, $title)
    use-TitleText $box $title
    $box.Height = 0.75 * $cm
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # align right
    $box.TextFrame.VerticalAnchor = 1 # align top 
}

function use-dayHolLabel {
    param($box)
    use-BodyText $box
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.Fill.Solid()
    $box.Fill.ForeColor.ObjectThemeColor = 7
    $box.Fill.ForeColor.Brightness = 0.8
    $box.Line.Visible = 1
    $box.Line.Weight = 1
    $box.Line.ForeColor.ObjectThemeColor = 7
    $box.Line.ForeColor.Brightness = 0.8
    $box.ZOrder(1)
}