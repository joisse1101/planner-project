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

function label-miniCalTitle {
    param($box, $m, $y)
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # align center
    $box.TextFrame.VerticalAnchor = 3 # align middle
    $box.TextFrame.TextRange.Text = [string]$m.ToUpper()
    $box.Name = "mini_" + $m + "_" + $y
}

function label-miniCal {
    param($box, $m, $y, $cal)
    $cal.Name = "mini_" + $m + "_" + $y + "_cal"
    $cal.Left = $box.Left
    $cal.Top = $box.Top + $box.Height
}