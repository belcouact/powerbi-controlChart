# Create control chart icon with transparent background
Add-Type -AssemblyName System.Drawing

# Create a bitmap with transparent background
$bitmap = New-Object System.Drawing.Bitmap(80, 80)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

# Chart area background (light gray)
$chartBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(245, 245, 245))
$graphics.FillRectangle($chartBrush, 8, 8, 64, 56)

# Grid lines
$gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200), 0.5)
$graphics.DrawLine($gridPen, 8, 16, 72, 16)
$graphics.DrawLine($gridPen, 8, 26, 72, 26)
$graphics.DrawLine($gridPen, 8, 36, 72, 36)
$graphics.DrawLine($gridPen, 8, 46, 72, 46)

# UCL (Upper Control Limit) - Red dashed
$uclPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, 53, 69), 1.5)
$uclPen.DashPattern = [float[]](3, 2)
$graphics.DrawLine($uclPen, 8, 14, 72, 14)

# LCL (Lower Control Limit) - Red dashed
$lclPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, 53, 69), 1.5)
$lclPen.DashPattern = [float[]](3, 2)
$graphics.DrawLine($lclPen, 8, 48, 72, 48)

# Center Line (CL) - Blue solid
$clPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(13, 110, 253), 1.5)
$graphics.DrawLine($clPen, 8, 31, 72, 31)

# Data line
$linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(13, 110, 253), 2)
$linePen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$points = [System.Drawing.Point[]](
    [System.Drawing.Point]::new(14, 35),
    [System.Drawing.Point]::new(22, 27),
    [System.Drawing.Point]::new(30, 33),
    [System.Drawing.Point]::new(38, 25),
    [System.Drawing.Point]::new(46, 29),
    [System.Drawing.Point]::new(54, 23),
    [System.Drawing.Point]::new(62, 31),
    [System.Drawing.Point]::new(68, 27)
)
$graphics.DrawLines($linePen, $points)

# Data points (blue = in control, red = out of control)
$blueBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(13, 110, 253))
$redBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 53, 69))

$graphics.FillEllipse($blueBrush, 10, 31, 8, 8)
$graphics.FillEllipse($blueBrush, 18, 23, 8, 8)
$graphics.FillEllipse($blueBrush, 26, 29, 8, 8)
$graphics.FillEllipse($redBrush, 34, 21, 8, 8)
$graphics.FillEllipse($blueBrush, 42, 25, 8, 8)
$graphics.FillEllipse($redBrush, 50, 19, 8, 8)
$graphics.FillEllipse($blueBrush, 58, 27, 8, 8)
$graphics.FillEllipse($blueBrush, 64, 23, 8, 8)

# Save with transparent background
$bitmap.Save("d:\Alex\PowerBI\controlChart\assets\icon.png", [System.Drawing.Imaging.ImageFormat]::Png)

# Cleanup
$bitmap.Dispose()
$graphics.Dispose()

Write-Host "Icon created successfully!"
