#
# Productivity Excelerator
# Guarenteed 10x productivity improvement in every serving.
#
# By Jimurrito 
#
# REAL Purpose:
# 
# Opens a Winform and Excel
# Every 60s, a new row is added to Excel
# Form is used to stop the loop
# This keeps your machine awake and status active on "work apps".
#

#
# Imports the dotnet/WinAPI Forms lib
Add-Type -AssemblyName System.Windows.Forms

#
# Creates form object itself
$form = New-Object System.Windows.Forms.Form
$form.Text = "(Excel)erator - [K'etzal Cycle Engine]"
$form.Width = 380
$form.Height = 120
$form.TopMost = $true
$form.FormBorderStyle = 'FixedDialog'
$form.StartPosition = "CenterScreen"

#
# Creates form button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Click 2 Stop Cadence Oscillation"
$button.Dock = "Fill"

#
# Button color behavior
$defaultBackColor = $button.BackColor
$hoverBackColor = [System.Drawing.Color]::FromArgb(255, 230, 230)
$clickBackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
# Hover enter
$button.Add_MouseEnter({
        $button.BackColor = $hoverBackColor
    })
# Hover leave
$button.Add_MouseLeave({
        $button.BackColor = $defaultBackColor
    })
# Click visual feedback
$button.Add_MouseDown({
        $button.BackColor = $clickBackColor
    })
$button.Add_MouseUp({
        $button.BackColor = $hoverBackColor
    })

#
# Adds form button to said form
$form.Controls.Add($button)

#
# Initialize Excel handler
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true # visible true so it keeps the machine awake

#
# Create workbook and sheet handles
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

#
# Productivity Inspiration 
$phrases = @(
    "Cycle stable"
    "Resonance holding"
    "Alignment nominal"
    "Intent lattice active"
    "Cadence oscillating"
    "Manifold engaged"
    "Glyph sequence valid"
    "Focus vector rising"
    "Cognitive field steady"
    "Tri-phasic loop intact"
    "Harmonic drift minimal"
    "Continuum acknowledged"
    "Ixatli signal detected"
    "K'etzal engine humming"
    "Temporal pocket forming"
    "Micro-oscillation complete"
    "Phase corridor clear"
    "Attention anchor set"
    "Motivation amplitude high"
    "Perception manifold open"
    "Stability threshold met"
    "Recursive pulse received"
    "Operational clarity achieved"
    "Intentional sweep passed"
    "Cycle integrity confirmed"
    "The cake is a lie"
)

#
# Defines the random object for inspiration selection
$rand = New-Object System.Random
# Sets script env var row pointer to 1 (excel is base-1)
$script:row = 1
# Default interval
$intervalSeconds = 60

#
# Create form timer that allows the excel job to run in the background 
$timer = New-Object System.Windows.Forms.Timer
# Interval is in milliseconds
$timer.Interval = $intervalSeconds * 1000
# When the interval is hit, this tick function will trigger via the timer
$timer.Add_Tick({
        $sheet.Cells.Item($script:row, 1).Value2 = (Get-Date).ToString("o")
        $sheet.Cells.Item($script:row, 2).Value2 = $phrases[$rand.Next($phrases.Count)]
        $sheet.Columns.AutoFit()
        $script:row++
    })
# Starts timer background job
$timer.Start()

#
# Stop button closes form
# Declared here so we can include the $timer reference
$button.Add_Click({
        $timer.Stop()
        $form.Close()
    })

#
# Blocking function that keeps everything going until the "Stop button" is pressed
$form.ShowDialog() | Out-Null

#
# Once you are here, we are running shutdown/cleanup
# 

#
# Closes workbook and excel handles
$workbook.Close($false)
$excel.Quit()

#
# manually releases the objects from the script
# this allows dotnet to clean them up via Garbage collection (your mom)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

#
# trigger GC manually
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
