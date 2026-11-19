# SureCover Tester
# Generates a fictional flight-delay claim (7 files with the names SureCover expects)
# and opens a ready-to-send email in classic Outlook.
# All generated documents are labelled as fictional demonstration data.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Footer = 'Fictional demonstration data. Not a SureCover policy document.'
$DefaultRecipient = 'claims@eeelkoizil.resend.app'
$OutRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'SureCover tests'

$AllFiles = @(
    'claim-form.pdf',
    'boarding-pass-sc101.pdf',
    'boarding-pass-sc101.png',
    'boarding-pass-sc202.pdf',
    'delay-letter.pdf',
    'scheduled-itinerary.pdf',
    'revised-itinerary.pdf'
)

# ---------------------------------------------------------------- PDF writer
function Esc([string]$s) { $s.Replace('\', '\\').Replace('(', '\(').Replace(')', '\)') }

function Txt($x, $y, $font, $size, $text, $color = '0.07 g') {
    "BT $color /$font $size Tf $x $y Td ($(Esc $text)) Tj ET"
}

# Items: @{K='Label';V='Value'}  @{P='paragraph line'}  @{Gap=$true}
#        @{Table=@(@('h1','h2',...), @('r1c1',...), ...)}
function Write-ClaimPdf([string]$Path, [string]$Title, $Items) {
    $ops = New-Object System.Collections.Generic.List[string]
    $ops.Add('0.075 0.161 0.294 rg 0 752 595.28 90 re f')
    $ops.Add('0.78 0.63 0.31 rg 0 748 595.28 4 re f')
    $ops.Add((Txt 50 784 'F2' 22 $Title '1 1 1 rg'))
    $ops.Add((Txt 50 728 'F3' 9 $Footer '0.33 g'))
    $y = 690
    foreach ($it in $Items) {
        if ($it.ContainsKey('K')) {
            $ops.Add((Txt 50 $y 'F2' 11 ($it.K + ':')))
            $ops.Add((Txt 200 $y 'F1' 11 $it.V))
            $y -= 22
        } elseif ($it.ContainsKey('P')) {
            $ops.Add((Txt 50 $y 'F1' 11 $it.P))
            $y -= 18
        } elseif ($it.ContainsKey('Gap')) {
            $y -= 14
        } elseif ($it.ContainsKey('Table')) {
            $cols = @(50, 140, 290, 370, 460)
            $first = $true
            foreach ($row in $it.Table) {
                $font = 'F1'; if ($first) { $font = 'F2' }
                for ($i = 0; $i -lt $row.Count; $i++) {
                    $ops.Add((Txt $cols[$i] $y $font 11 $row[$i]))
                }
                if ($first) { $ops.Add("0.6 G 0.5 w 50 $($y - 7) m 545 $($y - 7) l S") ; $y -= 26 } else { $y -= 20 }
                $first = $false
            }
        }
    }
    $stream = ($ops -join "`n")

    $objs = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Contents 4 0 R /Resources << /Font << /F1 5 0 R /F2 6 0 R /F3 7 0 R >> >> >>',
        "<< /Length $($stream.Length) >>`nstream`n$stream`nendstream",
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Oblique /Encoding /WinAnsiEncoding >>',
        "<< /Title ($(Esc $Title)) /Producer (SureCover Tester) >>"
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n")
    $offsets = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offsets += $sb.Length
        [void]$sb.Append("$($i + 1) 0 obj`n$($objs[$i])`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offsets) { [void]$sb.Append(("{0:D10} 00000 n `n" -f $o)) }
    [void]$sb.Append("trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R /Info 8 0 R >>`nstartxref`n$xref`n%%EOF`n")
    [System.IO.File]::WriteAllBytes($Path, [System.Text.Encoding]::GetEncoding(1252).GetBytes($sb.ToString()))
}

# ---------------------------------------------------------------- PNG writer
function Write-BoardingPassPng([string]$Path, [string]$Name, [string]$Flight, [datetime]$Date, [string]$From, [string]$To) {
    $bmp = New-Object System.Drawing.Bitmap 840, 1080
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $wood = [System.Drawing.Color]::FromArgb(160, 112, 70)
    $card = [System.Drawing.Color]::FromArgb(246, 234, 210)
    $navy = [System.Drawing.Color]::FromArgb(19, 41, 75)
    $gold = [System.Drawing.Color]::FromArgb(198, 160, 80)
    $ink = [System.Drawing.Color]::FromArgb(25, 25, 25)
    $grey = [System.Drawing.Color]::FromArgb(85, 85, 85)
    $g.Clear($wood)
    for ($y = 0; $y -lt 1080; $y += 120) {
        $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 96, 60))), 0, $y, 840, 3)
    }
    $g.TranslateTransform(420, 540); $g.RotateTransform(-1.5); $g.TranslateTransform(-420, -540)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(90, 0, 0, 0))), 70, 62, 720, 980)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $card), 56, 48, 720, 980)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $navy), 56, 48, 720, 150)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $gold), 56, 198, 720, 8)
    $white = New-Object System.Drawing.SolidBrush $card
    $g.DrawString('SC AIR', (New-Object System.Drawing.Font('Segoe UI', 34, [System.Drawing.FontStyle]::Bold)), $white, 96, 70)
    $g.DrawString('BOARDING PASS', (New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)), $white, 98, 135)
    $lab = New-Object System.Drawing.Font('Segoe UI', 13)
    $val = New-Object System.Drawing.Font('Segoe UI', 30, [System.Drawing.FontStyle]::Bold)
    $gb = New-Object System.Drawing.SolidBrush $grey; $ib = New-Object System.Drawing.SolidBrush $ink
    $fields = @(
        @(96, 250, 'PASSENGER', $Name),
        @(96, 400, 'FLIGHT', $Flight), @(440, 400, 'DATE', $Date.ToString('d MMM yyyy', $Inv)),
        @(96, 550, 'FROM', $From), @(440, 550, 'TO', $To)
    )
    foreach ($f in $fields) {
        $g.DrawString($f[2], $lab, $gb, $f[0], $f[1])
        $g.DrawString($f[3], $val, $ib, $f[0] - 4, $f[1] + 26)
    }
    $rnd = New-Object System.Random ($Name.GetHashCode())
    $x = 100
    while ($x -lt 730) { $w = $rnd.Next(1, 6); $g.FillRectangle($ib, $x, 740, $w, 130); $x += $w + $rnd.Next(2, 6) }
    $g.DrawString($Footer, (New-Object System.Drawing.Font('Segoe UI', 11)), $gb, 96, 900)
    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

# ---------------------------------------------------------------- claim data
function Get-Claim {
    $c = @{
        Name     = $tbName.Text.Trim()
        Policy   = $tbPolicy.Text.Trim()
        From     = $tbFrom.Text.Trim().ToUpper()
        To       = $tbTo.Text.Trim().ToUpper()
        ReturnPlace = $tbReturnPlace.Text.Trim()
        Out      = $dpOut.Value.Date
        Back     = $dpBack.Value.Date
        DocDelay = [int]$nuDelay.Value
        MailDelay= [int]$nuMailDelay.Value
        Cause    = $tbCause.Text.Trim()
    }
    if ([string]::IsNullOrWhiteSpace($c.ReturnPlace)) { $c.ReturnPlace = 'Singapore' }
    $t = [datetime]::ParseExact($tbDep.Text.Trim(), 'HH:mm', $Inv)
    $c.SchedDep = $c.Back.Add($t.TimeOfDay)
    $c.RevDep = $c.SchedDep.AddHours($c.DocDelay)
    return $c
}

function Build-Files($c, [string]$Folder, $Include) {
    $d = { param($dt) $dt.ToString('d MMMM yyyy', $Inv) }
    $s = { param($dt) $dt.ToString('d MMM yyyy', $Inv) }
    $flights = @(@('Flight', 'Date', 'From', 'To'),
                 @('SC101', (& $d $c.Out), $c.From, $c.To),
                 @('SC202', (& $d $c.Back), $c.To, $c.From))
    if ($Include -contains 'claim-form.pdf') {
        Write-ClaimPdf (Join-Path $Folder 'claim-form.pdf') 'Flight-delay claim form' @(
            @{K = 'Claimant'; V = $c.Name }, @{K = 'Policy reference'; V = $c.Policy },
            @{K = 'Claim type'; V = 'Flight delay' },
            @{K = ('Return to ' + $c.ReturnPlace); V = $c.Back.ToString('dddd d MMMM yyyy', $Inv) },
            @{Gap = $true }, @{P = 'Affected flights' }, @{Gap = $true }, @{Table = $flights })
    }
    foreach ($bp in @(@('sc101', 'SC101', $c.Out, $c.From, $c.To), @('sc202', 'SC202', $c.Back, $c.To, $c.From))) {
        if ($Include -contains "boarding-pass-$($bp[0]).pdf") {
            Write-ClaimPdf (Join-Path $Folder "boarding-pass-$($bp[0]).pdf") 'Boarding pass' @(
                @{K = 'Passenger'; V = $c.Name }, @{K = 'Flight'; V = $bp[1] }, @{K = 'Date'; V = (& $d $bp[2]) },
                @{K = 'From'; V = $bp[3] }, @{K = 'To'; V = $bp[4] })
        }
    }
    if ($Include -contains 'boarding-pass-sc101.png') {
        Write-BoardingPassPng (Join-Path $Folder 'boarding-pass-sc101.png') $c.Name 'SC101' $c.Out $c.From $c.To
    }
    if ($Include -contains 'delay-letter.pdf') {
        Write-ClaimPdf (Join-Path $Folder 'delay-letter.pdf') 'Airline delay notice' @(
            @{K = 'Passenger'; V = $c.Name }, @{Gap = $true },
            @{P = "Flight SC202 from $($c.To) to $($c.From) on $(& $d $c.Back) was delayed" },
            @{P = "because of $($c.Cause)." }, @{Gap = $true },
            @{K = 'Delay duration'; V = "$($c.DocDelay) hours" })
    }
    $hdr = @('Flight', 'Date', 'Departs', 'Route')
    $leg1 = @('SC101', (& $s $c.Out), '09:10', "$($c.From) - $($c.To)")
    if ($Include -contains 'scheduled-itinerary.pdf') {
        Write-ClaimPdf (Join-Path $Folder 'scheduled-itinerary.pdf') 'Scheduled itinerary' @(
            @{K = 'Passenger'; V = $c.Name }, @{Gap = $true },
            @{Table = @($hdr, $leg1, @('SC202', (& $s $c.SchedDep), $c.SchedDep.ToString('HH:mm', $Inv), "$($c.To) - $($c.From)")) })
    }
    if ($Include -contains 'revised-itinerary.pdf') {
        Write-ClaimPdf (Join-Path $Folder 'revised-itinerary.pdf') 'Revised itinerary' @(
            @{K = 'Passenger'; V = $c.Name }, @{Gap = $true },
            @{Table = @($hdr, $leg1, @('SC202', (& $s $c.RevDep), $c.RevDep.ToString('HH:mm', $Inv), "$($c.To) - $($c.From)")) })
    }
}

function Update-EmailText {
    try { $c = Get-Claim } catch { return }
    $tbSubject.Text = "Flight delay claim - $($c.Name)"
    $tbBody.Text = ("Dear SureCover team,`r`n`r`n" +
        "My return flight SC202 from $($c.To) to $($c.From) on $($c.Back.ToString('d MMMM', $Inv)) was delayed by $($c.MailDelay) hours because of $($c.Cause). " +
        "I'd like to claim under my travel policy $($c.Policy).`r`n`r`n" +
        "Please find all the required documents attached to this email.`r`n`r`n" +
        "Kind regards,`r`n$($c.Name)")
}

# ---------------------------------------------------------------- random person / presets
$First = 'Priya','Daniel','Mei Ling','Arjun','Sofia','Marcus','Aisha','Kenji','Nora','Rafael','Hana','Oliver','Siti','Lucas','Farah','Ethan'
$Last  = 'Raman','Tan','Okafor','Lim','Novak','Hassan','Sato','Fernandes','Koh','Petrov','Wong','Mendes','Ibrahim','Nguyen','Schmidt','Rahman'
$Causes = 'a typhoon warning at the destination airport','a late inbound aircraft','an unscheduled technical inspection','air traffic control restrictions','a crew scheduling problem'
$Rng = New-Object System.Random

function Set-RandomPerson {
    $tbName.Text = "$($First[$Rng.Next($First.Count)]) $($Last[$Rng.Next($Last.Count)])"
    $tbPolicy.Text = 'SC-DEMO-' + $Rng.Next(1100, 9999)
    $out = (Get-Date).Date.AddDays(-$Rng.Next(20, 120))
    $dpOut.Value = $out
    $dpBack.Value = $out.AddDays($Rng.Next(3, 10))
    $tbCause.Text = $Causes[$Rng.Next($Causes.Count)]
    Update-EmailText
}

function Set-Preset([string]$p) {
    for ($i = 0; $i -lt $clbFiles.Items.Count; $i++) { $clbFiles.SetItemChecked($i, $true) }
    switch ($p) {
        'Clean claim'              { $nuDelay.Value = 9; $nuMailDelay.Value = 9 }
        'Email contradicts documents' { $nuDelay.Value = 8; $nuMailDelay.Value = 6 }
        'Missing boarding pass'    { $nuDelay.Value = 9; $nuMailDelay.Value = 9
                                     $clbFiles.SetItemChecked($clbFiles.Items.IndexOf('boarding-pass-sc202.pdf'), $false) }
        'Short delay (2 hours)'    { $nuDelay.Value = 2; $nuMailDelay.Value = 2 }
    }
    Update-EmailText
}

# ---------------------------------------------------------------- actions
function Invoke-Generate([bool]$Email) {
    try { $c = Get-Claim } catch {
        [System.Windows.Forms.MessageBox]::Show('Scheduled departure must be a time like 11:40.', 'SureCover Tester') | Out-Null; return
    }
    if (-not $c.Name) { [System.Windows.Forms.MessageBox]::Show('Enter a claimant name.', 'SureCover Tester') | Out-Null; return }
    $include = @($clbFiles.CheckedItems | ForEach-Object { [string]$_ })
    $safe = ($c.Name -replace '[^\w\- ]', '').Trim()
    $folder = Join-Path $OutRoot ("{0} {1} - {2}" -f (Get-Date -Format 'yyyy-MM-dd HHmm'), $safe, $cbPreset.Text)
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    Build-Files $c $folder $include
    $info = "To: $($tbTo2.Text)`r`nSubject: $($tbSubject.Text)`r`n`r`n$($tbBody.Text)`r`n`r`nAttachments:`r`n" + (($AllFiles | Where-Object { $include -contains $_ }) -join "`r`n")
    [System.IO.File]::WriteAllText((Join-Path $folder 'email.txt'), $info)
    $lblStatus.Text = "Saved to $folder"

    if (-not $Email) { Start-Process explorer.exe $folder; return }
    try {
        $ol = New-Object -ComObject Outlook.Application
        $m = $ol.CreateItem(0)
        $m.To = $tbTo2.Text
        $m.Subject = $tbSubject.Text
        $m.Body = $tbBody.Text
        foreach ($f in $AllFiles) { if ($include -contains $f) { [void]$m.Attachments.Add((Join-Path $folder $f)) } }
        if ($chkSend.Checked) { $m.Send(); $lblStatus.Text = "Sent. Files in $folder" }
        else { $m.Display(); $lblStatus.Text = "Email opened in Outlook. Files in $folder" }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Couldn't open classic Outlook:`n$($_.Exception.Message)`n`nThe files and email.txt are in:`n$folder", 'SureCover Tester') | Out-Null
        Start-Process explorer.exe $folder
    }
}

# ---------------------------------------------------------------- UI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SureCover Tester'
$form.Size = New-Object System.Drawing.Size(760, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false

function Add-Label($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label; $l.Text = $text; $l.Location = "$x,$($y + 3)"; $l.AutoSize = $true
    $form.Controls.Add($l)
}
function Add-Box($x, $y, $w, $text = '') {
    $t = New-Object System.Windows.Forms.TextBox; $t.Location = "$x,$y"; $t.Width = $w; $t.Text = $text
    $form.Controls.Add($t); return $t
}
function Add-Date($x, $y) {
    $d = New-Object System.Windows.Forms.DateTimePicker; $d.Location = "$x,$y"; $d.Width = 160
    $d.Format = 'Custom'; $d.CustomFormat = 'd MMM yyyy'; $form.Controls.Add($d); return $d
}
function Add-Num($x, $y, $v) {
    $n = New-Object System.Windows.Forms.NumericUpDown; $n.Location = "$x,$y"; $n.Width = 60
    $n.Minimum = 0; $n.Maximum = 72; $n.Value = $v; $form.Controls.Add($n); return $n
}

Add-Label 'Test case' 16 16
$cbPreset = New-Object System.Windows.Forms.ComboBox
$cbPreset.Location = '130,16'; $cbPreset.Width = 220; $cbPreset.DropDownStyle = 'DropDownList'
[void]$cbPreset.Items.AddRange(@('Clean claim', 'Email contradicts documents', 'Missing boarding pass', 'Short delay (2 hours)'))
$form.Controls.Add($cbPreset)
$btnRandom = New-Object System.Windows.Forms.Button
$btnRandom.Text = 'New random person'; $btnRandom.Location = '370,15'; $btnRandom.Width = 150
$form.Controls.Add($btnRandom)

Add-Label 'Claimant' 16 56;          $tbName = Add-Box 130 56 220
Add-Label 'Policy ref' 370 56;        $tbPolicy = Add-Box 450 56 140
Add-Label 'Home airport' 16 90;       $tbFrom = Add-Box 130 90 60 'SIN'
Add-Label 'Destination' 370 90;       $tbTo = Add-Box 450 90 60 'NRT'
Add-Label 'SC101 out' 16 124;         $dpOut = Add-Date 130 124
Add-Label 'SC202 back' 370 124;       $dpBack = Add-Date 450 124
Add-Label 'SC202 departs' 16 158;     $tbDep = Add-Box 130 158 60 '11:40'
Add-Label 'Delay cause' 370 158;      $tbCause = Add-Box 450 158 270
Add-Label 'Return place' 16 192;      $tbReturnPlace = Add-Box 130 192 220 'Singapore'
Add-Label 'Delay in docs (h)' 16 226; $nuDelay = Add-Num 130 226 9
Add-Label 'Delay in email (h)' 370 226; $nuMailDelay = Add-Num 490 226 9

Add-Label 'Attachments' 16 264
$clbFiles = New-Object System.Windows.Forms.CheckedListBox
$clbFiles.Location = '130,264'; $clbFiles.Size = New-Object System.Drawing.Size(220, 130); $clbFiles.CheckOnClick = $true
[void]$clbFiles.Items.AddRange($AllFiles)
$form.Controls.Add($clbFiles)

Add-Label 'To' 16 410;      $tbTo2 = Add-Box 130 410 590 $DefaultRecipient
Add-Label 'Subject' 16 444; $tbSubject = Add-Box 130 444 590
Add-Label 'Body' 16 478
$tbBody = New-Object System.Windows.Forms.TextBox
$tbBody.Location = '130,478'; $tbBody.Size = New-Object System.Drawing.Size(590, 140)
$tbBody.Multiline = $true; $tbBody.ScrollBars = 'Vertical'
$form.Controls.Add($tbBody)
$btnText = New-Object System.Windows.Forms.Button
$btnText.Text = 'Refresh text'; $btnText.Location = '16,504'; $btnText.Width = 100
$form.Controls.Add($btnText)

$chkSend = New-Object System.Windows.Forms.CheckBox
$chkSend.Text = 'Send immediately (skip review)'; $chkSend.Location = '130,630'; $chkSend.AutoSize = $true
$form.Controls.Add($chkSend)

$btnFiles = New-Object System.Windows.Forms.Button
$btnFiles.Text = 'Generate files only'; $btnFiles.Location = '400,626'; $btnFiles.Width = 150
$form.Controls.Add($btnFiles)
$btnMail = New-Object System.Windows.Forms.Button
$btnMail.Text = 'Generate + email'; $btnMail.Location = '570,626'; $btnMail.Width = 150
$form.Controls.Add($btnMail)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = '16,670'; $lblStatus.Size = New-Object System.Drawing.Size(710, 36)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text = "Files are saved under $OutRoot"
$form.Controls.Add($lblStatus)

$cbPreset.Add_SelectedIndexChanged({ Set-Preset $cbPreset.Text })
$btnRandom.Add_Click({ Set-RandomPerson })
$btnText.Add_Click({ Update-EmailText })
$btnFiles.Add_Click({ Invoke-Generate $false })
$btnMail.Add_Click({ Invoke-Generate $true })

Set-RandomPerson
$cbPreset.SelectedIndex = 0
[void]$form.ShowDialog()
