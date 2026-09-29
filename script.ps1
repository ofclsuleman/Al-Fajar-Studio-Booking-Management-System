# ================================================================
# Al Fajar Studio — Booking Management System
# Complete single-file PowerShell Windows Forms application
# ================================================================
#
# Save this entire script as:
#     AlFajarStudio.ps1
#
# Data file created automatically:
#     Data\bookings.json
#
# Login:
#     112     -> Aqeel Ahmed / 03004115099
#     725442  -> M. Suleman Raghib / 03203622923
#
# Built with:
#     System.Windows.Forms
#     System.Drawing
#
# ================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ================================================================
# GLOBALS
# ================================================================

$script:StudioName = 'Al Fajar Studio'

if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $script:AppDirectory = [Environment]::CurrentDirectory
}
else {
    $script:AppDirectory = $PSScriptRoot
}

$script:DataDirectory = Join-Path $script:AppDirectory 'Data'
$script:DataFile = Join-Path $script:DataDirectory 'bookings.json'

$script:Bookings = @()
$script:CurrentUser = $null
$script:DataLoadError = $null

$script:CurrentBooking = $null
$script:IsEditMode = $false

$script:WizardForm = $null
$script:WizardPages = @{}
$script:WizardControls = @{}
$script:WizardEventRows = @()
$script:WizardSetupRows = @()
$script:WizardDeliverableRows = @()
$script:WizardStep = 1

# ================================================================
# THEME
# ================================================================

$script:Theme = @{
    Background  = [System.Drawing.Color]::FromArgb(246, 248, 251)
    Surface     = [System.Drawing.Color]::White
    Primary     = [System.Drawing.Color]::FromArgb(36, 91, 140)
    PrimaryDark = [System.Drawing.Color]::FromArgb(27, 68, 105)
    Accent      = [System.Drawing.Color]::FromArgb(71, 132, 183)
    Text        = [System.Drawing.Color]::FromArgb(35, 42, 50)
    Muted       = [System.Drawing.Color]::FromArgb(102, 112, 122)
    Border      = [System.Drawing.Color]::FromArgb(218, 224, 230)
    Danger      = [System.Drawing.Color]::FromArgb(180, 55, 55)
    Success     = [System.Drawing.Color]::FromArgb(32, 125, 78)
}

# ================================================================
# OPTIONS
# ================================================================

$script:EventOptions = @(
    'Mehndi',
    'Barat',
    'Walima',
    'Engagment',
    'Birthday',
    'Mehfil',
    'Custom'
)

$script:TimingOptions = @(
    '12:00pm To 04:00pm',
    '06:00pm To 10:00pm',
    'Custom'
)

$script:EquipmentOptions = @(
    'DSLR Photo Camera',
    'Sony Photo Camera',
    'DSLR Video Camera',
    'Sony Video Camera',
    'Cinematic Video Camera',
    'Drone Camera',
    'Custom'
)

$script:DeliverableOptions = @(
    'Complete Edited Film - Softcopy',
    'Complete Photos - Softcopy',
    'Raw Data - Softcopy',
    'Indian Album',
    'Printed Photo',
    'USB',
    'Custom'
)

# ================================================================
# DATA STORE
# ================================================================

function Initialize-DataStore {

    try {

        if (-not (Test-Path -LiteralPath $script:DataDirectory)) {

            New-Item `
                -ItemType Directory `
                -Path $script:DataDirectory `
                -Force |
                Out-Null
        }

        if (-not (Test-Path -LiteralPath $script:DataFile)) {

            [System.IO.File]::WriteAllText(
                $script:DataFile,
                '[]',
                [System.Text.UTF8Encoding]::new($false)
            )
        }
    }
    catch {

        throw (
            "Unable to initialize local data storage.`r`n`r`n" +
            $_.Exception.Message
        )
    }
}

function Normalize-Booking {

    param(
        [Parameter(Mandatory)]
        $Booking
    )

    $result = [PSCustomObject][ordered]@{
        BookingId        = [string]$Booking.BookingId
        ClientName       = [string]$Booking.ClientName
        ClientPhone      = [string]$Booking.ClientPhone
        ClientPhone2     = [string]$Booking.ClientPhone2
        EventDays        = @()
        Deliverables     = @()
        TotalCharges     = [decimal]0
        AdvancedReceived = [decimal]0
        BookedBy         = [string]$Booking.BookedBy
        BookedByContact  = [string]$Booking.BookedByContact
        StudioName       = $script:StudioName
        CreatedAt        = [string]$Booking.CreatedAt
        UpdatedAt        = [string]$Booking.UpdatedAt
        WhatsAppMessage  = [string]$Booking.WhatsAppMessage
    }

    try {
        $result.TotalCharges = [decimal]$Booking.TotalCharges
    }
    catch {
        $result.TotalCharges = [decimal]0
    }

    try {
        $result.AdvancedReceived = [decimal]$Booking.AdvancedReceived
    }
    catch {
        $result.AdvancedReceived = [decimal]0
    }

    foreach ($day in @($Booking.EventDays)) {

        if ($null -eq $day) {
            continue
        }

        $newDay = [PSCustomObject][ordered]@{
            Date            = [string]$day.Date
            EventType       = [string]$day.EventType
            CustomEventName = [string]$day.CustomEventName
            TimingType      = [string]$day.TimingType
            CustomTiming    = [string]$day.CustomTiming
            Location        = [string]$day.Location
            Equipment       = @()
        }

        foreach ($equipment in @($day.Equipment)) {

            if ($null -eq $equipment) {
                continue
            }

            $quantity = 0

            try {
                $quantity = [int]$equipment.Quantity
            }
            catch {
                $quantity = 0
            }

            $newDay.Equipment += [PSCustomObject][ordered]@{
                Name     = [string]$equipment.Name
                Quantity = $quantity
            }
        }

        $result.EventDays += $newDay
    }

    foreach ($deliverable in @($Booking.Deliverables)) {

        if ($null -eq $deliverable) {
            continue
        }

        $quantity = 0

        try {
            $quantity = [int]$deliverable.Quantity
        }
        catch {
            $quantity = 0
        }

        $result.Deliverables += [PSCustomObject][ordered]@{
            Name     = [string]$deliverable.Name
            Quantity = $quantity
        }
    }

    return $result
}

function Load-Bookings {

    $script:DataLoadError = $null

    try {

        Initialize-DataStore

        $raw = [System.IO.File]::ReadAllText(
            $script:DataFile
        )

        if ([string]::IsNullOrWhiteSpace($raw)) {
            $raw = '[]'
        }

        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop

        $loaded = @()

        if ($null -ne $parsed) {

            foreach ($booking in @($parsed)) {

                $loaded += Normalize-Booking $booking
            }
        }

        $script:Bookings = @($loaded)
    }
    catch {

        $script:Bookings = @()

        $script:DataLoadError =
            "Unable to read the booking database.`r`n`r`n" +
            $_.Exception.Message +
            "`r`n`r`nFile:`r`n" +
            $script:DataFile
    }
}

function Save-Bookings {

    param(
        [Parameter(Mandatory)]
        [array]$BookingCollection
    )

    if ($script:DataLoadError) {

        throw (
            "The booking database could not be safely read. " +
            "Saving has been blocked to protect existing information."
        )
    }

    $tempFile = $script:DataFile + '.tmp'
    $backupFile = $script:DataFile + '.bak'

    try {

        $safeCollection = @()

        foreach ($booking in @($BookingCollection)) {

            $safeCollection += Normalize-Booking $booking
        }

        if ($safeCollection.Count -eq 0) {
            $json = '[]'
        }
        else {
            $json = @($safeCollection) |
                ConvertTo-Json -Depth 12
        }

        [System.IO.File]::WriteAllText(
            $tempFile,
            $json,
            [System.Text.UTF8Encoding]::new($false)
        )

        if (Test-Path -LiteralPath $script:DataFile) {

            try {

                [System.IO.File]::Replace(
                    $tempFile,
                    $script:DataFile,
                    $backupFile,
                    $true
                )
            }
            catch {

                # Safe fallback for file systems where Replace is unavailable.
                [System.IO.File]::Copy(
                    $script:DataFile,
                    $backupFile,
                    $true
                )

                [System.IO.File]::Delete(
                    $script:DataFile
                )

                try {

                    [System.IO.File]::Move(
                        $tempFile,
                        $script:DataFile
                    )
                }
                catch {

                    if (Test-Path -LiteralPath $backupFile) {

                        [System.IO.File]::Copy(
                            $backupFile,
                            $script:DataFile,
                            $true
                        )
                    }

                    throw
                }
            }
        }
        else {

            [System.IO.File]::Move(
                $tempFile,
                $script:DataFile
            )
        }
    }
    catch {

        if (Test-Path -LiteralPath $tempFile) {

            Remove-Item `
                -LiteralPath $tempFile `
                -Force `
                -ErrorAction SilentlyContinue
        }

        throw (
            "Unable to safely save the booking database.`r`n`r`n" +
            $_.Exception.Message
        )
    }
}

# ================================================================
# MESSAGE HELPERS
# ================================================================

function Show-ErrorMessage {

    param(
        [string]$Message,
        [string]$Title = 'Al Fajar Studio'
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-InfoMessage {

    param(
        [string]$Message,
        [string]$Title = 'Al Fajar Studio'
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Show-ConfirmMessage {

    param(
        [string]$Message,
        [string]$Title = 'Confirm'
    )

    return (
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        ) -eq [System.Windows.Forms.DialogResult]::Yes
    )
}

# ================================================================
# CONTROL FACTORIES
# ================================================================

function New-Label {

    param(
        [string]$Text = '',
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 120,
        [int]$Height = 28,
        [int]$FontSize = 10,
        [switch]$Bold,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Empty
    )

    $label = New-Object System.Windows.Forms.Label

    $label.Text = $Text

    $label.Location = New-Object System.Drawing.Point(
        $X,
        $Y
    )

    $label.Size = New-Object System.Drawing.Size(
        $Width,
        $Height
    )

    if ($Bold) {

        $label.Font = New-Object System.Drawing.Font(
            'Segoe UI',
            $FontSize,
            [System.Drawing.FontStyle]::Bold
        )
    }
    else {

        $label.Font = New-Object System.Drawing.Font(
            'Segoe UI',
            $FontSize,
            [System.Drawing.FontStyle]::Regular
        )
    }

    if ($Color.IsEmpty) {
        $label.ForeColor = $script:Theme.Text
    }
    else {
        $label.ForeColor = $Color
    }

    $label.BackColor =
        [System.Drawing.Color]::Transparent

    return $label
}

function New-Button {

    param(
        [string]$Text = '',
        [int]$Width = 120,
        [int]$Height = 38,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::Empty,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Empty
    )

    $button = New-Object System.Windows.Forms.Button

    $button.Text = $Text

    $button.Size = New-Object System.Drawing.Size(
        $Width,
        $Height
    )

    $button.FlatStyle =
        [System.Windows.Forms.FlatStyle]::Flat

    $button.FlatAppearance.BorderSize = 0

    $button.Cursor =
        [System.Windows.Forms.Cursors]::Hand

    $button.Font = New-Object System.Drawing.Font(
        'Segoe UI Semibold',
        10
    )

    if ($BackColor.IsEmpty) {
        $button.BackColor = $script:Theme.Primary
    }
    else {
        $button.BackColor = $BackColor
    }

    if ($ForeColor.IsEmpty) {
        $button.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $button.ForeColor = $ForeColor
    }

    return $button
}

function New-TextBox {

    param(
        [int]$Width = 250,
        [int]$Height = 32,
        [switch]$Multiline,
        [switch]$ReadOnly,
        [switch]$Password
    )

    $textbox = New-Object System.Windows.Forms.TextBox

    $textbox.Width = $Width
    $textbox.Height = $Height

    $textbox.Font = New-Object System.Drawing.Font(
        'Segoe UI',
        10
    )

    $textbox.BorderStyle =
        [System.Windows.Forms.BorderStyle]::FixedSingle

    $textbox.BackColor =
        [System.Drawing.Color]::White

    $textbox.ForeColor =
        $script:Theme.Text

    $textbox.Multiline = $Multiline
    $textbox.ReadOnly = $ReadOnly

    if ($Password) {
        $textbox.UseSystemPasswordChar = $true
    }

    if ($Multiline) {
        $textbox.ScrollBars =
            [System.Windows.Forms.ScrollBars]::Vertical
    }

    return $textbox
}

function New-ComboBox {

    param(
        [int]$Width = 220
    )

    $combo = New-Object System.Windows.Forms.ComboBox

    $combo.Width = $Width
    $combo.Height = 32

    $combo.Font = New-Object System.Drawing.Font(
        'Segoe UI',
        9
    )

    $combo.DropDownStyle =
        [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $combo.FlatStyle =
        [System.Windows.Forms.FlatStyle]::Flat

    $combo.BackColor =
        [System.Drawing.Color]::White

    $combo.ForeColor =
        $script:Theme.Text

    return $combo
}

function New-SectionGroup {

    param(
        [string]$Text,
        [int]$Width = 900,
        [int]$Height = 150
    )

    $group = New-Object System.Windows.Forms.GroupBox

    $group.Text = $Text
    $group.Width = $Width
    $group.Height = $Height

    $group.Font = New-Object System.Drawing.Font(
        'Segoe UI Semibold',
        10
    )

    $group.ForeColor =
        $script:Theme.PrimaryDark

    $group.Padding =
        New-Object System.Windows.Forms.Padding(
            10,
            25,
            10,
            10
        )

    return $group
}

function Set-FormDefaults {

    param(
        [System.Windows.Forms.Form]$Form
    )

    $Form.BackColor =
        $script:Theme.Background

    $Form.Font = New-Object System.Drawing.Font(
        'Segoe UI',
        10
    )

    $Form.StartPosition =
        [System.Windows.Forms.FormStartPosition]::CenterScreen

    $Form.AutoScaleMode =
        [System.Windows.Forms.AutoScaleMode]::Dpi
}

# ================================================================
# VALUE HELPERS
# ================================================================

function Try-GetPositiveInt {

    param(
        [string]$Text,
        [ref]$Value
    )

    $number = 0

    $success = [int]::TryParse(
        $Text.Trim(),
        [System.Globalization.NumberStyles]::Integer,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )

    if ($success -and $number -gt 0) {

        $Value.Value = $number

        return $true
    }

    return $false
}

function Try-GetDecimalAmount {

    param(
        [string]$Text,
        [ref]$Value
    )

    $cleanText = $Text.Trim() -replace ',', ''

    $amount = [decimal]0

    $success = [decimal]::TryParse(
        $cleanText,
        [System.Globalization.NumberStyles]::Number,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$amount
    )

    if ($success -and $amount -ge 0) {

        $Value.Value = $amount

        return $true
    }

    return $false
}

function Format-Amount {

    param(
        [decimal]$Amount
    )

    return $Amount.ToString(
        '#,##0.##',
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

# ================================================================
# BOOKING OBJECTS
# ================================================================

function New-EventDayObject {

    param(
        [datetime]$Date = (Get-Date).Date
    )

    return [PSCustomObject][ordered]@{
        Date            = $Date.ToString('yyyy-MM-dd')
        EventType       = ''
        CustomEventName = ''
        TimingType      = ''
        CustomTiming    = ''
        Location        = ''
        Equipment       = @()
    }
}

function New-BookingDraft {

    return [PSCustomObject][ordered]@{
        BookingId        = ''
        ClientName       = ''
        ClientPhone      = ''
        ClientPhone2     = ''
        EventDays        = @(
            New-EventDayObject
        )
        Deliverables     = @()
        TotalCharges     = [decimal]0
        AdvancedReceived = [decimal]0
        BookedBy         = $script:CurrentUser.Name
        BookedByContact  = $script:CurrentUser.Contact
        StudioName       = $script:StudioName
        CreatedAt        = ''
        UpdatedAt        = ''
        WhatsAppMessage  = ''
    }
}

function Clone-BookingObject {

    param(
        [Parameter(Mandatory)]
        $Booking
    )

    return Normalize-Booking $Booking
}

function Get-ActiveEventName {

    param($Day)

    if ([string]$Day.EventType -eq 'Custom') {
        return [string]$Day.CustomEventName
    }

    return [string]$Day.EventType
}

function Get-ActiveTiming {

    param($Day)

    if ([string]$Day.TimingType -eq 'Custom') {
        return [string]$Day.CustomTiming
    }

    return [string]$Day.TimingType
}

# ================================================================
# BOOKING ID
# ================================================================

function Generate-BookingID {

    $testTime = Get-Date

    while ($true) {

        $candidate =
            'AFS-' +
            $testTime.ToString('yyMMHHmm')

        $duplicate = $script:Bookings |
            Where-Object {
                [string]$_.BookingId -eq $candidate
            } |
            Select-Object -First 1

        if ($null -eq $duplicate) {
            return $candidate
        }

        $testTime = $testTime.AddMinutes(1)
    }
}

# ================================================================
# CLIPBOARD
# ================================================================

function Copy-ToClipboard {

    param(
        [string]$Text
    )

    try {

        if ([string]::IsNullOrWhiteSpace($Text)) {
            throw 'The WhatsApp message is empty.'
        }

        [System.Windows.Forms.Clipboard]::SetText($Text)

        return $true
    }
    catch {

        Show-ErrorMessage `
            -Message (
                "The message could not be copied to the Windows Clipboard.`r`n`r`n" +
                $_.Exception.Message
            ) `
            -Title 'Clipboard Error'

        return $false
    }
}

# ================================================================
# WHATSAPP MESSAGE
#
# This function intentionally does NOT use the PowerShell -f
# formatting operator.
# ================================================================

function Generate-WhatsAppMessage {

    param(
        [Parameter(Mandatory)]
        $Booking
    )

    $lines =
        New-Object System.Collections.Generic.List[string]

    [void]$lines.Add('*AL FAJAR STUDIO*')
    [void]$lines.Add('')
    [void]$lines.Add('*BOOKING CONFIRMATION*')
    [void]$lines.Add('')
    [void]$lines.Add(
        '*Booking ID:* ' +
        [string]$Booking.BookingId
    )

    [void]$lines.Add('')
    [void]$lines.Add('*CLIENT INFORMATION*')
    [void]$lines.Add('')
    [void]$lines.Add(
        '*Client Name:* ' +
        [string]$Booking.ClientName
    )
    [void]$lines.Add(
        '*Phone Number:* ' +
        [string]$Booking.ClientPhone
    )

    if (-not [string]::IsNullOrWhiteSpace(
        [string]$Booking.ClientPhone2
    )) {

        [void]$lines.Add(
            '*Phone Number 2:* ' +
            [string]$Booking.ClientPhone2
        )
    }

    [void]$lines.Add('')
    [void]$lines.Add('---')
    [void]$lines.Add('')
    [void]$lines.Add('*EVENT DETAILS*')
    [void]$lines.Add('')

    $dayNumber = 1

    foreach ($day in @($Booking.EventDays)) {

        $eventName =
            Get-ActiveEventName $day

        $timing =
            Get-ActiveTiming $day

        $dateText =
            [string]$day.Date

        try {

            $dateText =
                ([datetime]::Parse($day.Date)).ToString(
                    'dd MMMM yyyy'
                )
        }
        catch {
        }

        [void]$lines.Add(
            '*Day ' +
            $dayNumber +
            ' — ' +
            $eventName +
            '*'
        )

        [void]$lines.Add(
            '*Date:* ' +
            $dateText
        )

        [void]$lines.Add(
            '*Timing:* ' +
            $timing
        )

        [void]$lines.Add(
            '*Location:* ' +
            [string]$day.Location
        )

        [void]$lines.Add('')
        [void]$lines.Add('*Equipment Setup:*')

        $equipmentList = @($day.Equipment)

        if ($equipmentList.Count -eq 0) {

            [void]$lines.Add(
                '• No equipment selected'
            )
        }
        else {

            foreach ($equipment in $equipmentList) {

                [void]$lines.Add(
                    '• ' +
                    [string]$equipment.Name +
                    ' — Qty: ' +
                    [string]$equipment.Quantity
                )
            }
        }

        [void]$lines.Add('')

        $dayNumber++
    }

    [void]$lines.Add('---')
    [void]$lines.Add('')
    [void]$lines.Add('*DELIVERABLES*')
    [void]$lines.Add('')

    $deliverableList = @(
        $Booking.Deliverables
    )

    if ($deliverableList.Count -eq 0) {

        [void]$lines.Add(
            '• No deliverables selected'
        )
    }
    else {

        foreach ($deliverable in $deliverableList) {

            [void]$lines.Add(
                '• ' +
                [string]$deliverable.Name +
                ' — Qty: ' +
                [string]$deliverable.Quantity
            )
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add('---')
    [void]$lines.Add('')
    [void]$lines.Add('*PAYMENT DETAILS*')
    [void]$lines.Add('')

    [void]$lines.Add(
        '*Total Charges:* Rs. ' +
        (Format-Amount ([decimal]$Booking.TotalCharges))
    )

    [void]$lines.Add(
        '*Advanced Received:* Rs. ' +
        (Format-Amount ([decimal]$Booking.AdvancedReceived))
    )

    [void]$lines.Add('')
    [void]$lines.Add('---')
    [void]$lines.Add('')
    [void]$lines.Add('*BOOKED BY*')

    [void]$lines.Add(
        [string]$Booking.BookedBy
    )

    [void]$lines.Add(
        [string]$Booking.BookedByContact
    )

    [void]$lines.Add('')
    [void]$lines.Add('*STUDIO*')
    [void]$lines.Add($script:StudioName)

    [void]$lines.Add('')
    [void]$lines.Add('---')
    [void]$lines.Add('')
    [void]$lines.Add('*NOTE*')
    [void]$lines.Add('')

    [void]$lines.Add(
        'Payment must be cleared on the last working day of the event.'
    )

    [void]$lines.Add('')
    [void]$lines.Add(
        '*Thank you for choosing Al Fajar Studio.*'
    )
    [void]$lines.Add(
        'We look forward to serving you.'
    )

    return (
        $lines -join [Environment]::NewLine
    )
}

# ================================================================
# FULL BOOKING VALIDATION
# ================================================================

function Validate-Booking {

    param(
        [Parameter(Mandatory)]
        $Booking
    )

    if ([string]::IsNullOrWhiteSpace(
        [string]$Booking.ClientName
    )) {

        return [PSCustomObject]@{
            Valid = $false
            Message = 'Client Name is required.'
        }
    }

    if ([string]::IsNullOrWhiteSpace(
        [string]$Booking.ClientPhone
    )) {

        return [PSCustomObject]@{
            Valid = $false
            Message = 'Client Phone Number is required.'
        }
    }

    $eventDays = @($Booking.EventDays)

    if ($eventDays.Count -lt 1) {

        return [PSCustomObject]@{
            Valid = $false
            Message = 'At least one event day is required.'
        }
    }

    $dayNumber = 1

    foreach ($day in $eventDays) {

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.Date
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message = 'Date is required for Day ' + $dayNumber + '.'
            }
        }

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.EventType
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message = 'Event is required for Day ' + $dayNumber + '.'
            }
        }

        if (
            ([string]$day.EventType -eq 'Custom') -and
            [string]::IsNullOrWhiteSpace(
                [string]$day.CustomEventName
            )
        ) {

            return [PSCustomObject]@{
                Valid = $false
                Message =
                    'Custom Event Name is required for Day ' +
                    $dayNumber +
                    '.'
            }
        }

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.TimingType
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message = 'Timing is required for Day ' + $dayNumber + '.'
            }
        }

        if (
            ([string]$day.TimingType -eq 'Custom') -and
            [string]::IsNullOrWhiteSpace(
                [string]$day.CustomTiming
            )
        ) {

            return [PSCustomObject]@{
                Valid = $false
                Message =
                    'Custom Timing is required for Day ' +
                    $dayNumber +
                    '.'
            }
        }

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.Location
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message =
                    'Location is required for Day ' +
                    $dayNumber +
                    '.'
            }
        }

        foreach ($equipment in @($day.Equipment)) {

            if ([string]::IsNullOrWhiteSpace(
                [string]$equipment.Name
            )) {

                return [PSCustomObject]@{
                    Valid = $false
                    Message =
                        'An equipment item on Day ' +
                        $dayNumber +
                        ' has no name.'
                }
            }

            $quantity = 0

            if (-not (
                Try-GetPositiveInt `
                    -Text ([string]$equipment.Quantity) `
                    -Value ([ref]$quantity)
            )) {

                return [PSCustomObject]@{
                    Valid = $false
                    Message =
                        'Invalid quantity for "' +
                        [string]$equipment.Name +
                        '" on Day ' +
                        $dayNumber +
                        '.'
                }
            }
        }

        $dayNumber++
    }

    foreach ($deliverable in @($Booking.Deliverables)) {

        if ([string]::IsNullOrWhiteSpace(
            [string]$deliverable.Name
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message = 'A selected deliverable has no name.'
            }
        }

        $quantity = 0

        if (-not (
            Try-GetPositiveInt `
                -Text ([string]$deliverable.Quantity) `
                -Value ([ref]$quantity)
        )) {

            return [PSCustomObject]@{
                Valid = $false
                Message =
                    'Invalid quantity for "' +
                    [string]$deliverable.Name +
                    '".'
            }
        }
    }

    $total = [decimal]0
    $advanced = [decimal]0

    if (-not (
        Try-GetDecimalAmount `
            -Text ([string]$Booking.TotalCharges) `
            -Value ([ref]$total)
    )) {

        return [PSCustomObject]@{
            Valid = $false
            Message = 'Total Charges must be a valid numeric amount.'
        }
    }

    if (-not (
        Try-GetDecimalAmount `
            -Text ([string]$Booking.AdvancedReceived) `
            -Value ([ref]$advanced)
    )) {

        return [PSCustomObject]@{
            Valid = $false
            Message = 'Advanced Received must be a valid numeric amount.'
        }
    }

    return [PSCustomObject]@{
        Valid = $true
        Message = ''
    }
}

# ================================================================
# LOGIN
# ================================================================

function Show-Login {

    $form = New-Object System.Windows.Forms.Form

    $form.Text =
        'Al Fajar Studio — Login'

    $form.ClientSize =
        New-Object System.Drawing.Size(
            440,
            270
        )

    Set-FormDefaults $form

    $form.FormBorderStyle =
        [System.Windows.Forms.FormBorderStyle]::FixedDialog

    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false

    $title = New-Label `
        -Text 'Secure Login' `
        -X 45 `
        -Y 26 `
        -Width 350 `
        -Height 36 `
        -FontSize 20 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $passwordLabel = New-Label `
        -Text 'Password' `
        -X 55 `
        -Y 80 `
        -Width 320 `
        -Height 27 `
        -FontSize 10 `
        -Bold

    $passwordBox = New-TextBox `
        -Width 330 `
        -Height 38 `
        -Password

    $passwordBox.Location =
        New-Object System.Drawing.Point(
            55,
            110
        )

    $unlockButton = New-Button `
        -Text 'Unlock' `
        -Width 330 `
        -Height 40 `
        -BackColor $script:Theme.Primary

    $unlockButton.Location =
        New-Object System.Drawing.Point(
            55,
            158
        )

    $errorLabel = New-Label `
        -Text '' `
        -X 55 `
        -Y 207 `
        -Width 330 `
        -Height 27 `
        -FontSize 9 `
        -Bold `
        -Color $script:Theme.Danger

    $errorLabel.TextAlign =
        [System.Drawing.ContentAlignment]::MiddleCenter

    $form.Controls.AddRange(
        @(
            $title,
            $passwordLabel,
            $passwordBox,
            $unlockButton,
            $errorLabel
        )
    )

    $loginAction = {

        $password = $passwordBox.Text

        if ($password -eq '112') {

            $script:CurrentUser =
                [PSCustomObject]@{
                    Name    = 'Aqeel Ahmed'
                    Contact = '03004115099'
                }

            $form.DialogResult =
                [System.Windows.Forms.DialogResult]::OK

            $form.Close()

            return
        }

        if ($password -eq '725442') {

            $script:CurrentUser =
                [PSCustomObject]@{
                    Name    = 'M. Suleman Raghib'
                    Contact = '03203622923'
                }

            $form.DialogResult =
                [System.Windows.Forms.DialogResult]::OK

            $form.Close()

            return
        }

        $script:CurrentUser = $null

        $errorLabel.Text =
            'Incorrect password. Please try again.'

        $passwordBox.Clear()
        $passwordBox.Focus()
    }

    $unlockButton.Add_Click($loginAction)

    $passwordBox.Add_KeyDown({

        param(
            $sender,
            $eventArgs
        )

        if (
            $eventArgs.KeyCode -eq
            [System.Windows.Forms.Keys]::Enter
        ) {

            & $loginAction
        }
    })

    $form.AcceptButton = $unlockButton

    $form.Add_Shown({
        $passwordBox.Focus()
    })

    $result = $form.ShowDialog()

    $form.Dispose()

    return (
        ($result -eq [System.Windows.Forms.DialogResult]::OK) -and
        ($null -ne $script:CurrentUser)
    )
}

# ================================================================
# DASHBOARD
# ================================================================

function Show-Dashboard {

    $form = New-Object System.Windows.Forms.Form

    $form.Text =
        'Al Fajar Studio — Booking Management System'

    $form.ClientSize =
        New-Object System.Drawing.Size(
            1020,
            650
        )

    Set-FormDefaults $form

    $form.FormBorderStyle =
        [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $form.MaximizeBox = $false

    # Header
    $header = New-Object System.Windows.Forms.Panel

    $header.Dock = 'Top'
    $header.Height = 105
    $header.BackColor = $script:Theme.Surface

    $title = New-Label `
        -Text $script:StudioName `
        -X 30 `
        -Y 17 `
        -Width 600 `
        -Height 36 `
        -FontSize 21 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $subtitle = New-Label `
        -Text 'Booking Management System' `
        -X 30 `
        -Y 55 `
        -Width 500 `
        -Height 25 `
        -FontSize 10 `
        -Color $script:Theme.Muted

    $header.Controls.AddRange(
        @(
            $title,
            $subtitle
        )
    )

    # User information
    $userCard = New-Object System.Windows.Forms.Panel

    $userCard.Location =
        New-Object System.Drawing.Point(
            690,
            14
        )

    $userCard.Size =
        New-Object System.Drawing.Size(
            300,
            76
        )

    $userCard.BackColor =
        [System.Drawing.Color]::FromArgb(
            239,
            245,
            250
        )

    $bookedByLabel = New-Label `
        -Text 'Booked By' `
        -X 15 `
        -Y 8 `
        -Width 100 `
        -Height 20 `
        -FontSize 8 `
        -Bold `
        -Color $script:Theme.Muted

    $userName = New-Label `
        -Text $script:CurrentUser.Name `
        -X 15 `
        -Y 29 `
        -Width 190 `
        -Height 24 `
        -FontSize 10 `
        -Bold

    $userContact = New-Label `
        -Text $script:CurrentUser.Contact `
        -X 195 `
        -Y 9 `
        -Width 95 `
        -Height 20 `
        -FontSize 8 `
        -Color $script:Theme.Muted

    $userCard.Controls.AddRange(
        @(
            $bookedByLabel,
            $userName,
            $userContact
        )
    )

    $header.Controls.Add($userCard)

    $form.Controls.Add($header)

    # Content
    $content = New-Object System.Windows.Forms.Panel

    $content.Dock = 'Fill'

    $welcome = New-Label `
        -Text ('Welcome, ' + $script:CurrentUser.Name) `
        -X 35 `
        -Y 25 `
        -Width 600 `
        -Height 35 `
        -FontSize 18 `
        -Bold

    $contact = New-Label `
        -Text ('Contact: ' + $script:CurrentUser.Contact) `
        -X 35 `
        -Y 60 `
        -Width 500 `
        -Height 25 `
        -FontSize 10 `
        -Color $script:Theme.Muted

    $addButton = New-Button `
        -Text 'Add Booking' `
        -Width 260 `
        -Height 95 `
        -BackColor $script:Theme.Primary

    $addButton.Location =
        New-Object System.Drawing.Point(
            155,
            145
        )

    $addButton.Font =
        New-Object System.Drawing.Font(
            'Segoe UI Semibold',
            14
        )

    $viewButton = New-Button `
        -Text 'View Booking' `
        -Width 260 `
        -Height 95 `
        -BackColor $script:Theme.Accent

    $viewButton.Location =
        New-Object System.Drawing.Point(
            445,
            145
        )

    $viewButton.Font =
        New-Object System.Drawing.Font(
            'Segoe UI Semibold',
            14
        )

    $savedPanel =
        New-Object System.Windows.Forms.Panel

    $savedPanel.Location =
        New-Object System.Drawing.Point(
            155,
            270
        )

    $savedPanel.Size =
        New-Object System.Drawing.Size(
            550,
            100
        )

    $savedPanel.BackColor =
        $script:Theme.Surface

    $savedPanel.BorderStyle =
        [System.Windows.Forms.BorderStyle]::FixedSingle

    $savedCount = New-Label `
        -Text (
            'Saved Bookings: ' +
            @($script:Bookings).Count
        ) `
        -X 18 `
        -Y 16 `
        -Width 480 `
        -Height 26 `
        -FontSize 11 `
        -Bold

    $dataPath = New-Label `
        -Text 'Local data: Data\bookings.json' `
        -X 18 `
        -Y 49 `
        -Width 480 `
        -Height 23 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $savedPanel.Controls.AddRange(
        @(
            $savedCount,
            $dataPath
        )
    )

    $footer = New-Label `
        -Text 'Studio Name: Al Fajar Studio' `
        -X 35 `
        -Y 555 `
        -Width 500 `
        -Height 25 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $content.Controls.AddRange(
        @(
            $welcome,
            $contact,
            $addButton,
            $viewButton,
            $savedPanel,
            $footer
        )
    )

    $form.Controls.Add($content)

    $addButton.Add_Click({

        Show-AddBooking -BookingToEdit $null

        Load-Bookings

        if (-not $form.IsDisposed) {

            $savedCount.Text =
                'Saved Bookings: ' +
                @($script:Bookings).Count
        }
    })

    $viewButton.Add_Click({

        Load-Bookings
        Show-ViewBookings
    })

    if ($script:DataLoadError) {

        Show-ErrorMessage `
            -Message $script:DataLoadError `
            -Title 'Booking Database Error'
    }

    [void]$form.ShowDialog()

    $form.Dispose()
}

# ================================================================
# WIZARD HEADER
# ================================================================

function New-WizardHeader {

    param(
        [System.Windows.Forms.Form]$Form
    )

    $header =
        New-Object System.Windows.Forms.Panel

    $header.Dock = 'Top'
    $header.Height = 88
    $header.BackColor = $script:Theme.Surface

    $title = New-Label `
        -Text 'Booking Wizard' `
        -X 25 `
        -Y 13 `
        -Width 500 `
        -Height 32 `
        -FontSize 18 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $subtitle = New-Label `
        -Text $script:StudioName `
        -X 25 `
        -Y 46 `
        -Width 400 `
        -Height 24 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $stepLabel = New-Label `
        -Text 'Step 1 of 4 — Event Details' `
        -X 655 `
        -Y 25 `
        -Width 300 `
        -Height 28 `
        -FontSize 9 `
        -Bold `
        -Color $script:Theme.Primary

    $stepLabel.TextAlign =
        [System.Drawing.ContentAlignment]::MiddleRight

    $header.Controls.AddRange(
        @(
            $title,
            $subtitle,
            $stepLabel
        )
    )

    $Form.Controls.Add($header)

    $script:WizardControls.StepLabel =
        $stepLabel
}

# ================================================================
# SET WIZARD STEP
# ================================================================

function Set-WizardStep {

    param(
        [int]$Step
    )

    $script:WizardStep = $Step

    foreach ($pageName in @(
        'Page1',
        'Page2',
        'Page3',
        'Page4'
    )) {

        if (
            $script:WizardPages.ContainsKey(
                $pageName
            )
        ) {

            $script:WizardPages[$pageName].Visible =
                $false
        }
    }

    $targetName =
        'Page' + $Step

    if (
        $script:WizardPages.ContainsKey(
            $targetName
        )
    ) {

        $script:WizardPages[$targetName].Visible =
            $true

        $script:WizardPages[$targetName].BringToFront()
    }

    switch ($Step) {

        1 {
            $script:WizardControls.StepLabel.Text =
                'Step 1 of 4 — Event Details'
        }

        2 {
            $script:WizardControls.StepLabel.Text =
                'Step 2 of 4 — Setup'
        }

        3 {
            $script:WizardControls.StepLabel.Text =
                'Step 3 of 4 — Deliverables & Payment'
        }

        4 {
            $script:WizardControls.StepLabel.Text =
                'Step 4 of 4 — Summary'
        }
    }

    $script:WizardControls.BackButton.Enabled =
        (
            $Step -gt 1 -and
            $Step -lt 4
        )

    $script:WizardControls.NextButton.Visible =
        ($Step -lt 3)

    $script:WizardControls.FinishButton.Visible =
        ($Step -eq 3)
}

# ================================================================
# PAGE 1
# ================================================================

function Build-Page1 {

    param(
        [System.Windows.Forms.Form]$Form
    )

    $panel =
        New-Object System.Windows.Forms.Panel

    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true

    # Client Information
    $clientGroup = New-SectionGroup `
        -Text 'Client Information' `
        -Width 920 `
        -Height 130

    $clientGroup.Location =
        New-Object System.Drawing.Point(
            25,
            18
        )

    $nameLabel = New-Label `
        -Text 'Client Name *' `
        -X 18 `
        -Y 25 `
        -Width 125 `
        -Height 25 `
        -FontSize 10 `
        -Bold

    $clientName = New-TextBox `
        -Width 230 `
        -Height 32

    $clientName.Location =
        New-Object System.Drawing.Point(
            145,
            20
        )

    $phoneLabel = New-Label `
        -Text 'Client Phone Number *' `
        -X 405 `
        -Y 25 `
        -Width 165 `
        -Height 25 `
        -FontSize 10 `
        -Bold

    $clientPhone = New-TextBox `
        -Width 215 `
        -Height 32

    $clientPhone.Location =
        New-Object System.Drawing.Point(
            570,
            20
        )

    $phone2Label = New-Label `
        -Text 'Client Phone Number 2' `
        -X 18 `
        -Y 76 `
        -Width 160 `
        -Height 25 `
        -FontSize 10 `
        -Bold

    $clientPhone2 = New-TextBox `
        -Width 230 `
        -Height 32

    $clientPhone2.Location =
        New-Object System.Drawing.Point(
            180,
            71
        )

    $clientGroup.Controls.AddRange(
        @(
            $nameLabel,
            $clientName,
            $phoneLabel,
            $clientPhone,
            $phone2Label,
            $clientPhone2
        )
    )

    $panel.Controls.Add($clientGroup)

    # Event Days
    $daysGroup = New-SectionGroup `
        -Text 'Event Days' `
        -Width 920 `
        -Height 420

    $daysGroup.Location =
        New-Object System.Drawing.Point(
            25,
            165
        )

    $addDayButton = New-Button `
        -Text 'Add Day' `
        -Width 105 `
        -Height 32 `
        -BackColor $script:Theme.Primary

    $addDayButton.Location =
        New-Object System.Drawing.Point(
            785,
            18
        )

    $removeDayButton = New-Button `
        -Text 'Remove Day' `
        -Width 120 `
        -Height 32 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $removeDayButton.Location =
        New-Object System.Drawing.Point(
            650,
            18
        )

    $daysGroup.Controls.AddRange(
        @(
            $removeDayButton,
            $addDayButton
        )
    )

    # Table
    $table =
        New-Object System.Windows.Forms.TableLayoutPanel

    $table.Location =
        New-Object System.Drawing.Point(
            15,
            58
        )

    $table.Size =
        New-Object System.Drawing.Size(
            880,
            340
        )

    $table.AutoScroll = $true
    $table.ColumnCount = 6
    $table.RowCount = 1

    $table.CellBorderStyle =
        [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single

    $columnWidths = @(
        125,
        120,
        145,
        140,
        145,
        180
    )

    for (
        $i = 0;
        $i -lt $columnWidths.Count;
        $i++
    ) {

        $table.ColumnStyles.Add(
            (
                New-Object System.Windows.Forms.ColumnStyle(
                    [System.Windows.Forms.SizeType]::Absolute,
                    $columnWidths[$i]
                )
            )
        )
    }

    $headers = @(
        'Date',
        'Event',
        'Custom Event',
        'Timing',
        'Custom Timing',
        'Location'
    )

    for (
        $i = 0;
        $i -lt $headers.Count;
        $i++
    ) {

        $header = New-Label `
            -Text $headers[$i] `
            -Width $columnWidths[$i] `
            -Height 32 `
            -FontSize 8 `
            -Bold `
            -Color $script:Theme.PrimaryDark

        $header.TextAlign =
            [System.Drawing.ContentAlignment]::MiddleCenter

        [void]$table.Controls.Add(
            $header,
            $i,
            0
        )
    }

    $table.RowStyles.Add(
        (
            New-Object System.Windows.Forms.RowStyle(
                [System.Windows.Forms.SizeType]::Absolute,
                38
            )
        )
    )

    $daysGroup.Controls.Add($table)
    $panel.Controls.Add($daysGroup)

    $script:WizardControls.ClientName =
        $clientName

    $script:WizardControls.ClientPhone =
        $clientPhone

    $script:WizardControls.ClientPhone2 =
        $clientPhone2

    $script:WizardControls.EventTable =
        $table

    $script:WizardPages.Page1 =
        $panel

    $Form.Controls.Add($panel)

    $script:WizardEventRows = @()

    $addDayButton.Add_Click({
        Add-EventDay
    })

    $removeDayButton.Add_Click({
        Remove-EventDay
    })

    if ($script:IsEditMode) {

        $clientName.Text =
            [string]$script:CurrentBooking.ClientName

        $clientPhone.Text =
            [string]$script:CurrentBooking.ClientPhone

        $clientPhone2.Text =
            [string]$script:CurrentBooking.ClientPhone2

        $existingDays =
            @($script:CurrentBooking.EventDays)

        if ($existingDays.Count -eq 0) {
            Add-EventDay
        }
        else {

            foreach ($existingDay in $existingDays) {

                Add-EventDay -FromExisting
            }
        }
    }
    else {
        Add-EventDay
    }
}

# ================================================================
# ADD EVENT DAY
# ================================================================

function Add-EventDay {

    param(
        [switch]$FromExisting
    )

    $index =
        $script:WizardEventRows.Count

    $table =
        $script:WizardControls.EventTable

    $datePicker =
        New-Object System.Windows.Forms.DateTimePicker

    $datePicker.Format =
        [System.Windows.Forms.DateTimePickerFormat]::Custom

    $datePicker.CustomFormat =
        'dd MMM yyyy'

    $datePicker.Width = 118
    $datePicker.Height = 30

    $datePicker.Font =
        New-Object System.Drawing.Font(
            'Segoe UI',
            8.5
        )

    $datePicker.Value =
        (Get-Date).Date

    $eventCombo =
        New-ComboBox -Width 112

    [void]$eventCombo.Items.AddRange(
        $script:EventOptions
    )

    $eventCombo.SelectedIndex = -1

    $customEvent =
        New-TextBox -Width 132 -Height 30

    $customEvent.Visible = $false

    $timingCombo =
        New-ComboBox -Width 130

    [void]$timingCombo.Items.AddRange(
        $script:TimingOptions
    )

    $timingCombo.SelectedIndex = -1

    $customTiming =
        New-TextBox -Width 132 -Height 30

    $customTiming.Visible = $false

    $location =
        New-TextBox -Width 168 -Height 30

    $row = [PSCustomObject][ordered]@{
        Date         = $datePicker
        Event        = $eventCombo
        CustomEvent  = $customEvent
        Timing       = $timingCombo
        CustomTiming = $customTiming
        Location     = $location
        ExistingDay  = $null
    }

    $eventCombo.Tag = $row
    $timingCombo.Tag = $row

    $eventCombo.Add_SelectedIndexChanged({

        $currentRow = $this.Tag

        if ($this.SelectedItem -eq 'Custom') {

            $currentRow.CustomEvent.Visible =
                $true
        }
        else {

            $currentRow.CustomEvent.Visible =
                $false
        }
    })

    $timingCombo.Add_SelectedIndexChanged({

        $currentRow = $this.Tag

        if ($this.SelectedItem -eq 'Custom') {

            $currentRow.CustomTiming.Visible =
                $true
        }
        else {

            $currentRow.CustomTiming.Visible =
                $false
        }
    })

    $script:WizardEventRows += $row

    $table.RowCount++

    $newRowIndex =
        $table.RowCount - 1

    $table.RowStyles.Add(
        (
            New-Object System.Windows.Forms.RowStyle(
                [System.Windows.Forms.SizeType]::Absolute,
                42
            )
        )
    )

    [void]$table.Controls.Add(
        $datePicker,
        0,
        $newRowIndex
    )

    [void]$table.Controls.Add(
        $eventCombo,
        1,
        $newRowIndex
    )

    [void]$table.Controls.Add(
        $customEvent,
        2,
        $newRowIndex
    )

    [void]$table.Controls.Add(
        $timingCombo,
        3,
        $newRowIndex
    )

    [void]$table.Controls.Add(
        $customTiming,
        4,
        $newRowIndex
    )

    [void]$table.Controls.Add(
        $location,
        5,
        $newRowIndex
    )

    if ($FromExisting) {

        $existingDays =
            @($script:CurrentBooking.EventDays)

        if ($index -lt $existingDays.Count) {

            $existing =
                $existingDays[$index]

            $row.ExistingDay =
                $existing

            try {

                $datePicker.Value =
                    [datetime]::Parse(
                        [string]$existing.Date
                    )
            }
            catch {

                $datePicker.Value =
                    (Get-Date).Date
            }

            if (-not [string]::IsNullOrWhiteSpace(
                [string]$existing.EventType
            )) {

                $eventCombo.SelectedItem =
                    [string]$existing.EventType
            }

            $customEvent.Text =
                [string]$existing.CustomEventName

            if (-not [string]::IsNullOrWhiteSpace(
                [string]$existing.TimingType
            )) {

                $timingCombo.SelectedItem =
                    [string]$existing.TimingType
            }

            $customTiming.Text =
                [string]$existing.CustomTiming

            $location.Text =
                [string]$existing.Location
        }
    }

    $table.PerformLayout()
}

# ================================================================
# REMOVE EVENT DAY
# ================================================================

function Remove-EventDay {

    if (
        $script:WizardEventRows.Count -le 1
    ) {

        Show-InfoMessage `
            -Message 'At least one event day is required.' `
            -Title 'Event Days'

        return
    }

    $lastIndex =
        $script:WizardEventRows.Count - 1

    $row =
        $script:WizardEventRows[$lastIndex]

    foreach ($control in @(
        $row.Date,
        $row.Event,
        $row.CustomEvent,
        $row.Timing,
        $row.CustomTiming,
        $row.Location
    )) {

        if (
            ($null -ne $control) -and
            ($null -ne $control.Parent)
        ) {

            $control.Parent.Controls.Remove(
                $control
            )

            $control.Dispose()
        }
    }

    $newRows = @()

    for (
        $i = 0;
        $i -lt $lastIndex;
        $i++
    ) {

        $newRows +=
            $script:WizardEventRows[$i]
    }

    $script:WizardEventRows =
        @($newRows)

    $table =
        $script:WizardControls.EventTable

    if ($table.RowCount -gt 1) {

        if ($table.RowStyles.Count -gt 1) {

            $table.RowStyles.RemoveAt(
                $table.RowStyles.Count - 1
            )
        }

        $table.RowCount--
    }

    $table.PerformLayout()
}

# ================================================================
# SYNC PAGE 1
# ================================================================

function Sync-Page1ToDraft {

    $booking =
        $script:CurrentBooking

    $booking.ClientName =
        $script:WizardControls.ClientName.Text.Trim()

    $booking.ClientPhone =
        $script:WizardControls.ClientPhone.Text.Trim()

    $booking.ClientPhone2 =
        $script:WizardControls.ClientPhone2.Text.Trim()

    $days = @()

    foreach ($row in @(
        $script:WizardEventRows
    )) {

        $eventType = ''

        if ($null -ne $row.Event.SelectedItem) {

            $eventType =
                [string]$row.Event.SelectedItem
        }

        $timingType = ''

        if ($null -ne $row.Timing.SelectedItem) {

            $timingType =
                [string]$row.Timing.SelectedItem
        }

        $equipment =
            if ($null -ne $row.ExistingDay) {
                @($row.ExistingDay.Equipment)
            }
            else {
                @()
            }

        $days +=
            [PSCustomObject][ordered]@{
                Date =
                    $row.Date.Value.ToString(
                        'yyyy-MM-dd'
                    )

                EventType =
                    $eventType

                CustomEventName =
                    $row.CustomEvent.Text.Trim()

                TimingType =
                    $timingType

                CustomTiming =
                    $row.CustomTiming.Text.Trim()

                Location =
                    $row.Location.Text.Trim()

                Equipment =
                    $equipment
            }
    }

    $booking.EventDays =
        @($days)
}

# ================================================================
# VALIDATE PAGE 1
# ================================================================

function Validate-Page1 {

    Sync-Page1ToDraft

    if ([string]::IsNullOrWhiteSpace(
        [string]$script:CurrentBooking.ClientName
    )) {

        Show-ErrorMessage `
            -Message 'Client Name is required.' `
            -Title 'Validation'

        $script:WizardControls.ClientName.Focus()

        return $false
    }

    if ([string]::IsNullOrWhiteSpace(
        [string]$script:CurrentBooking.ClientPhone
    )) {

        Show-ErrorMessage `
            -Message 'Client Phone Number is required.' `
            -Title 'Validation'

        $script:WizardControls.ClientPhone.Focus()

        return $false
    }

    if (@(
        $script:CurrentBooking.EventDays
    ).Count -lt 1) {

        Show-ErrorMessage `
            -Message 'At least one event day is required.' `
            -Title 'Validation'

        return $false
    }

    $dayNumber = 1

    foreach ($day in @(
        $script:CurrentBooking.EventDays
    )) {

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.EventType
        )) {

            Show-ErrorMessage `
                -Message (
                    'Please select an event for Day ' +
                    $dayNumber +
                    '.'
                ) `
                -Title 'Validation'

            return $false
        }

        if (
            ($day.EventType -eq 'Custom') -and
            [string]::IsNullOrWhiteSpace(
                [string]$day.CustomEventName
            )
        ) {

            Show-ErrorMessage `
                -Message (
                    'Please enter a Custom Event Name for Day ' +
                    $dayNumber +
                    '.'
                ) `
                -Title 'Validation'

            return $false
        }

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.TimingType
        )) {

            Show-ErrorMessage `
                -Message (
                    'Please select a timing for Day ' +
                    $dayNumber +
                    '.'
                ) `
                -Title 'Validation'

            return $false
        }

        if (
            ($day.TimingType -eq 'Custom') -and
            [string]::IsNullOrWhiteSpace(
                [string]$day.CustomTiming
            )
        ) {

            Show-ErrorMessage `
                -Message (
                    'Please enter Custom Timing for Day ' +
                    $dayNumber +
                    '.'
                ) `
                -Title 'Validation'

            return $false
        }

        if ([string]::IsNullOrWhiteSpace(
            [string]$day.Location
        )) {

            Show-ErrorMessage `
                -Message (
                    'Please enter a location for Day ' +
                    $dayNumber +
                    '.'
                ) `
                -Title 'Validation'

            return $false
        }

        $dayNumber++
    }

    return $true
}

# ================================================================
# BUILD SETUP PAGE
# ================================================================

function Build-SetupPage {

    param(
        [System.Windows.Forms.Form]$Form
    )

    if (
        $script:WizardPages.ContainsKey('Page2') -and
        $null -ne $script:WizardPages.Page2
    ) {

        $oldPage =
            $script:WizardPages.Page2

        $Form.Controls.Remove($oldPage)

        $oldPage.Dispose()
    }

    $panel =
        New-Object System.Windows.Forms.Panel

    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true

    $pageTitle = New-Label `
        -Text 'Setup' `
        -X 25 `
        -Y 12 `
        -Width 500 `
        -Height 35 `
        -FontSize 19 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $pageSubtitle = New-Label `
        -Text 'Select equipment for every event day and enter quantities for selected items.' `
        -X 25 `
        -Y 47 `
        -Width 850 `
        -Height 25 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $panel.Controls.AddRange(
        @(
            $pageTitle,
            $pageSubtitle
        )
    )

    $currentY = 80

    $setupRows = @()
    $dayNumber = 1

    foreach ($day in @(
        $script:CurrentBooking.EventDays
    )) {

        $eventName =
            Get-ActiveEventName $day

        $group = New-SectionGroup `
            -Text (
                'Day ' +
                $dayNumber +
                ' — ' +
                $eventName
            ) `
            -Width 900 `
            -Height 330

        $group.Location =
            New-Object System.Drawing.Point(
                25,
                $currentY
            )

        $items = @()
        $itemY = 30

        foreach ($option in $script:EquipmentOptions) {

            $check =
                New-Object System.Windows.Forms.CheckBox

            $check.Location =
                New-Object System.Drawing.Point(
                    20,
                    ($itemY + 2)
                )

            $check.Size =
                New-Object System.Drawing.Size(
                    24,
                    24
                )

            $equipmentLabel = New-Label `
                -Text $option `
                -X 50 `
                -Y $itemY `
                -Width 220 `
                -Height 25 `
                -FontSize 9 `
                -Bold

            $qtyLabel = New-Label `
                -Text 'Qty' `
                -X 275 `
                -Y $itemY `
                -Width 35 `
                -Height 24 `
                -FontSize 8 `
                -Color $script:Theme.Muted

            $quantityBox = New-TextBox `
                -Width 65 `
                -Height 30

            $quantityBox.Location =
                New-Object System.Drawing.Point(
                    310,
                    ($itemY - 3)
                )

            $quantityBox.TextAlign =
                [System.Windows.Forms.HorizontalAlignment]::Center

            $quantityBox.Enabled = $false

            $customLabel = $null
            $customNameBox = $null

            $isCustom =
                ($option -eq 'Custom')

            if ($isCustom) {

                $customLabel = New-Label `
                    -Text 'Custom Name' `
                    -X 410 `
                    -Y $itemY `
                    -Width 90 `
                    -Height 24 `
                    -FontSize 8 `
                    -Color $script:Theme.Muted

                $customNameBox = New-TextBox `
                    -Width 270 `
                    -Height 30

                $customNameBox.Location =
                    New-Object System.Drawing.Point(
                        505,
                        ($itemY - 3)
                    )

                $customLabel.Visible = $false
                $customNameBox.Visible = $false

                $group.Controls.AddRange(
                    @(
                        $customLabel,
                        $customNameBox
                    )
                )
            }

            $item = [PSCustomObject][ordered]@{
                Name        = $option
                Check       = $check
                Quantity    = $quantityBox
                CustomName  = $customNameBox
                CustomLabel = $customLabel
                IsCustom    = $isCustom
            }

            $check.Tag = $item

            $check.Add_CheckedChanged({

                $currentItem = $this.Tag

                $currentItem.Quantity.Enabled =
                    $currentItem.Check.Checked

                if ($currentItem.IsCustom) {

                    $currentItem.CustomName.Visible =
                        $currentItem.Check.Checked

                    $currentItem.CustomLabel.Visible =
                        $currentItem.Check.Checked
                }
            })

            $items += $item

            $group.Controls.AddRange(
                @(
                    $check,
                    $equipmentLabel,
                    $qtyLabel,
                    $quantityBox
                )
            )

            $itemY += 37
        }

        # Restore selected equipment
        $existingEquipment =
            @($day.Equipment)

        foreach ($item in @($items)) {

            if ($item.IsCustom) {

                $customExisting =
                    $existingEquipment |
                    Where-Object {
                        $script:EquipmentOptions -notcontains $_.Name
                    } |
                    Select-Object -First 1

                if ($null -ne $customExisting) {

                    $item.Check.Checked = $true

                    $item.Quantity.Text =
                        [string]$customExisting.Quantity

                    $item.CustomName.Text =
                        [string]$customExisting.Name
                }

                continue
            }

            $existing =
                $existingEquipment |
                Where-Object {
                    $_.Name -eq $item.Name
                } |
                Select-Object -First 1

            if ($null -ne $existing) {

                $item.Check.Checked = $true

                $item.Quantity.Text =
                    [string]$existing.Quantity
            }
        }

        $tip = New-Label `
            -Text 'Leave equipment unchecked when it is not required.' `
            -X 410 `
            -Y 282 `
            -Width 420 `
            -Height 24 `
            -FontSize 8 `
            -Color $script:Theme.Muted

        $group.Controls.Add($tip)

        $panel.Controls.Add($group)

        $setupRows +=
            [PSCustomObject][ordered]@{
                Day   = $day
                Items = $items
            }

        $currentY += 345
        $dayNumber++
    }

    $script:WizardSetupRows =
        @($setupRows)

    $script:WizardPages.Page2 =
        $panel

    $Form.Controls.Add($panel)

    $panel.BringToFront()
}

# ================================================================
# SYNC SETUP
# ================================================================

function Sync-SetupToDraft {

    $days =
        @($script:CurrentBooking.EventDays)

    for (
        $dayIndex = 0;
        $dayIndex -lt $script:WizardSetupRows.Count;
        $dayIndex++
    ) {

        $ui =
            $script:WizardSetupRows[$dayIndex]

        $equipmentList = @()

        foreach ($item in @($ui.Items)) {

            if (-not $item.Check.Checked) {
                continue
            }

            $quantity = 0

            [void](
                Try-GetPositiveInt `
                    -Text $item.Quantity.Text `
                    -Value ([ref]$quantity)
            )

            if ($item.IsCustom) {
                $equipmentName =
                    $item.CustomName.Text.Trim()
            }
            else {
                $equipmentName =
                    $item.Name
            }

            $equipmentList +=
                [PSCustomObject][ordered]@{
                    Name     = $equipmentName
                    Quantity = $quantity
                }
        }

        if ($dayIndex -lt $days.Count) {

            $days[$dayIndex].Equipment =
                @($equipmentList)
        }
    }

    $script:CurrentBooking.EventDays =
        @($days)
}

# ================================================================
# VALIDATE SETUP
# ================================================================

function Validate-SetupPage {

    Sync-SetupToDraft

    $dayNumber = 1

    foreach ($row in @(
        $script:WizardSetupRows
    )) {

        foreach ($item in @(
            $row.Items
        )) {

            if (-not $item.Check.Checked) {
                continue
            }

            if (
                $item.IsCustom -and
                [string]::IsNullOrWhiteSpace(
                    $item.CustomName.Text
                )
            ) {

                Show-ErrorMessage `
                    -Message (
                        'Please enter a Custom Equipment Name for Day ' +
                        $dayNumber +
                        '.'
                    ) `
                    -Title 'Validation'

                $item.CustomName.Focus()

                return $false
            }

            $quantity = 0

            if (-not (
                Try-GetPositiveInt `
                    -Text $item.Quantity.Text `
                    -Value ([ref]$quantity)
            )) {

                $equipmentName =
                    if (
                        $item.IsCustom -and
                        -not [string]::IsNullOrWhiteSpace(
                            $item.CustomName.Text
                        )
                    ) {
                        $item.CustomName.Text.Trim()
                    }
                    else {
                        $item.Name
                    }

                Show-ErrorMessage `
                    -Message (
                        'Please enter a valid positive quantity for "' +
                        $equipmentName +
                        '" on Day ' +
                        $dayNumber +
                        '.'
                    ) `
                    -Title 'Validation'

                $item.Quantity.Focus()

                return $false
            }
        }

        $dayNumber++
    }

    return $true
}

# ================================================================
# BUILD PAGE 3
# ================================================================

function Build-DeliverablesPage {

    param(
        [System.Windows.Forms.Form]$Form
    )

    if (
        $script:WizardPages.ContainsKey('Page3') -and
        $null -ne $script:WizardPages.Page3
    ) {

        $oldPage =
            $script:WizardPages.Page3

        $Form.Controls.Remove($oldPage)

        $oldPage.Dispose()
    }

    $panel =
        New-Object System.Windows.Forms.Panel

    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true

    $title = New-Label `
        -Text 'Deliverables & Payment' `
        -X 25 `
        -Y 12 `
        -Width 650 `
        -Height 35 `
        -FontSize 19 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $subtitle = New-Label `
        -Text 'Select deliverables and enter payment information.' `
        -X 25 `
        -Y 47 `
        -Width 850 `
        -Height 25 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $panel.Controls.AddRange(
        @(
            $title,
            $subtitle
        )
    )

    # ------------------------------------------------------------
    # DELIVERABLES
    # ------------------------------------------------------------

    $deliverableGroup =
        New-SectionGroup `
            -Text 'Deliverables' `
            -Width 900 `
            -Height 345

    $deliverableGroup.Location =
        New-Object System.Drawing.Point(
            25,
            80
        )

    $script:WizardDeliverableRows = @()

    $rowY = 30

    foreach ($option in $script:DeliverableOptions) {

        $check =
            New-Object System.Windows.Forms.CheckBox

        $check.Location =
            New-Object System.Drawing.Point(
                20,
                ($rowY + 2)
            )

        $check.Size =
            New-Object System.Drawing.Size(
                24,
                24
            )

        $nameLabel = New-Label `
            -Text $option `
            -X 50 `
            -Y $rowY `
            -Width 340 `
            -Height 25 `
            -FontSize 9 `
            -Bold

        $qtyLabel = New-Label `
            -Text 'Qty' `
            -X 400 `
            -Y $rowY `
            -Width 35 `
            -Height 24 `
            -FontSize 8 `
            -Color $script:Theme.Muted

        $quantityBox = New-TextBox `
            -Width 65 `
            -Height 30

        $quantityBox.Location =
            New-Object System.Drawing.Point(
                435,
                ($rowY - 3)
            )

        $quantityBox.TextAlign =
            [System.Windows.Forms.HorizontalAlignment]::Center

        $quantityBox.Enabled = $false

        $customLabel = $null
        $customNameBox = $null

        $isCustom =
            ($option -eq 'Custom')

        if ($isCustom) {

            $customLabel = New-Label `
                -Text 'Custom Name' `
                -X 515 `
                -Y $rowY `
                -Width 90 `
                -Height 24 `
                -FontSize 8 `
                -Color $script:Theme.Muted

            $customNameBox = New-TextBox `
                -Width 250 `
                -Height 30

            $customNameBox.Location =
                New-Object System.Drawing.Point(
                    610,
                    ($rowY - 3)
                )

            $customLabel.Visible = $false
            $customNameBox.Visible = $false

            $deliverableGroup.Controls.AddRange(
                @(
                    $customLabel,
                    $customNameBox
                )
            )
        }

        $row = [PSCustomObject][ordered]@{
            Name        = $option
            Check       = $check
            Quantity    = $quantityBox
            CustomName  = $customNameBox
            CustomLabel = $customLabel
            IsCustom    = $isCustom
        }

        $check.Tag = $row

        $check.Add_CheckedChanged({

            $current =
                $this.Tag

            $current.Quantity.Enabled =
                $current.Check.Checked

            if ($current.IsCustom) {

                $current.CustomName.Visible =
                    $current.Check.Checked

                $current.CustomLabel.Visible =
                    $current.Check.Checked
            }
        })

        $script:WizardDeliverableRows += $row

        $deliverableGroup.Controls.AddRange(
            @(
                $check,
                $nameLabel,
                $qtyLabel,
                $quantityBox
            )
        )

        $rowY += 40
    }

    # Restore selected deliverables
    $existingDeliverables =
        @($script:CurrentBooking.Deliverables)

    foreach ($item in @(
        $script:WizardDeliverableRows
    )) {

        if ($item.IsCustom) {

            $customExisting =
                $existingDeliverables |
                Where-Object {
                    $script:DeliverableOptions -notcontains $_.Name
                } |
                Select-Object -First 1

            if ($null -ne $customExisting) {

                $item.Check.Checked = $true

                $item.Quantity.Text =
                    [string]$customExisting.Quantity

                $item.CustomName.Text =
                    [string]$customExisting.Name
            }

            continue
        }

        $existing =
            $existingDeliverables |
            Where-Object {
                $_.Name -eq $item.Name
            } |
            Select-Object -First 1

        if ($null -ne $existing) {

            $item.Check.Checked = $true

            $item.Quantity.Text =
                [string]$existing.Quantity
        }
    }

    $panel.Controls.Add(
        $deliverableGroup
    )

    # ------------------------------------------------------------
    # PAYMENT
    # ------------------------------------------------------------

    $paymentGroup =
        New-SectionGroup `
            -Text 'Payment' `
            -Width 900 `
            -Height 145

    $paymentGroup.Location =
        New-Object System.Drawing.Point(
            25,
            445
        )

    $totalLabel = New-Label `
        -Text 'Total Charges *' `
        -X 25 `
        -Y 37 `
        -Width 145 `
        -Height 27 `
        -FontSize 10 `
        -Bold

    $totalCharges =
        New-TextBox `
            -Width 220 `
            -Height 32

    $totalCharges.Location =
        New-Object System.Drawing.Point(
            170,
            32
        )

    $advancedLabel = New-Label `
        -Text 'Advanced Received *' `
        -X 465 `
        -Y 37 `
        -Width 155 `
        -Height 27 `
        -FontSize 10 `
        -Bold

    $advancedReceived =
        New-TextBox `
            -Width 220 `
            -Height 32

    $advancedReceived.Location =
        New-Object System.Drawing.Point(
            620,
            32
        )

    $paymentGroup.Controls.AddRange(
        @(
            $totalLabel,
            $totalCharges,
            $advancedLabel,
            $advancedReceived
        )
    )

    $panel.Controls.Add($paymentGroup)

    $totalCharges.Text =
        Format-Amount (
            [decimal]$script:CurrentBooking.TotalCharges
        )

    $advancedReceived.Text =
        Format-Amount (
            [decimal]$script:CurrentBooking.AdvancedReceived
        )

    $script:WizardControls.TotalCharges =
        $totalCharges

    $script:WizardControls.AdvancedReceived =
        $advancedReceived

    $script:WizardPages.Page3 =
        $panel

    $Form.Controls.Add($panel)

    $panel.BringToFront()
}

# ================================================================
# SYNC DELIVERABLES/PAYMENT
# ================================================================

function Sync-DeliverablesToDraft {

    $selected = @()

    foreach ($item in @(
        $script:WizardDeliverableRows
    )) {

        if (-not $item.Check.Checked) {
            continue
        }

        $quantity = 0

        [void](
            Try-GetPositiveInt `
                -Text $item.Quantity.Text `
                -Value ([ref]$quantity)
        )

        if ($item.IsCustom) {
            $name = $item.CustomName.Text.Trim()
        }
        else {
            $name = $item.Name
        }

        $selected +=
            [PSCustomObject][ordered]@{
                Name     = $name
                Quantity = $quantity
            }
    }

    $script:CurrentBooking.Deliverables =
        @($selected)

    $total = [decimal]0
    $advanced = [decimal]0

    if (
        Try-GetDecimalAmount `
            -Text $script:WizardControls.TotalCharges.Text `
            -Value ([ref]$total)
    ) {

        $script:CurrentBooking.TotalCharges =
            $total
    }

    if (
        Try-GetDecimalAmount `
            -Text $script:WizardControls.AdvancedReceived.Text `
            -Value ([ref]$advanced)
    ) {

        $script:CurrentBooking.AdvancedReceived =
            $advanced
    }
}

# ================================================================
# VALIDATE PAGE 3
# ================================================================

function Validate-DeliverablesPaymentPage {

    Sync-DeliverablesToDraft

    foreach ($item in @(
        $script:WizardDeliverableRows
    )) {

        if (-not $item.Check.Checked) {
            continue
        }

        if (
            $item.IsCustom -and
            [string]::IsNullOrWhiteSpace(
                $item.CustomName.Text
            )
        ) {

            Show-ErrorMessage `
                -Message 'Please enter a Custom Deliverable Name.' `
                -Title 'Validation'

            $item.CustomName.Focus()

            return $false
        }

        $quantity = 0

        if (-not (
            Try-GetPositiveInt `
                -Text $item.Quantity.Text `
                -Value ([ref]$quantity)
        )) {

            $displayName =
                if (
                    $item.IsCustom -and
                    -not [string]::IsNullOrWhiteSpace(
                        $item.CustomName.Text
                    )
                ) {
                    $item.CustomName.Text.Trim()
                }
                else {
                    $item.Name
                }

            Show-ErrorMessage `
                -Message (
                    'Please enter a valid positive quantity for "' +
                    $displayName +
                    '".'
                ) `
                -Title 'Validation'

            $item.Quantity.Focus()

            return $false
        }
    }

    $total = [decimal]0

    if (-not (
        Try-GetDecimalAmount `
            -Text $script:WizardControls.TotalCharges.Text `
            -Value ([ref]$total)
    )) {

        Show-ErrorMessage `
            -Message 'Total Charges must be a valid numeric amount.' `
            -Title 'Validation'

        $script:WizardControls.TotalCharges.Focus()

        return $false
    }

    $advanced = [decimal]0

    if (-not (
        Try-GetDecimalAmount `
            -Text $script:WizardControls.AdvancedReceived.Text `
            -Value ([ref]$advanced)
    )) {

        Show-ErrorMessage `
            -Message 'Advanced Received must be a valid numeric amount.' `
            -Title 'Validation'

        $script:WizardControls.AdvancedReceived.Focus()

        return $false
    }

    $script:CurrentBooking.TotalCharges =
        $total

    $script:CurrentBooking.AdvancedReceived =
        $advanced

    return $true
}

# ================================================================
# SUMMARY PAGE
# ================================================================

function Build-SummaryPage {

    param(
        [System.Windows.Forms.Form]$Form
    )

    if (
        $script:WizardPages.ContainsKey('Page4') -and
        $null -ne $script:WizardPages.Page4
    ) {

        $oldPage =
            $script:WizardPages.Page4

        $Form.Controls.Remove($oldPage)

        $oldPage.Dispose()
    }

    $panel =
        New-Object System.Windows.Forms.Panel

    $panel.Dock = 'Fill'

    $successPanel =
        New-Object System.Windows.Forms.Panel

    $successPanel.Location =
        New-Object System.Drawing.Point(
            30,
            18
        )

    $successPanel.Size =
        New-Object System.Drawing.Size(
            900,
            88
        )

    $successPanel.BackColor =
        [System.Drawing.Color]::FromArgb(
            235,
            248,
            240
        )

    $successTitle = New-Label `
        -Text 'Booking Saved Successfully' `
        -X 22 `
        -Y 14 `
        -Width 700 `
        -Height 30 `
        -FontSize 16 `
        -Bold `
        -Color $script:Theme.Success

    $successText = New-Label `
        -Text 'WhatsApp message copied to clipboard.' `
        -X 22 `
        -Y 47 `
        -Width 700 `
        -Height 25 `
        -FontSize 9 `
        -Color $script:Theme.Success

    $successPanel.Controls.AddRange(
        @(
            $successTitle,
            $successText
        )
    )

    $panel.Controls.Add($successPanel)

    $idLabel = New-Label `
        -Text (
            'Booking ID: ' +
            $script:CurrentBooking.BookingId
        ) `
        -X 30 `
        -Y 120 `
        -Width 600 `
        -Height 28 `
        -FontSize 11 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $panel.Controls.Add($idLabel)

    $messageBox =
        New-TextBox `
            -Width 900 `
            -Height 390 `
            -Multiline `
            -ReadOnly

    $messageBox.Location =
        New-Object System.Drawing.Point(
            30,
            155
        )

    $messageBox.Font =
        New-Object System.Drawing.Font(
            'Consolas',
            9.5
        )

    $messageBox.Text =
        $script:CurrentBooking.WhatsAppMessage

    $panel.Controls.Add($messageBox)

    $copyAgain = New-Button `
        -Text 'Copy Message Again' `
        -Width 175 `
        -Height 38 `
        -BackColor $script:Theme.Primary

    $copyAgain.Location =
        New-Object System.Drawing.Point(
            30,
            565
        )

    $copyAgain.Add_Click({

        if (
            Copy-ToClipboard `
                $script:CurrentBooking.WhatsAppMessage
        ) {

            Show-InfoMessage `
                -Message 'WhatsApp message copied to clipboard.' `
                -Title 'Clipboard'
        }
    })

    $viewButton = New-Button `
        -Text 'View Bookings' `
        -Width 145 `
        -Height 38 `
        -BackColor $script:Theme.Accent

    $viewButton.Location =
        New-Object System.Drawing.Point(
            220,
            565
        )

    $viewButton.Add_Click({

        Load-Bookings

        Show-ViewBookings
    })

    $dashboardButton = New-Button `
        -Text 'Back to Dashboard' `
        -Width 160 `
        -Height 38 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $dashboardButton.Location =
        New-Object System.Drawing.Point(
            380,
            565
        )

    $dashboardButton.Add_Click({

        $Form.DialogResult =
            [System.Windows.Forms.DialogResult]::OK

        $Form.Close()
    })

    $panel.Controls.AddRange(
        @(
            $copyAgain,
            $viewButton,
            $dashboardButton
        )
    )

    $script:WizardPages.Page4 =
        $panel

    $Form.Controls.Add($panel)

    $panel.BringToFront()
}

# ================================================================
# COMPLETE SAVE
# ================================================================

function Complete-BookingSave {

    param(
        [System.Windows.Forms.Form]$Form
    )

    $validation =
        Validate-Booking $script:CurrentBooking

    if (-not $validation.Valid) {

        Show-ErrorMessage `
            -Message $validation.Message `
            -Title 'Validation'

        return $false
    }

    # New ID
    if (
        [string]::IsNullOrWhiteSpace(
            [string]$script:CurrentBooking.BookingId
        )
    ) {

        $script:CurrentBooking.BookingId =
            Generate-BookingID

        $script:CurrentBooking.CreatedAt =
            (Get-Date).ToString('o')
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$script:CurrentBooking.CreatedAt
        )
    ) {

        $script:CurrentBooking.CreatedAt =
            (Get-Date).ToString('o')
    }

    $script:CurrentBooking.UpdatedAt =
        (Get-Date).ToString('o')

    $script:CurrentBooking.BookedBy =
        $script:CurrentUser.Name

    $script:CurrentBooking.BookedByContact =
        $script:CurrentUser.Contact

    $script:CurrentBooking.StudioName =
        $script:StudioName

    # Generate final WhatsApp message.
    $script:CurrentBooking.WhatsAppMessage =
        Generate-WhatsAppMessage `
            $script:CurrentBooking

    try {

        $existingIndex = -1

        if ($script:IsEditMode) {

            for (
                $i = 0;
                $i -lt $script:Bookings.Count;
                $i++
            ) {

                if (
                    [string]$script:Bookings[$i].BookingId -eq
                    [string]$script:CurrentBooking.BookingId
                ) {

                    $existingIndex =
                        $i

                    break
                }
            }
        }

        $newCollection = @()

        if ($existingIndex -ge 0) {

            for (
                $i = 0;
                $i -lt $script:Bookings.Count;
                $i++
            ) {

                if ($i -eq $existingIndex) {

                    $newCollection +=
                        Normalize-Booking `
                            $script:CurrentBooking
                }
                else {

                    $newCollection +=
                        Normalize-Booking `
                            $script:Bookings[$i]
                }
            }
        }
        else {

            foreach ($existingBooking in @(
                $script:Bookings
            )) {

                $newCollection +=
                    Normalize-Booking `
                        $existingBooking
            }

            $newCollection +=
                Normalize-Booking `
                    $script:CurrentBooking
        }

        Save-Bookings $newCollection

        $script:Bookings =
            @($newCollection)
    }
    catch {

        Show-ErrorMessage `
            -Message $_.Exception.Message `
            -Title 'Save Error'

        return $false
    }

    [void](
        Copy-ToClipboard `
            $script:CurrentBooking.WhatsAppMessage
    )

    Build-SummaryPage $Form

    Set-WizardStep 4

    return $true
}

# ================================================================
# BOOKING WIZARD
# ================================================================

function Show-AddBooking {

    param(
        $BookingToEdit = $null
    )

    $script:IsEditMode =
        ($null -ne $BookingToEdit)

    if ($script:IsEditMode) {

        $script:CurrentBooking =
            Clone-BookingObject `
                $BookingToEdit
    }
    else {

        $script:CurrentBooking =
            New-BookingDraft
    }

    $script:CurrentBooking.StudioName =
        $script:StudioName

    $form =
        New-Object System.Windows.Forms.Form

    if ($script:IsEditMode) {

        $form.Text =
            'Al Fajar Studio — Edit Booking'
    }
    else {

        $form.Text =
            'Al Fajar Studio — New Booking'
    }

    $form.ClientSize =
        New-Object System.Drawing.Size(
            980,
            710
        )

    Set-FormDefaults $form

    $form.FormBorderStyle =
        [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $script:WizardForm =
        $form

    $script:WizardPages = @{}
    $script:WizardControls = @{}
    $script:WizardEventRows = @()
    $script:WizardSetupRows = @()
    $script:WizardDeliverableRows = @()
    $script:WizardStep = 1

    New-WizardHeader $form

    Build-Page1 $form
    Build-SetupPage $form
    Build-DeliverablesPage $form

    # Footer
    $footer =
        New-Object System.Windows.Forms.Panel

    $footer.Dock = 'Bottom'
    $footer.Height = 66
    $footer.BackColor = $script:Theme.Surface

    $backButton = New-Button `
        -Text 'Back' `
        -Width 105 `
        -Height 38 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $backButton.Location =
        New-Object System.Drawing.Point(
            560,
            14
        )

    $nextButton = New-Button `
        -Text 'Next' `
        -Width 105 `
        -Height 38 `
        -BackColor $script:Theme.Primary

    $nextButton.Location =
        New-Object System.Drawing.Point(
            675,
            14
        )

    $finishButton = New-Button `
        -Text 'Finish' `
        -Width 120 `
        -Height 38 `
        -BackColor $script:Theme.Success

    $finishButton.Location =
        New-Object System.Drawing.Point(
            675,
            14
        )

    $cancelButton = New-Button `
        -Text 'Cancel' `
        -Width 105 `
        -Height 38 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                115,
                122,
                130
            )
        )

    $cancelButton.Location =
        New-Object System.Drawing.Point(
            810,
            14
        )

    $footer.Controls.AddRange(
        @(
            $backButton,
            $nextButton,
            $finishButton,
            $cancelButton
        )
    )

    $form.Controls.Add($footer)

    $script:WizardControls.BackButton =
        $backButton

    $script:WizardControls.NextButton =
        $nextButton

    $script:WizardControls.FinishButton =
        $finishButton

    $script:WizardControls.CancelButton =
        $cancelButton

    # BACK
    $backButton.Add_Click({

        try {

            if ($script:WizardStep -eq 2) {

                Sync-SetupToDraft

                Set-WizardStep 1

                return
            }

            if ($script:WizardStep -eq 3) {

                Sync-DeliverablesToDraft

                Build-SetupPage $form

                Set-WizardStep 2

                return
            }
        }
        catch {

            Show-ErrorMessage `
                -Message $_.Exception.Message `
                -Title 'Navigation Error'
        }
    })

    # NEXT
    $nextButton.Add_Click({

        try {

            if ($script:WizardStep -eq 1) {

                if (-not (Validate-Page1)) {
                    return
                }

                Build-SetupPage $form

                Set-WizardStep 2

                return
            }

            if ($script:WizardStep -eq 2) {

                if (-not (Validate-SetupPage)) {
                    return
                }

                Build-DeliverablesPage $form

                Set-WizardStep 3

                return
            }
        }
        catch {

            Show-ErrorMessage `
                -Message $_.Exception.Message `
                -Title 'Navigation Error'
        }
    })

    # FINISH
    $finishButton.Add_Click({

        try {

            if (-not (
                Validate-DeliverablesPaymentPage
            )) {
                return
            }

            [void](
                Complete-BookingSave `
                    $form
            )
        }
        catch {

            Show-ErrorMessage `
                -Message $_.Exception.Message `
                -Title 'Save Error'
        }
    })

    # CANCEL
    $cancelButton.Add_Click({

        if ($script:WizardStep -eq 4) {

            $form.Close()

            return
        }

        $confirm =
            Show-ConfirmMessage `
                -Message 'Cancel this booking and discard unsaved changes?' `
                -Title 'Cancel Booking'

        if ($confirm) {

            $form.Close()
        }
    })

    Set-WizardStep 1

    [void]$form.ShowDialog()

    $form.Dispose()

    $script:WizardForm = $null
}

# ================================================================
# BOOKING SEARCH
# ================================================================

function Get-FilteredBookings {

    param(
        [string]$SearchText
    )

    $term =
        $SearchText.Trim().ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($term)) {

        return @($script:Bookings)
    }

    $filtered = @()

    foreach ($booking in @(
        $script:Bookings
    )) {

        $name =
            [string]$booking.ClientName

        $phone =
            [string]$booking.ClientPhone

        $id =
            [string]$booking.BookingId

        if (
            $name.ToLowerInvariant().Contains($term) -or
            $phone.ToLowerInvariant().Contains($term) -or
            $id.ToLowerInvariant().Contains($term)
        ) {

            $filtered += $booking
        }
    }

    return @($filtered)
}

function Get-FirstEventDaySummary {

    param(
        $Booking
    )

    $days =
        @($Booking.EventDays)

    if ($days.Count -eq 0) {

        return [PSCustomObject]@{
            Date     = ''
            Event    = ''
            Location = ''
        }
    }

    $day =
        $days[0]

    $dateText =
        [string]$day.Date

    try {

        $dateText =
            ([datetime]::Parse(
                [string]$day.Date
            )).ToString(
                'dd MMM yyyy'
            )
    }
    catch {
    }

    return [PSCustomObject]@{
        Date =
            $dateText

        Event =
            Get-ActiveEventName $day

        Location =
            [string]$day.Location
    }
}

# ================================================================
# BOOKING DETAILS
# ================================================================

function Show-BookingDetails {

    param(
        [Parameter(Mandatory)]
        $Booking
    )

    $form =
        New-Object System.Windows.Forms.Form

    $form.Text =
        'Al Fajar Studio — Booking Details'

    $form.ClientSize =
        New-Object System.Drawing.Size(
            920,
            720
        )

    Set-FormDefaults $form

    $header =
        New-Object System.Windows.Forms.Panel

    $header.Dock = 'Top'
    $header.Height = 80
    $header.BackColor = $script:Theme.Surface

    $title = New-Label `
        -Text 'Booking Details' `
        -X 22 `
        -Y 15 `
        -Width 500 `
        -Height 30 `
        -FontSize 18 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $id = New-Label `
        -Text ([string]$Booking.BookingId) `
        -X 22 `
        -Y 46 `
        -Width 350 `
        -Height 24 `
        -FontSize 9 `
        -Bold `
        -Color $script:Theme.Primary

    $header.Controls.AddRange(
        @(
            $title,
            $id
        )
    )

    $form.Controls.Add($header)

    $detailBox =
        New-TextBox `
            -Width 870 `
            -Height 565 `
            -Multiline `
            -ReadOnly

    $detailBox.Location =
        New-Object System.Drawing.Point(
            25,
            95
        )

    $detailBox.Font =
        New-Object System.Drawing.Font(
            'Consolas',
            9.5
        )

    $lines =
        New-Object System.Collections.Generic.List[string]

    [void]$lines.Add('CLIENT INFORMATION')
    [void]$lines.Add('===================')

    [void]$lines.Add(
        'Client Name        : ' +
        [string]$Booking.ClientName
    )

    [void]$lines.Add(
        'Client Phone       : ' +
        [string]$Booking.ClientPhone
    )

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Booking.ClientPhone2
        )
    ) {

        [void]$lines.Add(
            'Client Phone 2     : Not provided'
        )
    }
    else {

        [void]$lines.Add(
            'Client Phone 2     : ' +
            [string]$Booking.ClientPhone2
        )
    }

    [void]$lines.Add('')
    [void]$lines.Add('EVENT INFORMATION')
    [void]$lines.Add('=================')

    $dayNumber = 1

    foreach ($day in @(
        $Booking.EventDays
    )) {

        $dateText =
            [string]$day.Date

        try {

            $dateText =
                ([datetime]::Parse(
                    [string]$day.Date
                )).ToString(
                    'dd MMMM yyyy'
                )
        }
        catch {
        }

        [void]$lines.Add(
            'Day ' +
            $dayNumber +
            ' — ' +
            (Get-ActiveEventName $day)
        )

        [void]$lines.Add(
            '  Date      : ' +
            $dateText
        )

        [void]$lines.Add(
            '  Timing    : ' +
            (Get-ActiveTiming $day)
        )

        [void]$lines.Add(
            '  Location  : ' +
            [string]$day.Location
        )

        [void]$lines.Add(
            '  Equipment :'
        )

        $equipmentList =
            @($day.Equipment)

        if ($equipmentList.Count -eq 0) {

            [void]$lines.Add(
                '    - None selected'
            )
        }
        else {

            foreach ($equipment in $equipmentList) {

                [void]$lines.Add(
                    '    - ' +
                    [string]$equipment.Name +
                    ' (Qty: ' +
                    [string]$equipment.Quantity +
                    ')'
                )
            }
        }

        [void]$lines.Add('')

        $dayNumber++
    }

    [void]$lines.Add('DELIVERABLES')
    [void]$lines.Add('============')

    $deliverables =
        @($Booking.Deliverables)

    if ($deliverables.Count -eq 0) {

        [void]$lines.Add(
            '- None selected'
        )
    }
    else {

        foreach ($deliverable in $deliverables) {

            [void]$lines.Add(
                '- ' +
                [string]$deliverable.Name +
                ' (Qty: ' +
                [string]$deliverable.Quantity +
                ')'
            )
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add('PAYMENT')
    [void]$lines.Add('=======')

    [void]$lines.Add(
        'Total Charges      : Rs. ' +
        (Format-Amount (
            [decimal]$Booking.TotalCharges
        ))
    )

    [void]$lines.Add(
        'Advanced Received  : Rs. ' +
        (Format-Amount (
            [decimal]$Booking.AdvancedReceived
        ))
    )

    [void]$lines.Add('')
    [void]$lines.Add('BOOKING INFORMATION')
    [void]$lines.Add('===================')

    [void]$lines.Add(
        'Booking ID         : ' +
        [string]$Booking.BookingId
    )

    [void]$lines.Add(
        'Booked By          : ' +
        [string]$Booking.BookedBy
    )

    [void]$lines.Add(
        'Booked By Contact  : ' +
        [string]$Booking.BookedByContact
    )

    [void]$lines.Add(
        'Studio Name        : ' +
        $script:StudioName
    )

    $detailBox.Text =
        $lines -join [Environment]::NewLine

    $form.Controls.Add($detailBox)

    # Copy
    $copyButton = New-Button `
        -Text 'Copy WhatsApp Message' `
        -Width 185 `
        -Height 38 `
        -BackColor $script:Theme.Primary

    $copyButton.Location =
        New-Object System.Drawing.Point(
            25,
            665
        )

    $copyButton.Add_Click({

        if (
            Copy-ToClipboard `
                $Booking.WhatsAppMessage
        ) {

            Show-InfoMessage `
                -Message 'WhatsApp message copied to clipboard.' `
                -Title 'Clipboard'
        }
    })

    # Edit
    $editButton = New-Button `
        -Text 'Edit Booking' `
        -Width 135 `
        -Height 38 `
        -BackColor $script:Theme.Accent

    $editButton.Location =
        New-Object System.Drawing.Point(
            225,
            665
        )

    $editButton.Add_Click({

        $form.Close()

        Show-AddBooking `
            -BookingToEdit $Booking
    })

    # Close
    $closeButton = New-Button `
        -Text 'Close' `
        -Width 105 `
        -Height 38 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $closeButton.Location =
        New-Object System.Drawing.Point(
            790,
            665
        )

    $closeButton.Add_Click({
        $form.Close()
    })

    $form.Controls.AddRange(
        @(
            $copyButton,
            $editButton,
            $closeButton
        )
    )

    [void]$form.ShowDialog()

    $form.Dispose()
}

# ================================================================
# VIEW BOOKINGS
# ================================================================

function Show-ViewBookings {

    Load-Bookings

    $form =
        New-Object System.Windows.Forms.Form

    $form.Text =
        'Al Fajar Studio — View Bookings'

    $form.ClientSize =
        New-Object System.Drawing.Size(
            1080,
            680
        )

    Set-FormDefaults $form

    # Header
    $header =
        New-Object System.Windows.Forms.Panel

    $header.Dock = 'Top'
    $header.Height = 92
    $header.BackColor = $script:Theme.Surface

    $title = New-Label `
        -Text 'View Booking' `
        -X 25 `
        -Y 15 `
        -Width 400 `
        -Height 32 `
        -FontSize 18 `
        -Bold `
        -Color $script:Theme.PrimaryDark

    $subtitle = New-Label `
        -Text 'Search by Client Name, Client Phone Number, or Booking ID.' `
        -X 25 `
        -Y 48 `
        -Width 700 `
        -Height 24 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $header.Controls.AddRange(
        @(
            $title,
            $subtitle
        )
    )

    $form.Controls.Add($header)

    # Search
    $searchLabel = New-Label `
        -Text 'Search' `
        -X 25 `
        -Y 110 `
        -Width 65 `
        -Height 27 `
        -FontSize 10 `
        -Bold

    $searchBox =
        New-TextBox `
            -Width 445 `
            -Height 32

    $searchBox.Location =
        New-Object System.Drawing.Point(
            90,
            105
        )

    $clearButton = New-Button `
        -Text 'Clear' `
        -Width 90 `
        -Height 32 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $clearButton.Location =
        New-Object System.Drawing.Point(
            550,
            105
        )

    $form.Controls.AddRange(
        @(
            $searchLabel,
            $searchBox,
            $clearButton
        )
    )

    # Grid
    $grid =
        New-Object System.Windows.Forms.DataGridView

    $grid.Location =
        New-Object System.Drawing.Point(
            25,
            150
        )

    $grid.Size =
        New-Object System.Drawing.Size(
            1030,
            425
        )

    $grid.BackgroundColor =
        [System.Drawing.Color]::White

    $grid.BorderStyle =
        [System.Windows.Forms.BorderStyle]::FixedSingle

    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true

    $grid.SelectionMode =
        [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

    $grid.MultiSelect = $false
    $grid.AutoGenerateColumns = $false
    $grid.RowHeadersVisible = $false

    $grid.ColumnHeadersHeight = 38

    $grid.DefaultCellStyle.Font =
        New-Object System.Drawing.Font(
            'Segoe UI',
            9
        )

    $grid.ColumnHeadersDefaultCellStyle.Font =
        New-Object System.Drawing.Font(
            'Segoe UI Semibold',
            9
        )

    $columnDefinitions = @(
        @{
            Header = 'Booking ID'
            Width  = 145
            Name   = 'BookingId'
        },
        @{
            Header = 'Client Name'
            Width  = 210
            Name   = 'ClientName'
        },
        @{
            Header = 'Client Phone'
            Width  = 145
            Name   = 'ClientPhone'
        },
        @{
            Header = 'Event Date'
            Width  = 120
            Name   = 'EventDate'
        },
        @{
            Header = 'Event'
            Width  = 125
            Name   = 'Event'
        },
        @{
            Header = 'Location'
            Width  = 210
            Name   = 'Location'
        }
    )

    foreach ($definition in $columnDefinitions) {

        $column =
            New-Object System.Windows.Forms.DataGridViewTextBoxColumn

        $column.HeaderText =
            $definition.Header

        $column.Width =
            $definition.Width

        $column.Name =
            $definition.Name

        [void]$grid.Columns.Add(
            $column
        )
    }

    $form.Controls.Add($grid)

    # Buttons
    $closeButton = New-Button `
        -Text 'Close' `
        -Width 105 `
        -Height 38 `
        -BackColor (
            [System.Drawing.Color]::FromArgb(
                130,
                138,
                148
            )
        )

    $closeButton.Location =
        New-Object System.Drawing.Point(
            25,
            600
        )

    $detailsButton = New-Button `
        -Text 'Open Details' `
        -Width 135 `
        -Height 38 `
        -BackColor $script:Theme.Primary

    $detailsButton.Location =
        New-Object System.Drawing.Point(
            750,
            600
        )

    $editButton = New-Button `
        -Text 'Edit Booking' `
        -Width 135 `
        -Height 38 `
        -BackColor $script:Theme.Accent

    $editButton.Location =
        New-Object System.Drawing.Point(
            900,
            600
        )

    $status = New-Label `
        -Text '' `
        -X 165 `
        -Y 604 `
        -Width 520 `
        -Height 28 `
        -FontSize 9 `
        -Color $script:Theme.Muted

    $form.Controls.AddRange(
        @(
            $closeButton,
            $detailsButton,
            $editButton,
            $status
        )
    )

    $refreshGrid = {

        $grid.Rows.Clear()

        $filtered =
            Get-FilteredBookings `
                -SearchText $searchBox.Text

        foreach ($booking in @(
            $filtered
        )) {

            $summary =
                Get-FirstEventDaySummary `
                    $booking

            $values = [object[]]@(
                $booking.BookingId,
                $booking.ClientName,
                $booking.ClientPhone,
                $summary.Date,
                $summary.Event,
                $summary.Location
            )

            $rowIndex =
                $grid.Rows.Add(
                    $values
                )

            $grid.Rows[$rowIndex].Tag =
                $booking
        }

        $status.Text =
            'Showing ' +
            $filtered.Count +
            ' booking(s)'
    }

    $searchBox.Add_TextChanged({
        & $refreshGrid
    })

    $clearButton.Add_Click({
        $searchBox.Clear()
    })

    $detailsButton.Add_Click({

        if ($grid.SelectedRows.Count -eq 0) {

            Show-InfoMessage `
                -Message 'Please select a booking first.' `
                -Title 'View Booking'

            return
        }

        $booking =
            $grid.SelectedRows[0].Tag

        Show-BookingDetails $booking

        Load-Bookings

        & $refreshGrid
    })

    $editButton.Add_Click({

        if ($grid.SelectedRows.Count -eq 0) {

            Show-InfoMessage `
                -Message 'Please select a booking first.' `
                -Title 'View Booking'

            return
        }

        $booking =
            $grid.SelectedRows[0].Tag

        $form.Hide()

        Show-AddBooking `
            -BookingToEdit $booking

        Load-Bookings

        if (-not $form.IsDisposed) {

            $form.Show()

            & $refreshGrid
        }
    })

    $grid.Add_DoubleClick({

        if ($grid.SelectedRows.Count -eq 0) {
            return
        }

        $booking =
            $grid.SelectedRows[0].Tag

        Show-BookingDetails $booking

        Load-Bookings

        & $refreshGrid
    })

    $closeButton.Add_Click({
        $form.Close()
    })

    & $refreshGrid

    [void]$form.ShowDialog()

    $form.Dispose()
}

# ================================================================
# START APPLICATION
# ================================================================

try {

    Load-Bookings

    if ($script:DataLoadError) {

        Show-ErrorMessage `
            -Message $script:DataLoadError `
            -Title 'Booking Database Error'
    }

    # The first visible screen is the password screen.
    $loggedIn = Show-Login

    if ($loggedIn) {

        Show-Dashboard
    }
}
catch {

    Show-ErrorMessage `
        -Message (
            "The application encountered an unexpected error:`r`n`r`n" +
            $_.Exception.Message
        ) `
        -Title 'Al Fajar Studio — Error'
}