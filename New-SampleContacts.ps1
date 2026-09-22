<#
.SYNOPSIS
    Creates sample Outlook contacts (and personal distribution lists) from a CSV of fictitious names.

.DESCRIPTION
    Uses the Outlook Object Model (COM automation) to create a contact in the default Contacts
    folder of the currently configured Outlook profile for every row in the supplied CSV file.
    Each contact is given an SMTP email address in the format "firstname@<EmailDomain>", where
    "firstname" is derived from the CSV's FirstName column (trimmed, diacritics/punctuation
    stripped, lower-cased) and duplicates are disambiguated with a numeric suffix.

    Once all contacts have been created, the script creates a number of personal distribution
    lists (Outlook DistListItem objects), each populated with a random number of the contacts
    that were just created.

.PARAMETER CsvPath
    Path to the CSV file containing the sample names. Defaults to "Fictitious Names.csv" in the
    same folder as this script. The CSV must contain at least FirstName, LastName and FullName
    columns.

.PARAMETER EmailDomain
    The domain to use when building each contact's email address (e.g. "firstname@contoso.com").

.PARAMETER DistributionListCount
    The number of personal distribution lists to create.

.PARAMETER MinMembersPerList
    The minimum number of contacts to add to each distribution list.

.PARAMETER MaxMembersPerList
    The maximum number of contacts to add to each distribution list.

.PARAMETER DistributionListNamePrefix
    Prefix used when naming each generated distribution list (a sequential number is appended).

.EXAMPLE
    .\New-SampleContacts.ps1

    Creates contacts from ".\Fictitious Names.csv" using "contoso.com" as the email domain, then
    creates 10 personal distribution lists with between 4 and 25 members each.

.EXAMPLE
    .\New-SampleContacts.ps1 -DistributionListCount 5 -MinMembersPerList 2 -MaxMembersPerList 10

    Creates the sample contacts, then creates 5 distribution lists, each with 2-10 members.

.NOTES
    Requires a desktop version of Outlook to be installed and configured with a profile.
    Outlook must be able to run/automate on this machine (it will be started if not already running).
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$CsvPath = (Join-Path $PSScriptRoot "Fictitious Names.csv"),

    [string]$EmailDomain = "contoso.com",

    [ValidateRange(0, [int]::MaxValue)]
    [int]$DistributionListCount = 10,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MinMembersPerList = 4,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxMembersPerList = 25,

    [string]$DistributionListNamePrefix = "Sample Distribution List"
)

# Outlook Object Model constants (avoids a dependency on the Microsoft.Office.Interop.Outlook assembly).
$olFolderContacts = 10
$olContactItem = 2
$olDistributionListItem = 7

if ($MinMembersPerList -gt $MaxMembersPerList) {
    throw "MinMembersPerList ($MinMembersPerList) cannot be greater than MaxMembersPerList ($MaxMembersPerList)."
}

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

<#
.SYNOPSIS
    Converts a display name into a lower-case, ASCII-only, whitespace-free string suitable for
    use as the local part of an email address.
#>
function ConvertTo-EmailLocalPart {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $trimmed = $Name.Trim()
    $normalized = $trimmed.Normalize([System.Text.NormalizationForm]::FormD)

    $builder = New-Object System.Text.StringBuilder
    foreach ($character in $normalized.ToCharArray()) {
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($character)
        }
    }

    $ascii = $builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)
    $ascii = ($ascii -replace '[^a-zA-Z]', '').ToLowerInvariant()

    return $ascii
}

<#
.SYNOPSIS
    Builds a unique SMTP email address for a contact, given their first name.
#>
function New-UniqueEmailAddress {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FirstName,

        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$UsedLocalParts
    )

    $localPart = ConvertTo-EmailLocalPart -Name $FirstName
    if ([string]::IsNullOrEmpty($localPart)) {
        $localPart = "contact"
    }

    $candidate = $localPart
    $suffix = 1
    while ($UsedLocalParts.Contains($candidate)) {
        $suffix++
        $candidate = "$localPart$suffix"
    }

    [void]$UsedLocalParts.Add($candidate)
    return "$candidate@$Domain"
}

Write-Verbose "Importing sample names from '$CsvPath'..."
$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows -or $rows.Count -eq 0) {
    throw "No rows found in CSV file: $CsvPath"
}
Write-Verbose "Found $($rows.Count) sample name(s)."

Write-Verbose "Connecting to Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$contactsFolder = $namespace.GetDefaultFolder($olFolderContacts)

$usedLocalParts = New-Object 'System.Collections.Generic.HashSet[string]'
$createdContacts = New-Object System.Collections.Generic.List[object]

$rowIndex = 0
foreach ($row in $rows) {
    $rowIndex++

    $firstName = $(if ($row.FirstName) { $row.FirstName } else { "" }).Trim()
    $lastName = $(if ($row.LastName) { $row.LastName } else { "" }).Trim()
    $fullName = $(if ($row.FullName) { $row.FullName } else { "" }).Trim()

    if ([string]::IsNullOrWhiteSpace($firstName) -and [string]::IsNullOrWhiteSpace($lastName)) {
        Write-Warning "Skipping row $rowIndex - no FirstName or LastName present."
        continue
    }
    if ([string]::IsNullOrWhiteSpace($fullName)) {
        $fullName = "$firstName $lastName".Trim()
    }

    $emailAddress = New-UniqueEmailAddress -FirstName $firstName -Domain $EmailDomain -UsedLocalParts $usedLocalParts

    if ($PSCmdlet.ShouldProcess($fullName, "Create Outlook contact ($emailAddress)")) {
        $contact = $outlook.CreateItem($olContactItem)
        $contact.FirstName = $firstName
        $contact.LastName = $lastName
        $contact.FullName = $fullName
        $contact.Email1Address = $emailAddress
        $contact.Email1AddressType = "SMTP"
        $contact.Email1DisplayName = "$fullName ($emailAddress)"
        $contact.Save()

        Write-Verbose "Created contact '$fullName' <$emailAddress>."

        $createdContacts.Add([PSCustomObject]@{
            Contact  = $contact
            FullName = $fullName
            Email    = $emailAddress
        })
    }
}

Write-Host "Created $($createdContacts.Count) contact(s) in the '$($contactsFolder.Name)' folder."

if ($createdContacts.Count -eq 0) {
    Write-Warning "No contacts were created, so no distribution lists will be generated."
    return
}

if ($createdContacts.Count -lt $MinMembersPerList) {
    Write-Warning "Only $($createdContacts.Count) contact(s) were created, which is fewer than MinMembersPerList ($MinMembersPerList). Distribution lists will contain at most $($createdContacts.Count) member(s)."
}

$effectiveMaxMembers = [Math]::Min($MaxMembersPerList, $createdContacts.Count)
$effectiveMinMembers = [Math]::Min($MinMembersPerList, $effectiveMaxMembers)

for ($listNumber = 1; $listNumber -le $DistributionListCount; $listNumber++) {
    $listName = "$DistributionListNamePrefix $listNumber"
    $memberCount = Get-Random -Minimum $effectiveMinMembers -Maximum ($effectiveMaxMembers + 1)
    $members = $createdContacts | Get-Random -Count $memberCount

    if (-not $PSCmdlet.ShouldProcess($listName, "Create Outlook distribution list with $memberCount member(s)")) {
        continue
    }

    $distributionList = $outlook.CreateItem($olDistributionListItem)
    $distributionList.DLName = $listName

    $addedCount = 0
    foreach ($member in $members) {
        $recipient = $namespace.CreateRecipient($member.Email)
        [void]$recipient.Resolve()

        if ($recipient.Resolved) {
            $distributionList.AddMember($recipient)
            $addedCount++
        } else {
            Write-Warning "Could not resolve recipient '$($member.Email)' - skipping for list '$listName'."
        }
    }

    $distributionList.Save()
    Write-Verbose "Created distribution list '$listName' with $addedCount member(s)."
}

Write-Host "Created $DistributionListCount distribution list(s) with between $effectiveMinMembers and $effectiveMaxMembers member(s) each."

[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($contactsFolder)
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace)
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
