#New-ManagementRoleAssignment -Name "AllowImpersonation" -Role ApplicationImpersonation -User sanglyb@yourorg.com
#New-RoleGroup -name ImpersonationGroup -Roles ApplicationImpersonation

function Get-AllItemsFromFolder {
    param(
        [Microsoft.Exchange.WebServices.Data.FolderId]   $FolderId,
        [Microsoft.Exchange.WebServices.Data.SearchFilter] $Filter = $null,
        [DateTime] $StartDate = (Get-Date).AddDays(-30),
        [DateTime] $EndDate   = (Get-Date).AddDays(30)
    )

    $items = @()
    $offset = 0
    $pageSize = 100
    $moreItems = $true
	$currentStart = $StartDate
	
    if ($FolderId.FolderName -eq [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar) {
		$interval = 14 
		while ($currentStart -lt $EndDate) {
			$currentEnd = $currentStart.AddDays($interval)
			if ($currentEnd -gt $EndDate) { $currentEnd = $EndDate }
				write-host $currentStart $currentend
				$calendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($currentStart, $currentEnd, $pageSize)
				$find = $service.FindAppointments($FolderId, $calendarView)
			if ($find.Items.Count -gt 0) {
				$service.LoadPropertiesForItems($find.Items, $propertySet)
				$items += $find.Items
			}
			$currentStart = $currentEnd
			Start-Sleep -Milliseconds 100
		}
	}	
    else {        
        do {
            if ($Filter) {
                $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset)
                $find = $service.FindItems($FolderId, $Filter, $view)
            } else {
                $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset)
                $find = $service.FindItems($FolderId, $view)
            }

            if ($find.Items.Count -gt 0) {
                $service.LoadPropertiesForItems($find.Items, $propertySet)
                $items += $find.Items
            }
            $moreItems = $find.MoreAvailable
            if ($moreItems) {
                $offset += $pageSize
                Start-Sleep -Milliseconds 200
            }
        } while ($moreItems)
    }

    return $items
}

function Create-MeetingObject {
    param(
        [string]$Type,
        [object]$Item,
		[regex]$regex
    )
    $attendees = @($Item.RequiredAttendees + $Item.OptionalAttendees)
	if ($null -ne $item.JoinOnlineMeetingUrl){
		$JoinOnlineMeetingUrl=$item.JoinOnlineMeetingUrl
		$MeetingID=$item.JoinOnlineMeetingUrl.split("/")[-1]
	} elseif ($null -ne $regex.Matches($item.body.text)) {
		$JoinOnlineMeetingUrl=($regex.Matches($item.body.text))[0].value
		if ($null -ne $JoinOnlineMeetingUrl) {
			$MeetingID = $JoinOnlineMeetingUrl -split '/' | Select-Object -Last 1
		}
	}
	if (($item.Body.Text -replace '(?s)\.{20,}.*?(Присоединиться к собранию Skype|Skype Web App).*?\.{20,}', '').trim().length -gt 0){
		$HasBody=$true
	} else {
		$HasBody=$false
	}
    return [pscustomobject]@{
        UserName           = $username
        Title              = $title
        TargetMailbox      = $targetMailbox
        ItemType           = $Type
        Subject            = $Item.Subject
        Start              = $Item.Start 
        End                = $Item.End 
        Duration           = $Item.Duration
        Organizer          = $Item.Organizer.Address
        DateTimeSent       = $Item.DateTimeSent
        DateTimeReceived   = $Item.DateTimeReceived
        ICalDateTimeStamp  = $Item.ICalDateTimeStamp
        HasAttachments     = $Item.HasAttachments
        HasBody            = $HasBody
		JoinOnlineMeetingURL = $JoinOnlineMeetingUrl
		MeetingID = $MeetingID
#        Body               = $Item.Body.Text
        AttendeesCount     = $attendees.Count
        RequiredAttendees  = ($Item.RequiredAttendees | ForEach-Object { $_.Address }) -join '; '
        OptionalAttendees  = ($Item.OptionalAttendees | ForEach-Object { $_.Address }) -join '; '
    }
}

$users=Get-Content "c:\scripts\emails_list.txt"
$results = @()
foreach ($user in $users){
if (Get-Mailbox $user -erroraction silentlycontinue) { 

# $targetMailbox = "nmel@yourorg.com"
$targetMailbox = $user


$startDate = ((get-date).AddDays(-35))
$endDate = ((get-date).adddays(1))
$regex = [regex]'https://meet\.yourorg\.com/[^\s\?<>]+'
$server        = "172.21.38.152"
$pageSize      = 100

								

#обход ошибок сертификата
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
    param($sender, $certificate, $chain, $sslPolicyErrors)
    return $true
}

#user info
$samName=(get-mailbox $targetMailbox -erroraction silentlycontinue).samaccountname
if ($samName){
	$user=get-aduser $samName -properties title
    if ($null -ne $user.name){
		$username=$user.name
	} else {
		$username=$targetMailbox
	}
	$title=$user.title
}

#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
)
$service.UseDefaultCredentials = $true

#$cred=Get-Credential
#$service.UseDefaultCredentials = $false
#$service.Credentials = New-Object System.Net.NetworkCredential(
#    $cred.UserName,
#    $cred.GetNetworkCredential().Password
#)

$service.Url = "https://$server/EWS/Exchange.asmx"
#impersonation
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    $targetMailbox
)

$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$mailbox    = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($targetMailbox)
$calendarId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
    [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,
    $mailbox
)

#filters
$filterClass = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
    "IPM.Schedule.Meeting.Request"
)

$filterStart = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo(
[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,
$startDate
)

$filterEnd = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo(
[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,
$endDate
)

$calendarFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
    [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And,
    @($filterStart, $filterEnd)
)

$meetingFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
    [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And,
    @($filterClass, $filterStart, $filterEnd)
)


#allCalendaritems
$calendarItems = Get-AllItemsFromFolder -FolderId $calendarId -filter $calendarFilter -StartDate $startDate -EndDate $endDate |
    Where-Object { $_ -is [Microsoft.Exchange.WebServices.Data.Item] }

#allFolders
$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
    [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,
    $mailbox
)
$allFolders = $service.FindFolders($rootId, $folderView).Folders

#allInvitations
$inviteItems = @()
foreach ($folder in $allFolders) {
    try {
        $items = Get-AllItemsFromFolder -FolderId $folder.Id -Filter $meetingFilter -StartDate $startDate -EndDate $endDate
        $inviteItems += $items | Where-Object { $_ -is [Microsoft.Exchange.WebServices.Data.Item] }
    } catch {
        Write-Warning "Ошибка при обходе папки '$($folder.DisplayName)': $_"
    }
}


$calendarUids = $calendarItems | ForEach-Object { $_.ICalUid }
$inviteItems  = $inviteItems | Where-Object { $calendarUids -notcontains $_.ICalUid } | Group-Object -Property ICalUid | ForEach-Object { $_.Group[0] }


$results += $calendarItems | ForEach-Object { Create-MeetingObject -Type 'Appointment' -Item $_ -regex $regex}
$results += $inviteItems   | ForEach-Object { Create-MeetingObject -Type 'Invitation' -Item $_ -regex $regex}
} else {
write-host "$user mailbox not found"
}
}

#$results
$results | Export-Csv -Path "meetings.csv"  -NoTypeInformation -Encoding UTF8
