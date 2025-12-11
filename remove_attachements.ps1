<#
    calendar_cleaner.ps1 — REPORT всегда, CLEAN — опционально
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Mailbox,
    [switch]$Report=$false,
    [int]$DaysBack = 60,
#    [string]$Output = "past_meetings.csv",
    [string]$EwsUrl="https://exchangeserver/EWS/Exchange.asmx"
)

Write-Host "> Mailbox: $Mailbox" -ForegroundColor Cyan
Write-Host "> Report:  $Report"   -ForegroundColor Cyan
Write-Host "> EWS URL: $EwsUrl"   -ForegroundColor Cyan

#обход ошибок сертификата
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
    param($sender, $certificate, $chain, $sslPolicyErrors)
    return $true
}


function New-EwsService {
    param(
        [string]$ImpersonateMailbox,
        [string]$Url
    )
    $svc = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2019
    )

    $svc.UseDefaultCredentials = $true
    $svc.Url = $Url

    if ($ImpersonateMailbox) {
        $svc.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
            [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
            $ImpersonateMailbox
        )
    }
    return $svc
}

function Get-Appointments {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory = $true)]
        [DateTime]$Start,

        [Parameter(Mandatory = $true)]
        [DateTime]$End
    )

    # Разбиваем диапазон по 30 дней
    $step = 30
    $calendarFolders = Get-CalendarFolders -Service $Service
    $allAppointments = @()

    foreach ($folder in $calendarFolders) {

        $chunkStart = $Start

        while ($chunkStart -lt $End) {

            $chunkEnd = $chunkStart.AddDays($step)
            if ($chunkEnd -gt $End) { $chunkEnd = $End }

            # CalendarView на 30 дней
            $view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView(
                $chunkStart, $chunkEnd, 1000
            )
            $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties

            try {
                $page = $Service.FindAppointments($folder.Id, $view)
                if ($page.Items.Count -gt 0) {
                    $allAppointments += $page.Items
                }
            }
            catch {
                Write-Warning "Ошибка чтения папки $($folder.DisplayName): $_"
            }

            # Следующий интервал
            $chunkStart = $chunkEnd
        }
    }

    return $allAppointments
}

<#
function Get-Appointments {
    param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [DateTime]$Start,
        [DateTime]$End
    )
	$calendarFolders = get-calendarFolders $service
	$view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, 500)
	$view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties
	$appointments = @()
		foreach ($folder in $calendarFolders) {
		try {
			$items = $service.FindAppointments($folder.Id, $view)
			if ($items.Items.Count -gt 0) {
				$appointments += $items.Items
			}
		}
		catch {
			Write-Warning "Ошибка при чтении из папки $($folder.DisplayName): $_"
		}
	}
	return $appointments
}
#>

function Get-OriginalAppointment {
    param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$ICalUid,
        [DateTime]$Start,
        [DateTime]$End 
    )
	$appointments=Get-Appointments -service $service -start $start -end $end
	
    foreach ($appt in $appointments) {
        try { $appt.Load() } catch {}
        if ($appt.ICalUid -eq $ICalUid) {
            return $appt
        }
    }

    return $null
}

function get-calendarFolders {
    param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service        
    )
	$root = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot

	$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000)
	$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

	# Фильтр по типу папки
	$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
		[Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass,
		"IPF.Appointment"
	)
	return $service.FindFolders($root, $searchFilter, $folderView)	
}

function Get-Master {
    param(
        [Microsoft.Exchange.WebServices.Data.Appointment]$OrgItem,
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$OrgService
    )

    if ($OrgItem.AppointmentType -eq "Occurrence") {
        return [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster(
            $OrgService,
            $OrgItem.Id
        )
    }

    return $OrgItem
}

function Save-OrgItemAndSendUpdate {
    param(
        [Microsoft.Exchange.WebServices.Data.Appointment]$OrgItem,
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$OrgService
    )

    Write-Host "`nУдаление вложений из события $($OrgItem.Subject)..." -ForegroundColor green

    try {
        $OrgItem.Attachments.Clear()

        $OrgItem.Update(
            [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
            [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone
        )

        $OrgItem.Load()
		#Изменение тела собрания, для коррекной рассылки уведомлениий об изменении
		
        #$OrgItem.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody (
        #    $OrgItem.Body.Text + "<p>Attachement removed by script</p>"
        #)
		
		$body = $OrgItem.Body.Text
		$body = if ($body.EndsWith(" ")) {
			$body.TrimEnd()
		} else {
			$body + " "
		}
		$OrgItem.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody $body


        $OrgItem.Update(
            [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
            [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy
        )

        Write-Host "Обновление события выполнено" -ForegroundColor Green

    } catch {
        Write-Warning "Ошибка обнолвения события: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Warning "Inner: $($_.Exception.InnerException.Message)"
        }
    }
}

$service  = New-EwsService -ImpersonateMailbox $Mailbox -Url $EwsUrl

$startDate = (Get-Date).AddDays(-$DaysBack)
$endDate   = (Get-Date).AddDays(1)
$appointments=Get-Appointments -service $service -start $startDate -end $endDate

$groups = $appointments | where-object {$_.HasAttachments -eq $true} | Group-Object ICalUid,AppointmentType
#$groups.group | ft *type*
#report

Write-Host "`nПодготовка данных..." -ForegroundColor Green
$allItems=@()
$results = foreach ($grp in $groups) {

    $item = $grp.Group | Select-Object -First 1
    $item.Load()
    if ($item.End -ge (Get-Date)) { continue }
    if (-not $item.HasAttachments) { continue }
	if ($item.AppointmentType -eq "RecurringMaster" -and $item.LastOccurrence.End -ge (get-date)) { continue }
    if ($item.AppointmentType -eq "Occurrence") {		
		if ($item.Attachments.Count -ge 1){
			$itemTMP = Get-Master $item $service
			if ($itemTMP.AppointmentType -eq "RecurringMaster" -and $itemTMP.LastOccurrence.End -ge (get-date)) { 
				$test="skipped"
				continue 
			} else {
				$test="not skipped"
				$item=$itemTMP
			}
		}
    }
	$allItems+=$item
    [pscustomobject]@{
        Mailbox          = $Mailbox
        Subject          = $item.Subject
        Start            = $item.Start
        End              = $item.End
        Organizer        = $item.Organizer.Address
        HasAttachments   = $item.HasAttachments
        AttachmentsCount = $item.Attachments.Count
		Attachments		 = $item.Attachments
        ItemType         = $item.AppointmentType
        ICalUid          = $item.ICalUid
		IsCancelled		 = $item.IsCancelled
		LastOccurrence	 = $item.LastOccurrence.End
		skipped			 = $skipped
    }
	
}

$allItems = $allItems | sort-object {$_.Organizer.Address}

if ($allItems.count -eq 0){
	Write-host "Не найдено событий календарей с вложениями." -ForegroundColor green
	exit
} else {
	write-host "Найдены старые события в календаре с вложениями" -ForegroundColor green
	$results 
	#$results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $Output
	#Write-Host "Отчёт сохранён в $Output" -ForegroundColor Green

}

# Если указан Report, только отчёт
if ($Report) {
    Write-Host "`nРежим Report. Удаление вложений не производится." -ForegroundColor Cyan
    exit
}


#Clean
$item=$null
	
	foreach ($item in $allItems){
		$organizer = $item.Organizer.Address
			# --- если организатор другой человек ---
		if ($organizer -and $organizer -ne $Mailbox) {
				try {
					if ($prevorganizer -ne $organizer) {
						Write-Host "Имперсонация $organizer" -ForegroundColor Yellow
						$orgService = New-EwsService -ImpersonateMailbox $organizer -Url $EwsUrl
						$prevorganizer=$Organizer
					}
					$origItem = Get-OriginalAppointment -Service $orgService -ICalUid $item.ICalUid -start $item.Start.AddHours(-2) -end $item.End.AddHours(2) #-AppointmentType $item.AppointmentType.tostring()

					if (-not $origItem) {
						Write-Warning "`n`n`nНе найден оригинал встречи у организатора. Удаляем вложение у пользователя $mailbox ..."
						Save-OrgItemAndSendUpdate $item $service
						continue
					}

					$origItem.Load()

				} catch {
					Write-Warning "Ошибка impersonation: $($_.Exception.Message)"
					continue
				}

				if ($origItem.AppointmentType -eq "Occurrence") {
					$itemTMP = $null
					$itemTMP = Get-Master $origItem $orgService
					if ($itemTMP.AppointmentType -eq "RecurringMaster" -and $itemTMP.LastOccurrence.end -ge (get-date)) { 
						Write-Warning "Дата завершения повторяющихся событий еще не наступила..."
						continue 
					} else {
						$origItem=$itemTMP
					}
				}
				Save-OrgItemAndSendUpdate $origItem $orgService
		}
		
		else {
			Save-OrgItemAndSendUpdate $item $service
		}
	}

Write-Host "`nЗавершено" -ForegroundColor Green
