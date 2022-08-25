function Get-CallStatus {
    begin {
        #If LastNotification.xml doesn't exist, create it
        $LastNotificationPath = '~/LastNotification.xml'
        if (!(Test-Path $LastNotificationPath)) {
            Get-Date 0 | Export-Clixml $LastNotificationPath 
        }
    }
    process {
        #Check if it's more than 5min since last notification
        $CurrentDate = Get-Date
        [datetime]$LastNotification = Import-Clixml $LastNotificationPath
        $5minSinceLastNotification = ($CurrentDate - $LastNotification).Minutes -gt 4

        #Check if day is holiday
        $holiday = (Invoke-RestMethod -Uri "https://webapi.no/api/v1/holidays").data.date -contains (Get-Date -Format "MM/dd/yy")

        #Check if PowerHTML is installed.
        if (!(Get-InstalledModule PowerHTML -ErrorAction SilentlyContinue)) {
            Install-Module -Name PowerHTML -Force
        }

        #Scrape wallboard website for data
        $HTMLPage = Invoke-WebRequest -Uri http://uccxwb01.intern.mrfylke.no/wb-sadmit2/default.asp?Stats=CSQ
        $HTMLParsed = $HTMLPage | ConvertFrom-Html
        $table = $HTMLParsed.SelectNodes('//table') | Where-Object { $_.Element('tr') }
        [int]$AvailableAgents = $table.where({ $_.innertext -like "Available Agents*" }).childnodes[1].innertext
        [int]$CallsInQueue = $table.where({ $_.innertext -like "Calls In Queue*" }).childnodes[1].innertext
        [int]$TalkingAgents = $table.where({ $_.innertext -like "Talking Agents*" }).childnodes[1].innertext
        [timespan]$CurrentWaitTime = $table.where({ $_.innertext -like "Current Wait Time*" }).childnodes[1].innertext
    
        #Get how many agents that are set to ready
        [int]$LoggedInAgents = $AvailableAgents + $TalkingAgents

        #If there are [int] available agents and calls in queue are more than [int]
        if ($AvailableAgents -eq 0 -and $CallsInQueue -gt 0) {
            $message = "Det er $CallsInQueue calls in queue. Sjekk Cisco Finesse."
        }
        #If there are 1 or less agents set to ready
        elseif ($LoggedInAgents -le 1) {
            $message = "Det er $LoggedInAgents agenter online. Sjekk Cisco Finesse."
        }
        #If CurrentWaitTime is greater than 1min
        elseif ($CurrentWaitTime -gt '00:01:00') {
            $message = "Ventetid er $CurrentWaitTime. Sjekk Cisco Finesse."
        }

        if ($5minSinceLastNotification -and $message -and !$holiday) {
            #Execute function to send message
            Send-TeamsMessage -content $message -HookURI "https://mrfylke.webhook.office.com/webhookb2/87e9728b-4e8c-4256-bdfc-cba76bf9a8f7@b932ece7-9cdf-4d94-b4c1-15256e43c7ea/IncomingWebhook/5fa1d1eb28e747e181ec5c3cdbbacd9e/d2eb14ef-985d-49d1-8fd8-9d3f6365f057"
            #Log when notification is sent into xml file
            Get-Date | Export-Clixml $LastNotificationPath
        }
        Write-Host "Det er $CallsInQueue calls in queue."
        Write-Host "Det er $LoggedInAgents agenter online."
        Write-Host "Ventetid er $CurrentWaitTime."
    }
}
Get-CallStatus