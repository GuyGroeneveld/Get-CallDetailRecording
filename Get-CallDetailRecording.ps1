<#
  .SYNOPSIS
  This script retrieves CDR details for Teams meeting and tries to link it with calendar entries. We then make some stats.

    #############################################################################
    All the EWS part comes from Glen Scales and is documented on
    https://dev.to/gscales/accessing-microsoft-teams-summary-records-cdr-s-for-calls-and-meetings-using-exchange-web-services-3581

    Glen has also an excellent blog regarding EWS and Graph
    https://gsexdev.blogspot.com/

    #############################################################################

    This script assumes the OneDrive client is installed. We are getting the Microsoft.IdentityModel.Clients.ActiveDirectory.dll
    From the OneDrive installation point

    #############################################################################


    ===========================================
    Version:
	1.0 Initial version
    2.0 Changed the logic to keep only those having attended the live session
    2.1 Fixed session retrieval. Was getting only online meetings. It appears some are not appearing like this
    2.2 Forced UTC time everywhere I could think of
    2.3 fixed bug where $Attendee was not reinitialized and caused false entries
    2.4 fixed one attendee only sessions. There can me more than one CDR per meeting. I was only looking at the first one
        Changed the logic to get all existing CDR logs for a meeting.
    2.5 Fixed negative livetime
    2.6 Was too restrictive on who ends up in the detailed list. Changed "if" statement with an "Or"
    2.7 The Get Calendar Item was doing datetime to string and back to datetime. This was causing sometime issues with regional settings
    3.0 Changed the authentication function and fixed some bugs relative to CDR logs availability
    3.1 It may happen that a CDR log has "Unknown" as Meeting Cid. Added a function in that case to try to find a match for this
        CDR log in the calendar
    3.2 switched off from batch load with LoadPropertiesForItems to item per item load to avoid a random bug
    3.3 Added the possibility to only get the attendant with a minimum attendance duration
        Added the possibility to use another calendar mailbox (Useful for summer school)
    3.4 The minimum attendance duration allowed to expose a bug when you have several CDRLogs for a singlez meeting
        This made me review the way we parse the body
        This part is starting to become ugly. Will need someday to review it
    3.5 Fixed a bug when trying to match CDR flagged UNKNOWN

    Author: Guy Groeneveld <guyg@outlook.com>
    ===========================================

  .DESCRIPTION
  In a hidden folder in your mailbox the call details for Teams meeting and Calls is kept. Unfortunately those entries do not have the calendar details.

  Glen Scales has developped a script that gets calendar details and CDR details and tries to make a match.
  See https://dev.to/gscales/accessing-microsoft-teams-summary-records-cdr-s-for-calls-and-meetings-using-exchange-web-services-3581

  I have added some few functions to analyze in more details the body of the CDR to get who did attend the meeting


  .EXAMPLE
  To get CDR Items only for the months of July and August
  .\Get-CallDetailRecording -MailboxName john@contoso.com -startdatetime '07/07/2019' -enddatetime '07/31/2019'


  .PARAMETER MailboxName
  String parameter. Defines the mailbox to search for the CDR Logs. If the parameter CalendarMailbox is not set, this also defines
  the mailbox to be used to retrieve the calendar entries.

  .PARAMETER CalendarMailbox
  String parameter. Defines the mailbox to be used to retrieve the calendar entries. If this parameter is not set, we will use the
  MailboxName to get the calendar entries

  .PARAMETER WorkPath
  String parameter. Defines where the files are. By default it uses the current folder path.

  .PARAMETER StartDateTime
  Datetime parameter. Defines the start of the time window in which we are going to search for calendar entry and CDR logs.
  Defaults to three days back

  .PARAMETER EndDateTime
  Datetime parameter. Defines the end of the time window in which we are going to search for calendar entry and CDR logs.
  Defaults to two days back

  .PARAMETER SearchCalendarFor
  String parameter. Defines the calendar entries we try to match with CDR.

  .PARAMETER MinPercentAttendance
  Integer parameter. Determine the minimum percent the attendee maust have stayed in the meeting to be counted.

  .PARAMETER MatchedMeetings
  Switch parameter. If set, we will only retrieve the CDR log details for the CDR log we could match with a calendar entry.

  .PARAMETER UTC
  Switch parameter. If set, the start and end time are in UTC. By default, the start and end time are using default locals.

  .PARAMETER force
  Switch parameter. If set force to start from an empty resolved email address list. Every times we resolve an email address
  we keep it in a file named Resolved.xml. This avoid making too much EWS requests to Exchange online and getting throttled.
  #>


[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,Mandatory=$false,Position=0)]
        [string]$MailboxName="",
        [string]$CalendarMailbox="",
        [string]$WorkPath = (Get-Location).Path,
        [datetime]$StartDateTime=((Get-Date).AddDays(-3).AddHours(-(get-date).hour)),
        [datetime]$EndDateTime= ((Get-Date).AddDays(-2).AddHours(-(get-date).hour)),
        [string]$SearchCalendarFor="",
        [int]$MinPercentAttendance=1,
        [switch]$MatchedMeetings=$false,
        [switch]$UTC=$false,
        [switch]$force=$false
        )


if ($UTC)
{
    $startdatetime = [DateTime]::SpecifyKind($startdatetime,[DateTimeKind]::Utc)
    $EndDateTime = [DateTime]::SpecifyKind($EndDateTime,[DateTimeKind]::Utc)
}
else
{
    $startdatetime = [DateTime]::SpecifyKind($startdatetime,[DateTimeKind]::Local)
    $EndDateTime = [DateTime]::SpecifyKind($EndDateTime,[DateTimeKind]::Local)
}

if ($startdatetime -gt (Get-Date).AddDays(-3) -and $EndDateTime -gt (Get-Date).AddDays(-2))
{
    Write-Host "The Call Details Recording may take up to two days to appear in your mailbox.
You have chosen an earlier start or end date, if the script fails to get the call details retry later " -ForegroundColor Red
}

function Connect-EXCExchange {
	<#
	.SYNOPSIS
		A brief description of the Connect-EXCExchange function.

	.DESCRIPTION
		A detailed description of the Connect-EXCExchange function.

	.PARAMETER MailboxName
		A description of the MailboxName parameter.

	.PARAMETER Credentials
		A description of the Credentials parameter.

	.EXAMPLE
		PS C:\> Connect-EXCExchange -MailboxName 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,

		[Parameter(Position = 1, Mandatory = $False)]
		[System.Management.Automation.PSCredential]
		$Credentials,

		[Parameter(Position = 2, Mandatory = $False)]
		[switch]
		$ModernAuth,

		[Parameter(Position = 3, Mandatory = $False)]
		[String]
		$ClientId
	)
	Begin {
		## Load Managed API dll
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		if (Test-Path ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")) {
			Import-Module ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")
			$Script:EWSDLL = $script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll"
			write-verbose ("Using EWS dll from Local Directory")
		}
		else {


			## Load Managed API dll
			###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
			$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
			if (Test-Path $EWSDLL) {
				Import-Module $EWSDLL
				$Script:EWSDLL = $EWSDLL
			}
			else {
				"$(get-date -format yyyyMMddHHmmss):"
				"This script requires the EWS Managed API 1.2 or later."
				"Please download and install the current version of the EWS Managed API from"
				"http://go.microsoft.com/fwlink/?LinkId=255472"
				""
				"Exiting Script."
				exit


			}
		}

		## Set Exchange Version
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

		## Create Exchange Service Object
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials

		#Credentials Option 1 using UPN for the windows Account
        #$psCred = Get-Credential
        Write-Host "Connecting to Exchange Online" -ForegroundColor Yellow
		if ($ModernAuth.IsPresent) {
			Write-Verbose("Using Modern Auth")
			if ([String]::IsNullOrEmpty($ClientId)) {
				$ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
			}
            $user = $env:UserName
            if (Test-Path "C:\Users\$user\OneDrive - Microsoft\Documents 1\PowerShell\Modules\AzureAD\2.0.2.16\Microsoft.IdentityModel.Clients.ActiveDirectory.dll")
            {
               Import-Module "C:\Users\$user\OneDrive - Microsoft\Documents 1\PowerShell\Modules\AzureAD\2.0.2.16\Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -force
            }
            else
            {
				"$(get-date -format yyyyMMddHHmmss):"
                "This script requires the Azure Active Directory Dll"
                "We assume you have the OneDrive client installed and try to use it"
				"Please Install the OneDrive client or point the dll to another place in your directory"
				""
				"Exiting Script."
				exit
			}
			$Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
			if ($Credentials -eq $null) {
				$PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto

				$token = ($Context.AcquireTokenAsync("https://outlook.office365.com", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
				$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
			}else{
				$AADcredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList  $Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString()
				$token = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($Context,"https://outlook.office365.com",$ClientId,$AADcredential).result
				$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
			}
		}
		else {
			Write-Verbose("Using Negotiate Auth")
			if(!$Credentials){$Credentials = Get-Credential}
			$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
			$service.Credentials = $creds
		}

		#Credentials Option 2
		#service.UseDefaultCredentials = $true
		#$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates

		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
		$Params = New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable = $False
		$Params.GenerateInMemory = $True
		$Params.IncludeDebugInformation = $False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

		$TASource = @'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() {
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert,
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@
		$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
		$TAAssembly = $TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll

		## end code from http://poshcode.org/624

		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use

		#CAS URL Option 1 Autodiscover
		$service.AutodiscoverUrl($MailboxName, { $true })
		#Write-host ("Using CAS Server : " + $Service.url)

		#CAS URL Option 2 Hardcoded

		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"
		#$service.Url = $uri

		## Optional section for Exchange Impersonation

		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		if (!$service.URL) {
			throw "Error connecting to EWS"
		}
		else {
            return $service
		}
	}
}

function Get-TeamsMeetingsFolder{
    param(
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position =7, Mandatory = $false) ]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service

    )
    Process {
        if($service -eq $null){
            if ($basicAuth.IsPresent) {
                if (!$Credentials) {
                    $Credentials = Get-Credential
                }
                $service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth -Credentials $Credentials
            }
            else {
                $service = Connect-EXCExchange -MailboxName $MailboxName -ModernAuth  #-AccessToken $AccessToken
            }
            $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
            if ($useImpersonation.IsPresent) {
                $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
            }
        }
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
        $TeamMeetingsFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{07857F3C-AC6C-426B-8A8D-D1F7EA59F3C8}"), "TeamsMeetingsFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $psPropset.Add($TeamMeetingsFolderEntryId)
        $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
        $FolderIdVal = $null

        $TeamsMeetingFolder = $null
        if ($RootFolder.TryGetProperty($TeamMeetingsFolderEntryId,[ref]$FolderIdVal))
        {
            $TeamMeetingFolderId= new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-","")) -service $service))
            $TeamsMeetingFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$TeamMeetingFolderId);
        }
        return $TeamsMeetingFolder


    }
}

function Get-CalendarForCDRLog
{
    param (
        [datetime]$CDRLogStartTime,
        [hashtable]$CalendarHashTable,
        $CDRlog
        )

    $CalendarCdrMatch=[System.Collections.ArrayList]@()
    <#This function purpose it to try to match a CDR log with a calendar entry when the meeting Cid is missing in the CDR Log
    It happens sometimes that the CDR log is created with "Unknown" in the title. We will attempt to find a calendar entry close to the
    CDRLog and double check with the match of some of the attendees and some the meeting request distribution list#>
    foreach ($key in $CalendarHashTable.Keys)
    {
        $matchObj = "" | Select-Object NumOfMatch, numCDRAttendees, Key
        if ($CDRLogStartTime -gt ($CalendarHashTable[$key].start.ToUniversalTime()).addminutes(-5) -and $CDRLogStartTime -lt ($CalendarHashTable[$key].start.ToUniversalTime()).addminutes(5))
        {
            <#We have a calendar entry close to the CDR log. Now let see if there is some attendees matching#>
            $matchObj.Key = $key
            $TeamsAttendees = $CDRlog.DisplayTo.Split(";")
            $matchObj.numCDRAttendees = $TeamsAttendees.count
            foreach ($TeamAttendee in $TeamsAttendees)
            {
                if ($CalendarHashTable[$key].DisplayCc)
                {
                    if ($CalendarHashTable[$key].DisplayCc.Contains($TeamAttendee))
                    {
                        ([int]$matchObj.NumOfMatch)++
                    }
                }
                if ($CalendarHashTable[$key].DisplayTo)
                {
                    if ($CalendarHashTable[$key].DisplayTo.Contains($TeamAttendee))
                    {
                        ([int]$matchObj.NumOfMatch)++
                    }
                }
            }
        }
        if ($matchObj.Key)
        {
            [void]$CalendarCdrMatch.Add($matchObj)
        }

    }
    if ($CalendarCdrMatch.count -eq 1)
    {
        return $CalendarCdrMatch[0].Key
    }
    elseif ($CalendarCdrMatch.Count -gt 1)
    {
        <#We have found more than one calendar entry matching the CDRLog start time. We will get the one with the more matched attendees#>
        [int]$MaxMatch = $CalendarCdrMatch | Measure-Object -Maximum -Property NumOfMatch
        foreach ($CdrMatch In  $CalendarCdrMatch)
        {
            if ($CdrMatch.NumOfMatch -eq $MaxMatch)
            {
                return $CdrMatch.Key
            }
        }

    }
    else
    {
        return ""
    }

}

function Get-TeamsCDRItems{
    param(
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 10, Mandatory = $false)] [DateTime]$startdatetime = (Get-Date).AddDays(-60),
        [Parameter(Position = 11, Mandatory = $false)] [datetime]$enddatetime = (Get-Date)

    )
    Process {
        <#This function originated from gsexdev but has been widely modified as there was many issues where the match between the
        calendar item and the cdr log item didn't work#>
        if ($basicAuth.IsPresent) {
            if (!$Credentials) {
                $Credentials = Get-Credential
            }
            $script:service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth -Credentials $Credentials
        }
        else {
            #Changed $service scope to script as I will need it to resolve contact email later
            $script:service = $service = Connect-EXCExchange -MailboxName $MailboxName -ModernAuth
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        }
        $TeamsMeetingFolder = Get-TeamsMeetingsFolder -service $service -MailboxName $MailboxName
        $CalendarItemsHash = Get-CalendarItems -MailboxName $MailboxName -service $service -startdatetime $startdatetime -enddatetime $enddatetime
        $fiItems = $null
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
        $ItemPropsetIdOnly = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ivItemView.PropertySet = $ItemPropsetIdOnly
        $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $CleanObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $SkypeTeamsMeetingUrl = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x8AF2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $SkypeTeamsProperties = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SkypeTeamsProperties", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $PR_START_DATE  = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0060, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
        $PR_END_DATE  = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0061, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
        $ItemPropset.Add($CleanObjectId)
        $ItemPropset.Add($PR_START_DATE)
        $ItemPropset.Add($PR_END_DATE)
        $ItemPropset.Add($SkypeTeamsProperties)
        $ItemPropset.Add($SkypeTeamsMeetingUrl)
        $ItemPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
        $Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $startdatetime)
        #The CDR log is not immediately processed. It may take time to appear in the mailbox
        #As we search base on DateTimeReceive and this values is set when the item is created we may miss the CDRLog
        #Adding time at the expected time range to make sure we get it
        #The experience shows it may take up to two days. If it takes more, increase it below
        $enddatetime_for_cdrlogs = $enddatetime.AddDays(3)
        $Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $enddatetime_for_cdrlogs)
        $sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
        $sfCollection.add($Sfgt)
        $sfCollection.add($Sflt)
        Write-Host "Getting Call Detail Recording from mailbox $MailboxName" -ForegroundColor Yellow

        <#We keep the match that has been done. In the end this will give us the matches that couldn't be done#>
        $CalendarMatchDone = @{}
        $MeetingCDRLogWithoutCID = @{}

        do {
            $error.clear()
            try {
                #Getting the list of items
                $fiItems = $TeamsMeetingFolder.service.FindItems($TeamsMeetingFolder.Id,$sfCollection, $ivItemView)
            }
            catch {
                write-host ("Error " + $_.Exception.Message)
                if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                    Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                    Start-Sleep -Seconds 60
                }
                $fiItems = $TeamsMeetingFolder.service.FindItems($TeamsMeetingFolder.Id, $ivItemView)
            }
            if ($fiItems.Items.Count -gt 0) {
                #loading properties for retrieved items based on the defined propset
                [Void]$TeamsMeetingFolder.service.LoadPropertiesForItems($fiItems, $ItemPropset)
            }
            if ($fiItems.Items.Count -gt 0) {
                Write-Host ("Number of Call Detail Recording logs found: " + $fiItems.Items.Count ) -ForegroundColor Yellow
                foreach ($Item in $fiItems.Items)
                {
                    #Added some more properties, they might be helpful
                    $rptObj = "" | Select-Object Subject,Start,End,Organizer,MeetingCid,CDRLog,ConversationId,DisplayCc,DisplayTo,ScheduledStart,ScheduledEnd,Duration,IsMatched
                    $Regex = [Regex]::new("(?<=/Thread Id\:)(.*)(?=/Communication Id\:)")
                    $Match = $Regex.Match($Item.Subject)
                    <#We get the meeting cid from the CDR log#>
                    $rptObj.MeetingCid = $Match.Value.Trim()
                    <#It happens that we have a CDR log for a meeting without a valid meeting cid#>
                    if (($Match.Value.length -eq 1) -and ($Item.ItemClass -like "IPM.AppointmentSnapshot.SkypeTeams.Meeting") )
                    {
                        $Splitted_Body = $item.Body.Text.Split([Environment]::NewLine)
                        $Splitted_Body = $Splitted_Body[0]
                        $Splitted_Body = $Splitted_Body.replace("Start Time (UTC): ","")
                        [datetime]$startbdy = $Splitted_Body
                        $startbdy = [DateTime]::SpecifyKind($startbdy,[DateTimeKind]::Utc)
                        if ($startbdy -gt $startdatetime -and $startbdy -lt $enddatetime)
                        {
                            $MeetingCDRLogWithoutCID.Add($startbdy,$Item)
                            $MeetingCid = Get-CalendarForCDRLog -CDRLogStartTime $startbdy -CalendarHashTable $CalendarItemsHash -CDRlog $Item
                            if (!$CalendarMatchDone.contains($MeetingCid))
                            {
                                $rptObj.MeetingCid = $MeetingCid
                            }
                        }

                    }
                    $rptObj.Subject = $Item.Subject
                    $rptObj.Start = $Item.Start
                    $rptObj.End = $Item.End
                    $rptObj.ConversationId = $Item.ConversationId.UniqueId
                    $rptObj.DisplayCc = $Item.DisplayCc
                    $rptObj.DisplayTo = $Item.DisplayTo
                    $rptObj.IsMatched = $False
                    $PR_START_DATE_Value = $null
                    if($Item.TryGetProperty($PR_START_DATE,[ref]$PR_START_DATE_Value)){
                        [datetime]$rptObj.Start = $PR_START_DATE_Value.ToUniversalTime()

                    }
                    $PR_END_DATE_Value = $null
                    if($Item.TryGetProperty($PR_END_DATE,[ref]$PR_END_DATE_Value)){
                        [datetime]$rptObj.End = $PR_END_DATE_Value.ToUniversalTime()

                    }

                    #If we have a meeting cid we get the calendar entry properties
                    if($CalendarItemsHash.ContainsKey($rptObj.MeetingCid))
                    {
                        $rptObj.Organizer = $CalendarItemsHash[$rptObj.MeetingCid].Organizer.Name
                        $rptObj.Subject = $CalendarItemsHash[$rptObj.MeetingCid].Subject
                        $rptObj.ScheduledStart = $CalendarItemsHash[$rptObj.MeetingCid].Start.ToUniversalTime()
                        $rptObj.ScheduledEnd = $CalendarItemsHash[$rptObj.MeetingCid].End.ToUniversalTime()
                        $rptObj.Duration = $CalendarItemsHash[$rptObj.MeetingCid].Duration
                        $rptObj.IsMatched = $true

                        if (!$CalendarMatchDone.ContainsKey($rptObj.MeetingCid))
                        {
                            $CalendarMatchDone.Add($rptObj.MeetingCid,$CalendarItemsHash[$rptObj.MeetingCid].Subject)
                        }


                    }

                    $rptObj.CDRLog = $Item.Body.Text
                    #We need to check if the CDR log is within the expected time range between the specified start and end
                    #It happens that someone starts the meeting after the real schedule by mistake
                    if ($rptObj.start -gt $startdatetime -and $rptObj.End -lt ($enddatetime.AddHours(2)) -and ($rptObj.End - $rptObj.Start).totalminutes -gt 1 )
                    {
                        Write-Output $rptObj
                    }
                    else
                    {
                        $rptObj = $null
                    }

                }
            }
            $ivItemView.Offset += $fiItems.Items.Count
        }while ($fiItems.MoreAvailable)

        foreach ($CalMatch in $CalendarMatchDone.Keys)
        {
            $CalendarItemsHash.remove($CalMatch)
        }
        if ($CalendarItemsHash.count -gt 0)
        {
            foreach ($key in $CalendarItemsHash.Keys)
            {
               Write-Host ("Couldn't find CDR log for: " + $CalendarItemsHash[$key].subject + " - due to start at " + $CalendarItemsHash[$key].start) -ForegroundColor Magenta
            }

        }
    }
}



function Get-CalendarItems{
    param (
    [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
    [Parameter(Position = 2, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
    [Parameter(Position = 3, Mandatory = $false)] [DateTime]$startdatetime,
    [Parameter(Position = 4, Mandatory = $false)] [datetime]$enddatetime
    )
    process
    {
        if (!$CalendarMailbox.Length -gt 1)
        {
            $CalendarMailbox = $MailboxName
        }

        Write-Host "Getting calendar entries in mailbox $CalendarMailbox from:"  (($startdatetime.ToLocalTime()).ToString())  " to:" (($enddatetime.ToLocalTime()).ToString()) -ForegroundColor Yellow
        $rptHash = @{}
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$CalendarMailbox)
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
        $Recurring = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment, 0x8223,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);
        $TeamMeetingsFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{07857F3C-AC6C-426B-8A8D-D1F7EA59F3C8}"), "TeamsMeetingsFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        #Define the properties to retrieve
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $SkypeTeamsProperties = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SkypeTeamsProperties", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        $psPropset.Add($Recurring)
        $psPropset.Add($SkypeTeamsProperties)
        $psPropset.Add($TeamMeetingsFolderEntryId)
        $psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;

        #Define the calendar view
        $CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startdatetime,$enddatetime,1000)
        #Getting the items form the calendar based on the calendar view
        $fiItems = $service.FindAppointments($Calendar.Id,$CalendarView)
        if($fiItems.Items.Count -gt 0){
            $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
            $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
            $ItemColl = [Activator]::CreateInstance($type)
            foreach($Item in $fiItems.Items){
                $Item.Load($psPropset)
                $ItemColl.Add($Item)
                }
            <#Loading properties defined in the property set for the items we got
            We randomly get this error message when loading the item properties. I've not been able to debug it
                Exception calling "LoadPropertiesForItems" with "2" argument(s): "Requested value 'GroupMailbox' was not found."
                At C:\temp\script\Get-CallDetailRecording.ps1:478 char:13
                +             [Void]$service.LoadPropertiesForItems($ItemColl,$psPropse ...
                +             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
                + FullyQualifiedErrorId : ArgumentException
            This will need to be investigated sometime#>
            #[Void]$service.LoadPropertiesForItems($ItemColl,$psPropset)

            <#We keep only the calendar entries that are online meetings#>
            $ItemColl = $ItemColl | Where-Object { $_.isonlinemeeting -eq $true -or $_.Body.text -like "*teams.microsoft.com*" }
        }


        Write-Host "Found " ($ItemColl.count) " calendar entries" -ForegroundColor Yellow
        $itemcount = 0
        foreach($Item in $ItemColl){

        $check = "*" + $SearchCalendarFor + "*"

        If ($Item.subject -like $check)
        {

            # Changed from gscale source code as sometimes the cid is not documented where it was looking for
            # but it may be in the extended properties when it does exist
            if ($SearchCalendarFor)
            {
                Write-Verbose "Found a calendar entry for: $SearchCalendarFor"
            }

            $itemcount++

            #First we try to get the meeting cid from the extended properties if it is documented
            if ( $Item.ExtendedProperties[1].Value -ne $null)

            {

                $HashKey = ($Item.ExtendedProperties[1].Value | ConvertFrom-Json).cid
                if (!$rptHash.ContainsKey($HashKey))
                {

                    $rptHash.Add($HashKey,$Item)
                }
            }

            #Second we check the item body to find and parse the the meeting cids
                if ($Item.body.text -ne $null)
                {

                    if ($Item.body.text.contains("19%3ameeting"))
                    {

                        #We may have several meeting lines in the body as the meeting can be organized from another meeting
                        $linesmatch = ($Item.body.text).Split([Environment]::NewLine) | Select-String -Pattern '<https://teams.microsoft.com/l/meetup-join/19%3ameeting_'

                        for ($l=0 ; $l -lt $linesmatch.count ; $l++)
                        {

                            $MeetingLine = $linesmatch[$l].ToString()
                            $MeetingLine = $MeetingLine.Substring($MeetingLine.IndexOf("meetup-join/") + 12,$MeetingLine.Length - ($MeetingLine.IndexOf("meetup-join/")+12))
                            $MeetingLine = $MeetingLine.Substring(0,$MeetingLine.IndexOf("/0"))
                            $CidFromBody = $MeetingLine.Replace("%3a",":")
                            $CidFromBody = $CidFromBody.Replace("%40","@")

                            if (!$rptHash.ContainsKey($CidFromBody) -and ($CidFromBody.EndsWith(".v2") -or $CidFromBody.EndsWith(".skype")))
                            {
                                $rptHash.Add($CidFromBody,$Item)
                            }
                        }

                    }
                }
        }

    }


        return $rptHash

    }
}


function ConvertId{
    param (
            [Parameter(Position=1, Mandatory=$false)] [String]$HexId,
            [Parameter(Position=2, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
       )
    process{
        $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
        $aiItem.Mailbox = $MailboxName
        $aiItem.UniqueId = $HexId
        $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId
        $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
     return $convertedId.UniqueId
    }
   }




function Get-ContactInfo{
    param(
        [Parameter(Position = 1, Mandatory = $true)] [string]$LineToResolve,
        [Parameter(Position = 2, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [string]$Event

    )

    #We extract from the GAL the details based on the email address founf in the CDRLog

    $emlObj = "" | Select-Object Displayname,Manager,Department,OfficeLocation,JobTitle
    $Email = $LineToResolve.Substring($LineToResolve.IndexOf("]")+2)
    $Email = $Email.Replace($Event,"")

    #We are keeping the resolved email in a hash table to avoid repetitive EWS call and maybe be throttled
    #We check if the email ahs already been resolved and get the properties from the hash table instead of calling EWS
    if ($ResolvedEmail.containskey($Email))
    {
        $ResolvedContact = $ResolvedEmail[$Email]
    }

    else
    {
        #We didn't find the contact in the hash, so we call EWS to get the properties and we also store it in the hash
        $ResolvedContact = $service.ResolveName($Email,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true)
        $ResolvedEmail.add($Email,$ResolvedContact)

    }
    if ($ResolvedContact[0].contact.displayname)
    {
        $emlObj.Displayname = $ResolvedContact[0].contact.displayname
        $emlObj.Manager = $ResolvedContact[0].contact.manager
        $emlObj.Department = $ResolvedContact[0].contact.Department
        $emlObj.OfficeLocation = $ResolvedContact[0].contact.OfficeLocation
        $emlObj.JobTitle = $ResolvedContact[0].contact.JobTitle
    }
    Else
    {
        #The email could not be found in the GAL. Either it is someone not working in MS or somebody who left MS
        $emlObj.Displayname = $Email
        if ($Email.Contains("microsoft"))
        {
            $emlObj.OfficeLocation = "Not anymore in Microsoft"
        }
        else
        {
            $emlObj.OfficeLocation = "External"
        }

    }


    return $emlObj
}


Function Get-TimeEvent
{
Param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$LineToResolve="",

        [string]$event = ""
        )

        $tmObj = "" | Select-Object Time
        #UTC not being understood, we need to change it to 'Z' as this is the known notation for UTC
        $TimePart = $LineToResolve.Replace(' (UTC)',' Z')
        [datetime]$EventTime = $TimePart.Substring(1,$TimePart.IndexOf(']')-1)

        [datetime]$tmObj.Time = $EventTime.ToUniversalTime()


        Return $tmObj

}


Function Get-CDRDetails
{
<#This function takes the resulted list of Get-CDRItems where we retrieved the details of the Teams log
We need now to parse the body to retrieve the attendees list and the time they attended#>
[CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        $MeetingList,
        [string]$SubjectSearch="",
        [string]$SearchPath= (Get-Location).Path,
        [string]$MailboxName="",
        [switch]$ForceFileReset=$force

    )

        $MeetingAttendees = [System.Collections.ArrayList]@()
        $script:ResolvedEmail=@{}
        [int]$SessionsCount=0
        $SessionsDone = @{}


        #We load the resolved email address hash table from previous run unless we want to force its reload
        if ((Test-Path -Path ($WorkPath + "\ResolvedEmail.xml")) -and ($Force -eq $False))
            {
                $ResolvedEmail = Import-Clixml -Path ($WorkPath + "\ResolvedEmail.xml")
            }
        Else
            {
                Write-Host "No resolved email address to inject. Will start from scratch" -ForegroundColor Yellow
            }



        for ($z = 0; $z -lt $MeetingList.count ; $z++)
        {

            $MeetingDetails = [System.Collections.ArrayList]@()
            $CDRToBeDone = [System.Collections.ArrayList]@()
            <#We check if we want to see only the list of attendees for the CDR logs that could be matched with a calendar entry#>
            if (($MeetingList[$z].IsMatched -eq $false) -and ($MatchedMeetings -eq $true))

            {
                $skippedmeetingsubject = $MeetingList[$z].subject
                Write-Verbose "Skipped unresolved meeting $skippedmeetingsubject"
            }
            <#We may have several CDR logs for the same meeting as a log is created each time the first one logs on to the call.
            As we merge the logs the first time we have a log in the list, to make sure we don't do it twice we set $SessionDone
            This way the next time we find a log matching the name we make sure we don't parse all the entries several times#>
            Elseif (!( $SessionsDone.containskey($MeetingList[$z].Subject)))
            {
                Write-Progress -Activity "Getting meeting call detail recording details for session:" -Status ($MeetingList[$z].Subject) -PercentComplete ($z*100/($MeetingList.count)) -Id 1

                $SessionsCount++

                $SessionsDone.Add(($MeetingList[$z].Subject),($MeetingList[$z].Subject))
                #We may have more than one meeting CDR with the same subject. We need to parse both

                $CDRToBeDone = @($MeetingList | Where-Object { $_.subject -eq ($MeetingList[$z].Subject) })
                $CDRList = [System.Collections.ArrayList]@()
                $meeting_text = $null
                for ($sub=0 ; $sub -lt $CDRToBeDone.count ; $sub++)
                {
                    $cdrLogObj = "" | Select-Object CDRStart,CDREnd,CDRDuration,CDRBody
                    $meeting_text = $CDRToBeDone[$sub].CDRLog
                    $splitted_CDRLogText = $meeting_text.Split([Environment]::NewLine)
                    $splitted_CDRLogText = $splitted_CDRLogText | Where-Object { $_ -ne "" }
                    [string]$CDRStart = $splitted_CDRLogText | Select-String -Pattern "Start Time"
                    [string]$CDREnd = $splitted_CDRLogText | Select-String -Pattern "End Time"
                    [string]$CDRDuration = $splitted_CDRLogText | Select-String -Pattern "Duration:"
                    [datetime]$cdrLogObj.CDRStart = $CDRStart.Replace("Start Time (UTC): ","")
                    $cdrLogObj.CDRStart = [DateTime]::SpecifyKind($cdrLogObj.CDRStart,[DateTimeKind]::Utc)
                    [datetime]$cdrLogObj.CDREnd = $CDREnd.Replace("End Time (UTC): ","")
                    $cdrLogObj.CDREnd = [DateTime]::SpecifyKind($cdrLogObj.CDREnd,[DateTimeKind]::Utc)
                    [timespan]$cdrLogObj.CDRDuration = $CDRDuration.Replace("Duration: ","")
                    $cdrLogObj.CDRBody = $splitted_CDRLogText

                    <#We keep only the CDR logs details for CDR logs that started before the scheduled end and ended after the scheduled start
                    We don't care about CDR logs created before or after as they don't contains people who did attend the meeting during the scheduled time#>
                    if ($cdrLogObj.CDRStart -lt $MeetingList[$z].ScheduledEnd -and  $cdrLogObj.CDREnd -gt $MeetingList[$z].ScheduledStart )
                    {
                        [void]$CDRList.add($cdrLogObj)
                    }

                }

                $splitted_text = [System.Collections.ArrayList]@()
                $RealCDRStart = $MeetingList[$z].start
                $RealCDREnd = $MeetingList[$z].end
                [timespan]$RealDuration = 0

                foreach ($CDRListItem in $CDRList)
                {
                    foreach ($CDRLogLine in $CDRListItem.CDRBody)
                    {
                        [void]$splitted_text.Add($CDRLogLine)
                    }

                    if ($RealCDRStart -gt $CDRListItem.CDRStart -and $CDRListItem.CDRStart -gt $RealCDRStart.addminutes(-10) )
                    {
                        $RealCDRStart = $CDRListItem.CDRStart
                    }
                    if ($RealCDREnd -lt $CDRListItem.CDREnd -and $CDRListItem.CDREnd -lt $RealCDREnd.AddDays(1))
                    {
                        $RealCDREnd = $CDRListItem.CDREnd
                    }
                    if ($RealDuration -lt $CDRListItem.CDRDuration)
                    {
                        $RealDuration = $CDRListItem.CDRDuration
                    }


                }

                <#Now that we have a merged body text from all the CDR logs entry we may have for this subject, we can now parse the text to retrieve log details
                As we have potentially merged several CDR logs, or we may have people joining and leaving several time in the log, we need to retrieve that info#>


                for ($j=0;$j -lt $splitted_text.count ; $j++)
                {
                    [timespan]$LiveDuration = $RealDuration
                    $Attendee = $null

                    Write-Progress -Activity "Getting attendees details" -Status ($splitted_text[$j]) -PercentComplete ($j*100/($splitted_text.Count)) -Id 2

                    if ($splitted_text[$j].contains("joined"))
                    {
                        $EventInfo = "Joined"
                        $Attendee = Get-ContactInfo -LineToResolve $splitted_text[$j] -Event " joined." -service $service

                        if ($LiveDuration -lt $MeetingList[$z].Duration)
                        {
                            $TimeInfo = Get-TimeEvent -LineToResolve $splitted_text[$j] -event $EventInfo
                            $TimeInfo.time = $TimeInfo.time.ToUniversalTime()
                        }
                        Else
                        {
                            $TimeInfo = Get-TimeEvent -LineToResolve $splitted_text[$j] -event $EventInfo
                            $TimeInfo.time = $TimeInfo.time.ToUniversalTime()
                        }
                    }
                    if ($splitted_text[$j].contains("left"))
                    {
                        $EventInfo = "Left"
                        $Attendee = Get-ContactInfo -LineToResolve $splitted_text[$j] -Event " left." -service $service
                        $TimeInfo = Get-TimeEvent -LineToResolve $splitted_text[$j] -event $EventInfo
                        $TimeInfo.time = $TimeInfo.time.ToUniversalTime()
                    }

                    <#Now that we have a "Joined" or a "Left" event, we add the details and store it in $MeetingDetails#>
                    $meeting_data = "" | Select-Object Attendee,Event,Time,Manager,Department,OfficeLocation,JobTitle,Subject,MeetingScheduledStart,MeetingStart,MeetingScheduledEnd,MeetingEnd,MeetingScheduledDuration,Organizer
                    if($Attendee)
                    {

                        $meeting_data.Attendee = $Attendee.displayname
                        $meeting_data.Event = $EventInfo
                        $meeting_data.Time = $TimeInfo.Time.touniversaltime()
                        $meeting_data.Manager = $Attendee.manager
                        $meeting_data.Department = $Attendee.department
                        $meeting_data.OfficeLocation = $Attendee.OfficeLocation
                        $meeting_data.JobTitle = $Attendee.JobTitle
                        $meeting_data.Subject = $MeetingList[$z].subject
                        if ($MeetingList[$z].IsMatched -eq $true)
                        {
                            $meeting_data.MeetingScheduledStart = $MeetingList[$z].ScheduledStart.touniversaltime()
                            $meeting_data.MeetingScheduledEnd = $MeetingList[$z].ScheduledEnd.touniversaltime()
                        }
                        else
                        {
                            $meeting_data.MeetingScheduledStart = $RealCDRStart.touniversaltime()
                            $meeting_data.MeetingScheduledEnd = $RealCDREnd.touniversaltime()
                        }
                        $meeting_data.MeetingStart = $RealCDRStart.touniversaltime()
                        $meeting_data.MeetingEnd = $RealCDREnd.touniversaltime()
                        $meeting_data.MeetingScheduledDuration = $MeetingList[$z].ScheduledEnd - $MeetingList[$z].ScheduledStart
                        $meeting_data.Organizer = $MeetingList[$z].Organizer

                        [void]$MeetingDetails.Add($meeting_data)
                    }

                }
                <#Merging attendee events to get live minutes attending. We have one or more "Joined" event and one or more "Left" event
                We group $MeetingDetails per attendees. This gives the list of events for an attendee. We then merge them all trying to get
                the first "Joined" of the list and the last "Left" of the list. We then compute the livetime based on "Left" - "Joined"#>
                $grouped = $MeetingDetails | Group-Object attendee
                for ($att = 0 ; $att -lt $grouped.length ; $att++)
                    {
                    Write-Progress -Activity "Merging Details and getting live time attended" -Status ($grouped[$att].group[0].Attendee) -PercentComplete ($att*100/($grouped.Count)) -Id 3
                    [datetime]$joined = 0
                    [datetime]$Leaved = 0
                    [timespan]$LiveTime = 0
                    $GroupTimeSorted = $grouped[$att].group | Sort-Object Time
                    $evtObj = "" | Select-Object Attendee,Manager,Department,OfficeLocation,JobTitle,Subject,JoinedAt,LeftAt,LiveTime,PercentAttended,MeetingScheduledStart,MeetingStart,MeetingScheduledEnd,MeetingEnd,Organizer
                    $evtObj.Attendee = $GroupTimeSorted[0].Attendee
                    $evtObj.Manager = $GroupTimeSorted[0].Manager
                    $evtObj.Department = $GroupTimeSorted[0].Department
                    $evtObj.OfficeLocation = $GroupTimeSorted[0].OfficeLocation
                    $evtObj.JobTitle = $GroupTimeSorted[0].JobTitle
                    $evtObj.Subject = $GroupTimeSorted[0].Subject


                    if ($MeetingList[$z].IsMatched -eq $true)
                    {
                        $evtObj.MeetingScheduledStart = $GroupTimeSorted[0].MeetingScheduledStart.ToUniversalTime()
                        $evtObj.MeetingScheduledEnd = $GroupTimeSorted[0].MeetingScheduledEnd.ToUniversalTime()
                    }
                    else
                    {
                        $meeting_data.MeetingScheduledStart = $RealCDRStart.touniversaltime()
                        $meeting_data.MeetingScheduledEnd = $RealCDREnd.touniversaltime()
                    }
                    $evtObj.MeetingStart = $RealCDRStart.ToUniversalTime()
                    $evtObj.MeetingEnd = $RealCDREnd.ToUniversalTime()
                    $evtObj.Organizer = $GroupTimeSorted[0].Organizer


                    for ($grp = 0; $grp -lt $GroupTimeSorted.count ; $grp++)
                    {

                        if ($GroupTimeSorted[$grp].Event -eq "Joined")

                        {
                            if (!$evtObj.JoinedAt)
                            {
                                $Joined  = $GroupTimeSorted[$grp].Time.ToUniversalTime()
                                $evtObj.JoinedAt = $Joined
                            }


                        }

                        elseif ($GroupTimeSorted[$grp].Event -eq "Left" -and $GroupTimeSorted[$grp].Time.ToUniversalTime() -lt $GroupTimeSorted[$grp].MeetingScheduledEnd.ToUniversalTime())
                        {
                            $Leaved = $GroupTimeSorted[$grp].Time.ToUniversalTime()
                            if ($evtObj.LeftAt -lt $Leaved)
                            {
                                $evtObj.LeftAt = $Leaved
                            }
                        }
                        else
                        {
                            #the left event is happening out of band. The only effective time is the one spend live
                            $Leaved = $GroupTimeSorted[$grp].MeetingScheduledEnd.ToUniversalTime()
                            $evtObj.LeftAt = $GroupTimeSorted[$grp].Time
                        }

                        <#We may have a "Left" event before a "Joined" in that case the value of $Joined is 1970, this gives a huge livetime
                        Fixed this with this check#>
                        if (($Leaved -gt $joined) -and ($joined -gt $startdatetime))
                        {
                            $LiveTime = $Leaved.ToUniversalTime() - $joined.ToUniversalTime()
                            $leaved = 0

                        }
                        Elseif ($Leaved -gt 0 -and $joined -gt $GroupTimeSorted[$grp].MeetingScheduledEnd.ToUniversalTime())
                        {
                            $LiveTime = 0
                            $leaved = 0
                            $joined = 0
                        }

                        Elseif ($Leaved -eq $joined)
                        {
                            $LiveTime = 0
                            $leaved = 0
                            $joined = 0
                        }

                    }
                    $evtObj.LiveTime = $LiveTime
                    #Removing 5 minutes as the presenter use to wait that long before start to make sure they have most of attendees logged on

                    [int]$percentCalculated = 0
                    if ($MeetingList[$z].IsMatched -eq $true)
                    {
                        <#If the meeting lasted less then the planned time, we calculate the attendancy based on the effective meeting duration#>
                        if ($evtObj.MeetingEnd -lt $evtObj.MeetingScheduledEnd)
                        {
                            $percentCalculated = $LiveTime.TotalMinutes * 100 / ((($evtObj.MeetingEnd - $evtObj.MeetingScheduledStart).totalminutes) -5 )
                        }
                        else
                        {
                            <#If the meeting ended later than the planned time, we cannot know if it is due to one of the attendees who forgot to quit
                            In that case we have to use the meeting planned duration#>
                            $percentCalculated = $LiveTime.TotalMinutes * 100 / (($GroupTimeSorted[$grp-1].MeetingScheduledDuration.TotalMinutes)-5)
                        }

                    }
                    else
                    {
                        $percentCalculated = 100
                    }

                    <#When the livetime is less than one minute it may get rounded up to 0%
                    We here make sure it will be a minimum of 1%#>
                    if ($percentCalculated -gt 0 -and $percentCalculated -lt 1)
                    {
                        $percentCalculated = 1
                    }

                    <#If the attendee spent more time than needed, we set the max to 100%#>
                    if ($percentCalculated -gt 100 )
                    {
                        $percentCalculated = 100
                    }
                    $evtObj.PercentAttended =  (( "{0:N0}" -f ($percentCalculated)).tostring()) + "%"


                    <#We exlude "Applicaton...", attendee who joined after the end of the meeting and when livetime is zero#>
                    if (!($evtObj.Attendee -like "Application*") -and ($evtObj.JoinedAt -lt $evtObj.MeetingEnd) -and ($percentCalculated -ge $MinPercentAttendance))
                    {
                        [void]$MeetingAttendees.Add($evtObj)
                    }
                }
        }
    }

 write-host "Processed $SessionsCount Call Detail Recording logs according to the time period and filter" -ForegroundColor Yellow


 <#Saving resolved emails to a file in order to be able to reuse it next run#>
 $ResolvedEmail | Export-Clixml -Path ($WorkPath + "\ResolvedEmail.xml") -Force

 return $MeetingAttendees
}


$CDRItems = Get-TeamsCDRItems -MailboxName $MailboxName -startdatetime $StartDateTime -enddatetime $EndDateTime
$CDRDetails = Get-CDRDetails -MeetingList $CDRItems -MailboxName $MailboxName -SubjectSearch $SearchCalendarFor
Write-Host


return $CDRDetails


