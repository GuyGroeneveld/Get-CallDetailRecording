# Get-CallDetailRecording
This script retrieves CDR details for Teams meeting and tries to link it with calendar entries. We then make some stats.

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
