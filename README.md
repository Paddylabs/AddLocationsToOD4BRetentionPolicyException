# AddLocationsToOD4BRetentionPolicyException

I was looking for a way to programatically add users OneDrive for Business to a Retention
Policy exception list.

This script requires a csv input file with a column heading "OneDriveUrl" with each user
in scopes OneDrive address. A quick web search will explain how to format those.

Also uses the ExchangeOnlineManagement to accomodate authenticating with MFA.
