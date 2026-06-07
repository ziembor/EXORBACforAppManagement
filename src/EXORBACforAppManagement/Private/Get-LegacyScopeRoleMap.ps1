function Get-LegacyScopeRoleMap {
    # Maps a legacy Application Access Policy permission scope (the Microsoft Graph
    # application permission / EWS scope name as granted in Microsoft Entra) to the
    # equivalent Exchange Online App RBAC role name. The keys are the "Supported
    # permissions" from the Application Access Policies documentation; every value is a
    # key that exists in Get-AppRoleMap. Permissions not present here are not covered by
    # Application Access Policies and so cannot be converted to an RBAC for App role.
    [OutputType([hashtable])]
    param()

    return [ordered]@{
        'Mail.Read'                 = 'Application Mail.Read'
        'Mail.ReadBasic'            = 'Application Mail.ReadBasic'
        'Mail.ReadBasic.All'        = 'Application Mail.ReadBasic'   # App RBAC has no .All variant
        'Mail.ReadWrite'            = 'Application Mail.ReadWrite'
        'Mail.Send'                 = 'Application Mail.Send'
        'MailboxSettings.Read'      = 'Application MailboxSettings.Read'
        'MailboxSettings.ReadWrite' = 'Application MailboxSettings.ReadWrite'
        'Calendars.Read'            = 'Application Calendars.Read'
        'Calendars.ReadWrite'       = 'Application Calendars.ReadWrite'
        'Contacts.Read'             = 'Application Contacts.Read'
        'Contacts.ReadWrite'        = 'Application Contacts.ReadWrite'
        'full_access_as_app'        = 'Application EWS.AccessAsApp'
    }
}
