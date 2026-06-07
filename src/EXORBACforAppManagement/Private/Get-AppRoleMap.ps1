function Get-AppRoleMap {
    [OutputType([hashtable])]
    param()

    return [ordered]@{
        'Application Mail.Read'                = 'AppMailR'
        'Application Mail.ReadBasic'           = 'AppMailRBasic'
        'Application Mail.ReadWrite'           = 'AppMailRW'
        'Application Mail.Send'                = 'AppMailSend'
        'Application MailboxSettings.Read'     = 'AppMbxSetR'
        'Application MailboxSettings.ReadWrite'= 'AppMbxSetRW'
        'Application Calendars.Read'           = 'AppCldR'
        'Application Calendars.ReadWrite'      = 'AppCldRW'
        'Application Contacts.Read'            = 'AppCntR'
        'Application Contacts.ReadWrite'       = 'AppCntRW'
        'Application MailboxFolder.Read'       = 'AppMbxFldR'
        'Application MailboxFolder.ReadWrite'  = 'AppMbxFldRW'
        'Application MailboxItem.Read'         = 'AppMbxItemR'
        'Application MailboxItem.ImportExport' = 'AppMbxItemIE'
        'Application Mail Full Access'         = 'AppMailFull'
        'Application Exchange Full Access'     = 'AppExFull'
        'Application EWS.AccessAsApp'          = 'AppEWSAAA'
        'Application SMTP.SendAsApp'           = 'AppSMTPSendAA'
    }
}
