# TODO

[x] convert [ApplicationAccessPolices](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-access-policies) entries into RBAC for application (Convert-ApplicationAccessPolicyToRBAC)

[x] by design get-RBACforAppEntry should show only RoleAssigneeType -eq 'ServicePrincipal' and (RecipientScope -eq Group) or (RecipientScope -eq CustomRecipientScope))

[x] add to results of  New-RBACforAppEntry  and New-RBACforAppUnifiedGroup name of used owner (both requested, and added similarly to member)
