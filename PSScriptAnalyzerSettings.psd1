@{
    # PSScriptAnalyzer settings for the EXORBACforAppManagement module.
    # The build fails only on Error-severity findings; Warning/Information are reported but
    # non-blocking. A few rules that add noise without value here are excluded.
    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        # Private helpers are intentionally undocumented internal functions.
        'PSProvideCommentHelp'
    )
}
