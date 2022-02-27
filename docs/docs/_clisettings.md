## Available settings

Following is the list of configuration settings available in CLI for Microsoft 365.

Setting name|Definition|Default value
------------|----------|-------------
`autoOpenBrowserOnLogin`|Automatically open the browser to the Azure AD login page after running `m365 login` command in device code mode|`false`
`csvEscape`|Single character used for escaping; only apply to characters matching the quote and the escape options|`"`
`csvHeader`|Display the column names on the first line|`true`
`csvQuote`|The quote characters surrounding a field. An empty quote value will preserve the original field, whether it contains quotation marks or not.|` `
`csvQuoted`|Quote all the non-empty fields even if not required|`false`
`csvQuotedEmpty`|Quote empty strings and overrides quoted_string on empty strings when defined|`false`
`errorOutput`|Defines if errors should be written to `stdout` or `stderr`|`stderr`
`output`|Defines the default output when issuing a command|`json`
`printErrorsAsPlainText`|When output mode is set to `json`, print error messages as plain-text rather than JSON|`true`
`showHelpOnFailure`|Automatically display help when executing a command failed|`true`
