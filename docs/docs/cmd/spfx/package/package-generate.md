# spfx package generate

Generates SharePoint Framework solution package with a no-framework web part rendering the specified HTML snippet

## Usage

```sh
m365 spfx package generate
```

## Options

`-t, --webPartTitle <webPartTitle>`
: Title of the web part to generate. Displayed in the tool box when adding web part to page

`-d, --webPartDescription <webPartDescription>`
: Description of the web part to generate. Displayed in the tool box when adding web part to page

`-n, --packageName <packageName>`
: Name of the package to generate. Used among others for the .sppkg file. Must be unique in the app catalog to avoid collisions with other solutions.

`--html <html>`
: HTML snippet to embed in the web part. Can contain `<script>` and `<style>` tags.

`--enableForTeams [enableForTeams]`
: Specify, to make the generated web part available in Microsoft Teams. Specify `tab` to make the web part available as a configurable tab, `personalTab` to make it available as a personal tab or `all` to make it available both as a configurable and personal tab. By default the web part will not be available in Microsoft Teams.

`--exposePageContextGlobally`
: Set, to make the `legacyPageContext` exposed by SharePoint Framework available at `window._spPageContextInfo` for use in the HTML snippet of the web part

`--exposeTeamsContextGlobally`
: Set, to make the Microsoft Teams context exposed by SharePoint Framework available at `window._teamsContextInfo` for use in the HTML snippet of the web part

`--allowTenantWideDeployment`
: Set, to allow the solution package to be deployed globally to all sites

`--developerName [developerName]`
: Name of your organization. Displayed in Microsoft Teams when adding the solution as a (personal) tab. If not specified set to `Contoso`.

`--developerPrivacyUrl [developerPrivacyUrl]`
: URL of the privacy policy for this solution. Displayed in Microsoft Teams when adding the solution as a (personal) tab. If not specified, set to `https://contoso.com/privacy`.

`--developerTermsOfUseUrl [developerTermsOfUseUrl]`
: URL of the terms of use for this solution. Displayed in Microsoft Teams when adding the solution as a (personal) tab. If not specified, set to `https://contoso.com/terms-of-use`.

`--developerWebsiteUrl [developerWebsiteUrl]`
: URL of your organization's website. Displayed in Microsoft Teams when adding the solution as a (personal) tab. If not specified, set to `https://contoso.com/my-app`.

`--developerMpnId [developerMpnId]`
: Microsoft Partner Network ID of your organization. If not specified, set to `000000`.

## Examples

Generate a web part that shows the weather for Amsterdam. Load web part contents from a local file. Allow the web part to be deployed to all sites. Expose the web part in Teams as a personal tab.

```sh
m365 spfx package generate --webPartTitle "Amsterdam weather" --webPartDescription "Shows weather in Amsterdam" --packageName amsterdam-weather --html @amsterdam-weather.html --allowTenantWideDeployment --enableForTeams all
```
