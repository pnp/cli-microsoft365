# aad app set

Updates Azure AD app registration

## Usage

```sh
m365 aad app set [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration to update. Specify either `appId`, `objectId` or `name`

`--objectId [objectId]`
: Object ID of the Azure AD application registration to update. Specify either `appId`, `objectId` or `name`

`--name [name]`
: Name of the Azure AD application registration to update. Specify either `appId`, `objectId` or `name`

`-u, --uri [uri]`
: Comma-separated list of Application ID URIs to update

`-r, --redirectUris [redirectUris]`
: Comma-separated list of redirect URIs to add to the app registration. Requires `platform` to be specified

`-p, --platform [platform]`
: Platform for which the `redirectUri` should be configured. Allowed values `spa,web,publicClient`

`--redirectUrisToRemove [redirectUrisToRemove]`
: Comma-separated list of existing redirect URIs to remove. Specify, when you want to replace existing redirect URIs with another

`--certificateFile [certificateFile]`
: Path to the file with certificate public key. Specify either `certificateFile` or `certificateBase64Encoded`

`--certificateBase64Encoded [certificateBase64Encoded]`
: Base64-encoded string with certificate public key. Specify either `certificateFile` or `certificateBase64Encoded`

`--certificateDisplayName [certificateDisplayName]`
: Display name for the certificate. If not given, the displayName will be set to the certificate subject. When specified, also specify either `certificateFile` or `certificateBase64Encoded`

--8<-- "docs/cmd/_global.md"

## Remarks

For best performance use the `objectId` option to reference the Azure AD application registration to update. If you use `appId` or `name`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

When a certificate is specified it will be added to the list of certificates of the app without changing existing certificates.

## Examples

Update the app URI of the Azure AD application registration specified by its object ID

```sh
m365 aad app set --objectId d75be2e1-0204-4f95-857d-51a37cf40be8 --uri https://contoso.com/e75be2e1-0204-4f95-857d-51a37cf40be8
```

Update the app URI of the Azure AD application registration specified by its app (client) ID

```sh
m365 aad app set --appId e75be2e1-0204-4f95-857d-51a37cf40be8 --uri https://contoso.com/e75be2e1-0204-4f95-857d-51a37cf40be8
```

Update the app URI of the Azure AD application registration specified by its name

```sh
m365 aad app set --name "My app" --uri https://contoso.com/e75be2e1-0204-4f95-857d-51a37cf40be8
```

Add a new redirect URI for SPA authentication

```sh
m365 aad app set --objectId 95cfe30d-ed44-4f9d-b73d-c66560f72e83 --redirectUris https://contoso.com/auth --platform spa
```

Replace one redirect URI with another for SPA authentication

```sh
m365 aad app set --objectId 95cfe30d-ed44-4f9d-b73d-c66560f72e83 --redirectUris https://contoso.com/auth --platform spa --redirectUrisToRemove https://contoso.com/old-auth
```

Add a certificate to the app

```sh
m365 aad app set --appId e75be2e1-0204-4f95-857d-51a37cf40be8 --certificateDisplayName "Some certificate name" --certificateFile "c:\temp\some-certificate.cer"
```
