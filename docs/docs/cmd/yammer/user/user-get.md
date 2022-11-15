# yammer user get

Retrieves the current user or searches for a user by ID or e-mail

## Usage

```sh
m365 yammer user get [options]
```

## Options

`-i, --id [id]`
: Retrieve a user by ID

`--email [email]`
: Retrieve a user by e-mail

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

All operations return a single user object. Operations executed with the `email` parameter return an array of user objects.

## Examples
  
Returns the current user

```sh
m365 yammer user get
```

Returns the user with the ID 1496550697

```sh
m365 yammer user get --id 1496550697
```

Returns an array of users matching the e-mail john.smith@contoso.com

```sh
m365 yammer user get --email john.smith@contoso.com
```

Returns an array of users matching the e-mail john.smith@contoso.com in JSON. The JSON output returns a full user object

```sh
m365 yammer user get --email john.smith@contoso.com --output json
```

## Response

=== "JSON"

    ```json
    {
      "type": "user",
      "id": 172006440961,
      "network_id": 5897756673,
      "state": "active",
      "job_title": "",
      "location": null,
      "interests": null,
      "summary": null,
      "expertise": null,
      "full_name": "johndoe",
      "activated_at": "2021/10/08 11:45:32 +0000",
      "auto_activated": false,
      "show_ask_for_photo": true,
      "first_name": "",
      "last_name": "",
      "network_name": "Contoso",
      "network_domains": [
        "contoso.onmicrosoft.com"
      ],
      "url": "https://www.yammer.com/api/v1/users/172006440961",
      "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/users/172006440961",
      "name": "admvalo",
      "mugshot_url": "https://mugshot0eu-1.assets-yammer.com/mugshot/images/no_photo.png?P1=1668205841&P2=104&P3=1&P4=rnM2FPGOQRZja018qxAFshyyDKKH5SoUXkdpCeizRsdD7Ggb9sLSsdEaq-icgk8g-QTHFd0Te4e1gWAZTGEQekSQop6G6zDcipIVbZMJStEzfKKUpSPckcXnRhfiI55yq5AOLhVcH2PP_ZBFF-0KXMaP8Hy4dGDIRzmnUGhFuik0yjNBoGaYL86ltEaDMQdpS6rS3lmIMzLPGEMfr30vethAxRT7SKBbNYxZ9iPxO6TY26cYCfv0VyyMQkGGviPU4__EVjOoklhD_AqFGFGHtRTcsafpKOxCE70Z-nUpIPbYCel3las7w105u4SvPPC00Q5LUMDynUvzPiR4-vbWPg&size=48x48",
      "mugshot_redirect_url": "https://www.yammer.com/mugshot/images/redirect/48x48/no_photo.png",
      "mugshot_url_template": "https://mugshot0eu-1.assets-yammer.com/mugshot/images/no_photo.png?P1=1668205841&P2=104&P3=1&P4=rnM2FPGOQRZja018qxAFshyyDKKH5SoUXkdpCeizRsdD7Ggb9sLSsdEaq-icgk8g-QTHFd0Te4e1gWAZTGEQekSQop6G6zDcipIVbZMJStEzfKKUpSPckcXnRhfiI55yq5AOLhVcH2PP_ZBFF-0KXMaP8Hy4dGDIRzmnUGhFuik0yjNBoGaYL86ltEaDMQdpS6rS3lmIMzLPGEMfr30vethAxRT7SKBbNYxZ9iPxO6TY26cYCfv0VyyMQkGGviPU4__EVjOoklhD_AqFGFGHtRTcsafpKOxCE70Z-nUpIPbYCel3las7w105u4SvPPC00Q5LUMDynUvzPiR4-vbWPg&size={width}x{height}",
      "mugshot_redirect_url_template": "https://www.yammer.com/mugshot/images/redirect/{width}x{height}/no_photo.png",
      "birth_date": "",
      "birth_date_complete": "",
      "timezone": "Pacific Time (US & Canada)",
      "external_urls": [],
      "admin": "true",
      "verified_admin": "true",
      "m365_yammer_admin": "false",
      "supervisor_admin": "false",
      "o365_tenant_admin": "true",
      "can_broadcast": "true",
      "department": null,
      "email": "johndoe@contoso.onmicrosoft.com",
      "guest": false,
      "aad_guest": false,
      "can_view_delegations": false,
      "can_create_new_network": false,
      "can_browse_external_networks": false,
      "reaction_accent_color": "none",
      "significant_other": "",
      "kids_names": "",
      "previous_companies": [],
      "schools": [],
      "contact": {
        "im": {
          "provider": "",
          "username": ""
        },
        "phone_numbers": [],
        "email_addresses": [
          {
            "type": "primary",
            "address": "johndoe@contoso.onmicrosoft.com"
          }
        ],
        "has_fake_email": false
      },
      "stats": {
        "updates": 0,
        "following": 0,
        "followers": 0
      },
      "settings": {
        "xdr_proxy": ""
      },
      "show_invite_lightbox": false,
      "age_bucket": "notSet"
    }
    ```

=== "Text"

    ```text
    email    : johndoe@contoso.onmicrosoft.com
    full_name: johndoe
    id       : 172006440961
    job_title:
    state    : active
    url      : https://www.yammer.com/api/v1/users/172006440961
    ```

=== "CSV"

    ```csv
    id,full_name,email,job_title,state,url
    172006440961,admvalo,johndoe@contoso.onmicrosoft.com,,active,https://www.yammer.com/api/v1/users/172006440961
    ```
