# yammer user list

Returns users from the current network

## Usage

```sh
m365 yammer user list [options]
```

## Options

`-g, --groupId [groupId]`
: Returns users within a given group

`-l, --letter [letter]`
: Returns users with usernames beginning with the given character

`--reverse`
: Returns users in reverse sorting order

`--limit [limit]`
: Limits the users returned

`--sortBy [sortBy]`
: Returns users sorted by a number of messages or followers, instead of the default behavior of sorting alphabetically. Allowed values are `messages,followers`

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples
  
Returns all Yammer network users

```sh
m365 yammer user list
```

Returns all Yammer network users with usernames beginning with "a"

```sh
m365 yammer user list --letter a
```

Returns all Yammer network users sorted alphabetically in descending order

```sh
m365 yammer user list --reverse
```

Returns the first 10 Yammer network users within the group 5785177

```sh
m365 user list --groupId 5785177 --limit 10
```

## Response

=== "JSON"

    ```json
    [
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
        "full_name": "John Doe",
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
        "name": "johndoe",
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
        "show_invite_lightbox": false
      }
    ]
    ```

=== "Text"

    ```text
    id            full_name         email
    ------------  ----------------  -------------------------------
    36425097217   John Doe          johndoe@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    id,full_name,email
    36425097217,John Doe,johndoe@contoso.onmicrosoft.com
    ```
