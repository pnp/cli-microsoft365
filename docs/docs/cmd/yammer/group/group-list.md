# yammer group list

Returns the list of groups in a Yammer network or the groups for a specific user

## Usage

```sh
m365 yammer group list [options]
```

## Options

`--userId [userId]`
: Returns the groups for a specific user

`--limit [limit]`
: Limits the groups returned

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns all Yammer network groups

```sh
m365 yammer group list
```

Returns all Yammer network groups for the user with the ID `5611239081`

```sh
m365 yammer group list --userId 5611239081
```

Returns the first 10 Yammer network groups

```sh
m365 yammer group list --limit 10
```

Returns the first 10 Yammer network groups for the user with the ID `5611239081`

```sh
m365 yammer group list --userId 5611239081 --limit 10
```

## Response

=== "JSON"

    ```json
    [
      {
        "type": "group",
        "id": 31158067201,
        "email": "",
        "full_name": "Contoso Hub",
        "network_id": 5897756673,
        "name": "contosohub",
        "description": "",
        "privacy": "public",
        "url": "https://www.yammer.com/api/v1/groups/31158067201",
        "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/#/threads/inGroup?type=in_group&feedId=31158067201",
        "mugshot_url": "https://mugshot0eu-1.assets-yammer.com/mugshot/images/group_profile.png?P1=1668205176&P2=104&P3=1&P4=l98Wk4FkhCqVX1J8bQ_8yZDbK4cfU1lQGgkK0Ak1k2g-tfLV9_ecm6k7FyFApCq3Xnzl7NPKpGLWT2IVD-Ft5q3VSCwzv5c0A1l-SFC5MrfN25BIsR9ux8K-LlYbFUF3yeh-vFk_IxwE-AI2xEVCuq0aoINzHiIW4Gi5IxC6mDDni72sE2LuM3X4LooEowEYrzfz5d-m9hMveU1E8KPPEmq3WTejhJ_Bc3zY3XA3n4jEPDnZ09uPUyVCBpa84Ysh-GGSkFWsPBAldAQAbbzcjip_SzrfKz868BolCLlbM3DwRQfyDH9Of9IYEZpu1U85hBuNoolF68rKPVL6-bxl2w&size=48x48",
        "mugshot_redirect_url": "https://www.yammer.com/mugshot/images/redirect/48x48/group_profile.png",
        "mugshot_url_template": "https://mugshot0eu-1.assets-yammer.com/mugshot/images/group_profile.png?P1=1668205176&P2=104&P3=1&P4=l98Wk4FkhCqVX1J8bQ_8yZDbK4cfU1lQGgkK0Ak1k2g-tfLV9_ecm6k7FyFApCq3Xnzl7NPKpGLWT2IVD-Ft5q3VSCwzv5c0A1l-SFC5MrfN25BIsR9ux8K-LlYbFUF3yeh-vFk_IxwE-AI2xEVCuq0aoINzHiIW4Gi5IxC6mDDni72sE2LuM3X4LooEowEYrzfz5d-m9hMveU1E8KPPEmq3WTejhJ_Bc3zY3XA3n4jEPDnZ09uPUyVCBpa84Ysh-GGSkFWsPBAldAQAbbzcjip_SzrfKz868BolCLlbM3DwRQfyDH9Of9IYEZpu1U85hBuNoolF68rKPVL6-bxl2w&size={width}x{height}",
        "mugshot_redirect_url_template": "https://www.yammer.com/mugshot/images/redirect/{width}x{height}/group_profile.png",
        "mugshot_id": null,
        "show_in_directory": "true",
        "created_at": "2022/11/11 20:54:52 +0000",
        "aad_guests": 0,
        "color": "#2c5b85",
        "external": false,
        "moderated": false,
        "header_image_url": "https://mugshot0eu-1.assets-yammer.com/mugshot/images/group-header-coffee.png?P1=1668204451&P2=104&P3=1&P4=hPZP6QJbY1Oj4KQZAodyMQyvjUahlwoqSCMqioVYvDoB-9Fx3qEB3ZTM7I_TF-mceKqGVDtasUIH8ZDYEfjTg9zgWWDpmkREJySioTZ0WcPtHIUkh2GUWJOfr-5aX9QhdpE1Fpp94mltGCtBc_nqlEbgIAYCJtBKgLAgUFZ4L2WSkQNn5Y_JLp5cM9Gnf7Z3MmHniN0Na1oemDhZ1vOsGCtaU09WPB5oNoSUMfwqYSKjF5IqXdd55Y3F2NZuuyTHoZS65BFZR9OJaICXJs6Q2dNExLqMvGQ76_aZsgli-BG67MVwfDsmqpxsjZZOBIZGQOEKc4D_bx8iQUHZD7p2xA",
        "category": "unclassified",
        "default_thread_starter_type": "normal",
        "restricted_posting": false,
        "company_group": false,
        "creator_type": "user",
        "creator_id": 36425097217,
        "state": "active",
        "stats": {
          "members": 1,
          "aad_guests": 0,
          "updates": 0,
          "last_message_id": null,
          "last_message_at": null
        }
      }
    ]
    ```

=== "Text"

    ```text
    email    :
    external : false
    id       : 31158067201
    moderated: false
    name     : contosohub
    privacy  : public
    ```

=== "CSV"

    ```csv
    id,name,email,privacy,external,moderated
    31158067201,wombathub,,public,false,false
    ```
