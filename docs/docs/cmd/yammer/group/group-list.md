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

=== "Markdown"

    ```md
    # yammer group list --limit "1"

    Date: 2023-05-16

    ## leadership (123412865024)

    Property | Value
    ---------|-------
    type | group
    id | 123412865024
    email | leadership+contoso.onmicrosoft.com@yammer.com
    full\_name | Leadership
    network\_id | 98327945216
    name | leadership
    description | Share what's on your mind and get important announcements from Patti and the rest of the Leadership Team.
    privacy | public
    url | https://www.yammer.com/api/v1/groups/123412865024
    web\_url | https://www.yammer.com/contoso.onmicrosoft.com/#/threads/inGroup?type=in\_group&feedId=123412865024
    mugshot\_url | https://mugshot0.assets-yammer.com/mugshot/images/5jjCjcSTJsdzFn0Ps50Vz0tqNdWdgnWs?P1=1684267104&P2=104&P3=1&P4=igdO9ZCQbSd5YS7tzwuIFj9CmMPsPWWpAjsk0xGDGrciD-3XKKsHbYx-e6H22yZ6OqLc3zt\_5ZOWefd8l537cWNUOPzeDg2lz\_fNxx1bowFMIdz6mRCHcCwygEwtKI0HxX5eHd4cdJBg54c4R6VN1\_Oex7Ug9Are6hVux4DsLg7eoNMMYvvcjXUp2zcT7o6bXYcZM2WBf\_r1IC24Sb-PLaSfAtKJZsswBBTkmz\_B7O5PZFcY4TQJvd5XzwEL17aqWm1hV1MCUSEd3Ms7Clc7KwxA0Hhv1rWYF064siAHEDiVlKZrE1yN7j-gCt0K1\_xUHWc54TrUIjFxnrwMDGZvzw&size=48x48
    mugshot\_redirect\_url | https://www.yammer.com/mugshot/images/redirect/48x48/5jjCjcSTJsdzFn0Ps50Vz0tqNdWdgnWs
    mugshot\_url\_template | https://mugshot0.assets-yammer.com/mugshot/images/5jjCjcSTJsdzFn0Ps50Vz0tqNdWdgnWs?P1=1684267104&P2=104&P3=1&P4=igdO9ZCQbSd5YS7tzwuIFj9CmMPsPWWpAjsk0xGDGrciD-3XKKsHbYx-e6H22yZ6OqLc3zt\_5ZOWefd8l537cWNUOPzeDg2lz\_fNxx1bowFMIdz6mRCHcCwygEwtKI0HxX5eHd4cdJBg54c4R6VN1\_Oex7Ug9Are6hVux4DsLg7eoNMMYvvcjXUp2zcT7o6bXYcZM2WBf\_r1IC24Sb-PLaSfAtKJZsswBBTkmz\_B7O5PZFcY4TQJvd5XzwEL17aqWm1hV1MCUSEd3Ms7Clc7KwxA0Hhv1rWYF064siAHEDiVlKZrE1yN7j-gCt0K1\_xUHWc54TrUIjFxnrwMDGZvzw&size={width}x{height}
    mugshot\_redirect\_url\_template | https://www.yammer.com/mugshot/images/redirect/{width}x{height}/5jjCjcSTJsdzFn0Ps50Vz0tqNdWdgnWs
    mugshot\_id | 5jjCjcSTJsdzFn0Ps50Vz0tqNdWdgnWs
    show\_in\_directory | true
    created\_at | 2022/12/12 12:51:11 +0000
    aad\_guests | 0
    color | #0e4f7a
    external | false
    moderated | false
    header\_image\_url | https://mugshot0.assets-yammer.com/mugshot/images/group-header-megaphone.png?P1=1684266783&P2=104&P3=1&P4=FObDxfvTV7O201-7u4v-u4Y25mAZNrpD9QhUqSXbUyC8UaqvGJH7mT5yPtx0Qls\_QUkM3606i0F2GnkQHOwC1tVW8Vse0yNZHWDTyqA\_wSRX\_fn6cP47uoC4wvSsGAmWeb6epr-hJpDW\_qn-1CHQF7cen2Ti9Ap-XncmOiu2Tfd2DTuGyuHKivI6cxGGbIQ5ERU1NgiVEXqKClOMb9qPUBu4dqPc1gfaFDaA1umUslwTG3DRfAIVviECiG1eHI5cjkTX5qifscUXCmEOQU5lLih9J409qVUOPa0vs1clNspm6XtkVaAfC8FB2gaBmEqbVtFBVbAwyoUJhu2KM0Vp7w
    category | unclassified
    default\_thread\_starter\_type | normal
    restricted\_posting | false
    company\_group | false
    creator\_type | user
    creator\_id | 1842176974848
    state | active
    ```
