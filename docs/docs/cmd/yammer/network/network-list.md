# Yammer network list

Returns a list of networks to which the current user has access

## Usage

```sh
m365 yammer network list [options]
```

## Options

`--includeSuspended`
: Include the networks in which the user is suspended

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns the current user's networks

```sh
m365 yammer network list
```

Returns the current user's networks including the networks in which the user is suspended

```sh
m365 yammer network list --includeSuspended
```

## Response

=== "JSON"

    ```json
    [
      {
        "type": "network",
        "id": 5897756673,
        "email": "",
        "name": "Contoso",
        "community": false,
        "permalink": "contoso.onmicrosoft.com",
        "web_url": "https://www.yammer.com/contoso.onmicrosoft.com",
        "show_upgrade_banner": false,
        "header_background_color": "#396B9A",
        "header_text_color": "#FFFFFF",
        "navigation_background_color": "#38699F",
        "navigation_text_color": "#FFFFFF",
        "paid": true,
        "moderated": false,
        "is_freemium": false,
        "is_org_chart_enabled": false,
        "is_group_enabled": true,
        "is_chat_enabled": true,
        "is_translation_enabled": false,
        "created_at": "2020/02/26 10:33:56 +0000",
        "is_storyline_enabled": true,
        "is_storyline_preview_enabled": false,
        "is_stories_enabled": true,
        "is_stories_preview_enabled": false,
        "is_premium_preview_enabled": false,
        "profile_fields_config": {
          "enable_work_history": true,
          "enable_education": true,
          "enable_job_title": true,
          "enable_work_phone": true,
          "enable_mobile_phone": true,
          "enable_summary": true,
          "enable_interests": true,
          "enable_expertise": true,
          "enable_location": true,
          "enable_im": true,
          "enable_skype": true,
          "enable_websites": true
        },
        "browser_deprecation_url": null,
        "external_messaging_state": "disabled",
        "state": "enabled",
        "enforce_office_authentication": true,
        "office_authentication_committed": true,
        "is_gif_shortcut_enabled": true,
        "is_link_preview_enabled": true,
        "attachments_in_private_messages": false,
        "secret_groups": false,
        "force_connected_groups": true,
        "force_spo_files": false,
        "connected_all_company": true,
        "m365_native_mode": true,
        "force_optin_modern_client": false,
        "admin_modern_client_flexible_optin": false,
        "aad_guests_enabled": false,
        "all_company_group_creation_state": null,
        "unseen_message_count": -1,
        "preferred_unseen_message_count": -1,
        "private_unseen_thread_count": 0,
        "inbox_unseen_thread_count": 0,
        "private_unread_thread_count": 0,
        "unseen_notification_count": 0,
        "has_fake_email": false,
        "is_primary": true,
        "allow_attachments": true,
        "attachment_types_allowed": "ALL",
        "privacy_link": "https://go.microsoft.com/fwlink/p/?linkid=857875"
      }
    ]
    ```

=== "Text"

    ```text
    community: false
    email    :
    id       : 5897756673
    name     : Contoso
    permalink: contoso.onmicrosoft.com
    web_url  : https://www.yammer.com/contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    id,name,email,community,permalink,web_url
    5897756673,Contoso,,,contoso.onmicrosoft.com,https://www.yammer.com/contoso.onmicrosoft.com
    ```
