# teams meeting transcript get

Downloads a transcript for a given meeting

## Usage

```sh
m365 teams meeting transcript get [options]
```

## Options

`-u, --userId [userId]`
: The id of the user, omit to get meeting transcript for current signed in user. Use either  `id`, `userName` or `email`, but not multiple.

`-n, --userName [userName]`
: The name of the user, omit to get meeting transcript for current signed in user. Use either `id`, `userName` or `email`, but not multiple.

`--email [email]`
: The email of the user, omit to get meeting transcript reports for current signed in user. Use either `id`, `userName` or `email`, but not multiple.

`-m, --meetingId <meetingId>`
: The Id of the meeting.

`-i, --id <id>`
: The Id of the transcript.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the specified transcript made for the current signed in user and Microsoft Teams meeting with given id.

```sh
m365 teams meeting transcript get --meetingId MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ --id MSMjMCMjNzU3ODc2ZDYtOTcwMi00MDhkLWFkNDItOTE2ZDNmZjkwZGY4
```

Saves the specified transcript made for the _[garthf@contoso.com](mailto:garthf@contoso.com)_ and Microsoft Teams meeting with given id.

```sh
m365 teams meeting transcript get --userName garthf@contoso.com --meetingId MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ --id MSMjMCMjNzU3ODc2ZDYtOTcwMi00MDhkLWFkNDItOTE2ZDNmZjkwZGY4 --outputFile c:/Transcript.vtt
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

## Response

=== "JSON"

    ```json
    {
      "id": "MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj",
      "meetingId": "MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy",
      "meetingOrganizerId": "e1251b10-1ba4-49e3-b35a-933e3f21772b",
      "transcriptContentUrl": "https://graph.microsoft.com/beta/users/e1251b10-1ba4-49e3-b35a-933e3f21772b/onlineMeetings/MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy/transcripts/MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj/content",
      "createdDateTime": "2024-04-08T05:26:21.1936844Z",
      "meetingOrganizer": {
        "application": null,
        "device": null,
        "user": {
          "id": "e1251b10-1ba4-49e3-b35a-933e3f21772b",
          "displayName": null,
          "userIdentityType": "aadUser",
          "tenantId": "de348bc7-1aeb-4406-8cb3-97db021cadb4"
        }
      }
    }
    ```

=== "Text"

    ```text
    createdDateTime     : 2024-04-08T05:26:21.1936844Z
    id                  : MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj
    meetingId           : MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy
    meetingOrganizer    : {"application":null,"device":null,"user":{"id":"e1251b10-1ba4-49e3-b35a-933e3f21772b","displayName":null,"userIdentityType":"aadUser","tenantId":"de348bc7-1aeb-4406-8cb3-97db021cadb4"}}
    meetingOrganizerId  : e1251b10-1ba4-49e3-b35a-933e3f21772b
    transcriptContentUrl: https://graph.microsoft.com/beta/users/e1251b10-1ba4-49e3-b35a-933e3f21772b/onlineMeetings/MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy/transcripts/MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj/content
    ```

=== "CSV"

    ```csv
    id,meetingId,meetingOrganizerId,transcriptContentUrl,createdDateTime
    MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj,MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy,e1251b10-1ba4-49e3-b35a-933e3f21772b,https://graph.microsoft.com/beta/users/e1251b10-1ba4-49e3-b35a-933e3f21772b/onlineMeetings/MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy/transcripts/MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj/content,2024-04-08T05:26:21.1936844Z
    ```

=== "Markdown"

    ```md
    # teams meeting transcript get --meetingId "MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy" --id "MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj"

    Date: 4/9/2024

    ## MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj

    Property | Value
    ---------|-------
    id | MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj
    meetingId | MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy
    meetingOrganizerId | e1251b10-1ba4-49e3-b35a-933e3f21772b
    transcriptContentUrl | https://graph.microsoft.com/beta/users/e1251b10-1ba4-49e3-b35a-933e3f21772b/onlineMeetings/MSplMTI1MWIxMC0xYmE0LTQ5ZTMtYjM1YS05MzNlM2YyMTc3MmIqMCoqMTk6bWVldGluZ19OREJpWVROa05XVXRaakptWlMwMFl6QTRMVGd3TlRRdE16WTNaR014T1Rjek1tUTBAdGhyZWFkLnYy/transcripts/MSMjMCMjNmU2OTc2OTUtZWNmMC00MTE2LWEyNzYtYjcyOTE5NTBiNzRj/content
    createdDateTime | 2024-04-08T05:26:21.1936844Z
    ```
