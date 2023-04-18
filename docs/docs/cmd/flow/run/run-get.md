# flow run get

Gets information about a specific run of the specified Power Automate flow

## Usage

```sh
m365 flow run get [options]
```

## Options

`-n, --name <name>`
: The name of the run to get information about

`-f, --flowName <flowName>`
: The name of the Power Automate flow for which to retrieve information

`-e, --environmentName <environmentName>`
: The name of the environment where the flow is located

`--includeTriggerInformation`
: If specified, include information about the trigger details

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

If the run with the name you specified doesn't exist, you will get the `The provided workflow run name is not valid.` error.

If the option `includeTriggerInformation` is specified, but the trigger does not contain an outputsLink such as for example with a `Recurrence` trigger, this option will be ignored.

## Examples

Get information about the given run of the specified Power Automate flow

```sh
m365 flow run get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62
```

Get information about the given run of the specified Power Automate flow including trigger information

```sh
m365 flow run get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62 --includeTriggerInformation
```

## Response

### Standard response

=== "JSON"

    ```json
    {
      "name": "08585329112602833828909892130CU17",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-de348bc7-1aeb-4406-8cb3-97db021cadb4/flows/170fb67e-a514-4d84-8727-582022bd13a9/runs/08585329112602833828909892130CU17",
      "type": "Microsoft.ProcessSimple/environments/flows/runs",
      "properties": {
        "startTime": "2022-11-17T14:33:45.2763872Z",
        "status": "Running",
        "correlation": {
          "clientTrackingId": "08585329112602833829909892130CU00"
        },
        "trigger": {
          "name": "When_a_new_response_is_submitted",
          "inputsLink": {
            "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329112602833828909892130CU17/contents/TriggerInputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329112602833828909892130CU17%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=",
            "contentVersion": "6ZrBBE+MJg7IvhMgyJLMmA==",
            "contentSize": 349,
            "contentHash": {
              "algorithm": "md5",
              "value": "6ZrBBE+MJg7IvhMgyJLMmA=="
            }
          },
          "outputsLink": {
            "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329112602833828909892130CU17/contents/TriggerOutputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329112602833828909892130CU17%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=",
            "contentVersion": "Z/4a8tfYygNAR1xpc44iww==",
            "contentSize": 493,
            "contentHash": {
              "algorithm": "md5",
              "value": "Z/4a8tfYygNAR1xpc44iww=="
            }
          },
          "startTime": "2022-11-17T14:33:45.1914506Z",
          "endTime": "2022-11-17T14:33:45.1914506Z",
          "originHistoryName": "08585329112602833829909892130CU00",
          "correlation": {
            "clientTrackingId": "08585329112602833829909892130CU00"
          },
          "status": "Succeeded"
        }
      },
      "startTime": "2022-11-17T14:33:45.2763872Z",
      "endTime": "",
      "status": "Running",
      "triggerName": "When_a_new_response_is_submitted"
    }
    ```

=== "Text"

    ```text
    endTime    : 2023-03-04T09:05:22.5880202Z
    name       : 08585236861638480597867166179CU104
    startTime  : 2023-03-04T09:05:21.8066368Z
    status     : Succeeded
    triggerName: When_an_email_is_flagged_(V4)
    ```

=== "CSV"

    ```csv
    name,startTime,endTime,status,triggerName
    08585236861638480597867166179CU104,2023-03-04T09:05:21.8066368Z,2023-03-04T09:05:22.5880202Z,Succeeded,When_an_email_is_flagged_(V4)
    ```

=== "Markdown"

    ```md
    # flow run get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62

    Date: 04/03/2023

    ## 08586653536760200319026785874CU62 (/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/5923cb07-ce1a-4a5c-ab81-257ce820109a/runs/08586653536760200319026785874CU62)

    Property | Value
    ---------|-------
    name | 08586653536760200319026785874CU62
    id | /providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/c3c707b5-fefd-4f7a-a96c-b8e0d5ca3cc1/runs/08585208964855963748594654409CU47
    type | Microsoft.ProcessSimple/environments/flows/runs
    properties | {"startTime":"2023-04-05T15:59:59.8822066Z","endTime":"2023-04-05T16:00:02.3071033Z","status":"Succeeded","correlation":{"clientTrackingId":"08585208964855963748594654409CU47"},"trigger":{"name":"Recurrence","startTime":"2023-04-05T15:59:59.8696099Z","endTime":"2023-04-05T15:59:59.8696099Z","scheduledTime":"2023-04-05T16:00:00Z","originHistoryName":"08585208964855963748594654409CU47","correlation":{"clientTrackingId":"08585208964855963748594654409CU47"},"code":"OK","status":"Succeeded"}}
    startTime | 2023-03-04T09:05:21.8066368Z
    endTime | 2023-03-04T09:05:22.5880202Z
    status | Succeeded
    triggerName | When\_an\_email\_is\_flagged\_(V4)
    ```

### `includeTriggerInformation` response

When using the option `includeTriggerInformation`, the response for the json and md-output will differ.

=== "JSON"

    ```json
    {
      "name": "08585236861638480597867166179CU104",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/24335774-daf6-4183-acb7-f5155c2cd2fe/runs/08585236861638480597867166179CU104",
      "type": "Microsoft.ProcessSimple/environments/flows/runs",
      "properties": {
        "startTime": "2023-03-04T09:05:21.8066368Z",
        "endTime": "2023-03-04T09:05:22.5880202Z",
        "status": "Succeeded",
        "correlation": {
          "clientTrackingId": "08585236861638480598867166179CU131"
        },
        "trigger": {
          "name": "When_an_email_is_flagged_(V4)",
          "inputsLink": {
            "uri": "https://prod-130.westeurope.logic.azure.com:443/workflows/3ebadb794f6641e0b7f4fda131cdfb0b/runs/08585236861638480597867166179CU104/contents/TriggerInputs?api-version=2016-06-01&se=2023-03-04T14%3A00%3A00.0000000Z&sp=%2Fruns%2F08585236861638480597867166179CU104%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=",
            "contentVersion": "2v/VLXFrKV6JvwSdcN7aHg==",
            "contentSize": 343,
            "contentHash": {
              "algorithm": "md5",
              "value": "2v/VLXFrKV6JvwSdcN7aHg=="
            }
          },
          "outputsLink": {
            "uri": "https://prod-130.westeurope.logic.azure.com:443/workflows/3ebadb794f6641e0b7f4fda131cdfb0b/runs/08585236861638480597867166179CU104/contents/TriggerOutputs?api-version=2016-06-01&se=2023-03-04T14%3A00%3A00.0000000Z&sp=%2Fruns%2F08585236861638480597867166179CU104%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=",
            "contentVersion": "AHZEeWNlQ0bLe48yDmpzrQ==",
            "contentSize": 3478,
            "contentHash": {
              "algorithm": "md5",
              "value": "AHZEeWNlQ0bLe48yDmpzrQ=="
            }
          },
          "startTime": "2023-03-04T09:05:21.6192576Z",
          "endTime": "2023-03-04T09:05:21.7442626Z",
          "scheduledTime": "2023-03-04T09:05:21.573561Z",
          "originHistoryName": "08585236861638480598867166179CU131",
          "correlation": {
            "clientTrackingId": "08585236861638480598867166179CU131"
          },
          "code": "OK",
          "status": "Succeeded"
        }
      },
      "startTime": "2023-03-04T09:05:21.8066368Z",
      "endTime": "2023-03-04T09:05:22.5880202Z",
      "status": "Succeeded",
      "triggerName": "When_an_email_is_flagged_(V4)",
      "triggerInformation": {
        "from": "john@contoso.com",
        "toRecipients": "doe@contoso.com",
        "subject": "Dummy email",
        "body": "<html><head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body><p>This is dummy content</p></body></html>",
        "importance": "normal",
        "bodyPreview": "This is dummy content",
        "hasAttachments": false,
        "id": "AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E_hLMK5kAALUqy81AAA=",
        "internetMessageId": "<DB7PR03MB5018879914324FC65695809FE1AD9@DB7PR03MB5018.eurprd03.prod.outlook.com>",
        "conversationId": "AAQkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgAQAMqP9zsK8a1CnIYEgHclLTk=",
        "receivedDateTime": "2023-03-01T15:06:57+00:00",
        "isRead": false,
        "attachments": [],
        "isHtml": true
      }
    }
    ```

=== "Markdown"

    ```md
    # flow run get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62

    Date: 04/03/2023

    ## 08586653536760200319026785874CU62 (/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/5923cb07-ce1a-4a5c-ab81-257ce820109a/runs/08586653536760200319026785874CU62)

    Property | Value
    ---------|-------
    name | 08586653536760200319026785874CU62
    id | /providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/c3c707b5-fefd-4f7a-a96c-b8e0d5ca3cc1/runs/08585208964855963748594654409CU47
    type | Microsoft.ProcessSimple/environments/flows/runs
    properties | {"startTime":"2023-04-05T15:59:59.8822066Z","endTime":"2023-04-05T16:00:02.3071033Z","status":"Succeeded","correlation":{"clientTrackingId":"08585208964855963748594654409CU47"},"trigger":{"name":"Recurrence","startTime":"2023-04-05T15:59:59.8696099Z","endTime":"2023-04-05T15:59:59.8696099Z","scheduledTime":"2023-04-05T16:00:00Z","originHistoryName":"08585208964855963748594654409CU47","correlation":{"clientTrackingId":"08585208964855963748594654409CU47"},"code":"OK","status":"Succeeded"}}
    startTime | 2023-03-04T09:05:21.8066368Z
    endTime | 2023-03-04T09:05:22.5880202Z
    status | Succeeded
    triggerName | When\_an\_email\_is\_flagged\_(V4)
    triggerInformation | {"from":"noreply-capmarketrevenues@creg.be","toRecipients":"mathijs@mathijsdev2.onmicrosoft.com","subject":"Validation debtor by CREG: Debtor Mathijs 100000000","body":"<html><head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body><p>CREG has validated the debtor with the following details:<br><br>Name of the debtor: Debtor Mathijs<br>Legal form: nv<br>VAT number: BE0123321111<br>Street + number: Straat a<br>City: Vosselaar<br>Postal code: 2350<br>First and last name of the contact: &nbsp;Mathijs Verbeeck<br>Email: mathijs@mathijsdev2.onmicrosoft.com<br>Telephone number of the contact: +32476345130<br><br>If you believe you received this email in error, please contact us by sending an email to capmarketrevenues@creg.be.<br><br>Please do not reply to this message. This email address is not monitored so there will be no response to any messages sent to this address.<br><br>Thank you,<br>&nbsp;<br>CREG<br></p></body></html>","importance":"normal","bodyPreview":"CREG has validated the debtor with the following details:\r\n\r\nName of the debtor: Debtor Mathijs\r\nLegal form: nv\r\nVAT number: BE0123321111\r\nStreet + number: Straat a\r\nCity: Vosselaar\r\nPostal code: 2350\r\nFirst and last name of the contact:  Mathijs Verbeeck","hasAttachments":false,"id":"AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E\_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E\_hLMK5kAALUqy81AAA=","internetMessageId":"<DB7PR03MB5018879914324FC65695809FE1AD9@DB7PR03MB5018.eurprd03.prod.outlook.com>","conversationId":"AAQkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgAQAMqP9zsK8a1CnIYEgHclLTk=","receivedDateTime":"2023-03-01T15:06:57+00:00","isRead":true,"attachments":[],"isHtml":true}
    ```
