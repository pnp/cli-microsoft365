# flow run list

Lists runs of the specified Microsoft Flow

## Usage

```sh
m365 flow run list [options]
```

## Options

`-f, --flowName <flowName>`
: The name of the Microsoft Flow to retrieve the runs for

`-e, --environmentName <environmentName>`
: The name of the environment to which the flow belongs

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

## Examples

List runs of the specified Microsoft Flow

```sh
m365 flow run list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a
```

## Response

### Standard response

=== "JSON"

    ``` json
    [      
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
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329112602833828909892130CU17/contents/TriggerInputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329112602833828909892130CU17%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=jmdMRWvY7uGoxTmqd3_a2bJtegXuVyuKTKKUVLiwh38",
              "contentVersion": "6ZrBBE+MJg7IvhMgyJLMmA==",
              "contentSize": 349,
              "contentHash": {
                "algorithm": "md5",
                "value": "6ZrBBE+MJg7IvhMgyJLMmA=="
              }
            },
            "outputsLink": {
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329112602833828909892130CU17/contents/TriggerOutputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329112602833828909892130CU17%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=Y3qqjuWrrcQJrmF7uvm6LVzQy5w-dNOFWJ8Yt8khXvA",
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
        "status": "Running"
      },
      {
        "name": "08585329113604313818196774173CU24",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-de348bc7-1aeb-4406-8cb3-97db021cadb4/flows/170fb67e-a514-4d84-8727-582022bd13a9/runs/08585329113604313818196774173CU24",
        "type": "Microsoft.ProcessSimple/environments/flows/runs",
        "properties": {
          "startTime": "2022-11-17T14:32:05.132289Z",
          "endTime": "2022-11-17T14:32:29.2551547Z",
          "status": "Cancelled",
          "correlation": {
            "clientTrackingId": "08585329113604313819196774173CU15"
          },
          "trigger": {
            "name": "When_a_new_response_is_submitted",
            "inputsLink": {
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329113604313818196774173CU24/contents/TriggerInputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329113604313818196774173CU24%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=4gwAiDrqI-GUi57SkW_uX0D7yryuMV5rjRh51TarmUQ",
              "contentVersion": "6ZrBBE+MJg7IvhMgyJLMmA==",
              "contentSize": 349,
              "contentHash": {
                "algorithm": "md5",
                "value": "6ZrBBE+MJg7IvhMgyJLMmA=="
              }
            },
            "outputsLink": {
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329113604313818196774173CU24/contents/TriggerOutputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329113604313818196774173CU24%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=ue-KemAjqZOD2RqGNvsNVNNnn2sO7xmQAH5HTfhfAss",
              "contentVersion": "ATVPdF/i0gBgMtrlHuygFQ==",
              "contentSize": 493,
              "contentHash": {
                "algorithm": "md5",
                "value": "ATVPdF/i0gBgMtrlHuygFQ=="
              }
            },
            "startTime": "2022-11-17T14:32:05.0397953Z",
            "endTime": "2022-11-17T14:32:05.0397953Z",
            "originHistoryName": "08585329113604313819196774173CU15",
            "correlation": {
              "clientTrackingId": "08585329113604313819196774173CU15"
            },
            "status": "Succeeded"
          }
        },
        "startTime": "2022-11-17T14:32:05.132289Z",
        "status": "Cancelled"
      },
      {
        "name": "08585329114263413369372385226CU17",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-de348bc7-1aeb-4406-8cb3-97db021cadb4/flows/170fb67e-a514-4d84-8727-582022bd13a9/runs/08585329114263413369372385226CU17",
        "type": "Microsoft.ProcessSimple/environments/flows/runs",
        "properties": {
          "startTime": "2022-11-17T14:30:59.2453442Z",
          "endTime": "2022-11-17T14:30:59.9036291Z",
          "status": "Failed",
          "code": "ActionFailed",
          "error": {
            "code": "ActionFailed",
            "message": "An action failed. No dependent actions succeeded."
          },
          "correlation": {
            "clientTrackingId": "08585329114263413370372385226CU09"
          },
          "trigger": {
            "name": "When_a_new_response_is_submitted",
            "inputsLink": {
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329114263413369372385226CU17/contents/TriggerInputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329114263413369372385226CU17%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=VRtGkx4IBBNBJ5k0pxp38Gxpjkyyb6P58AKmRCCmrb4",
              "contentVersion": "6ZrBBE+MJg7IvhMgyJLMmA==",
              "contentSize": 349,
              "contentHash": {
                "algorithm": "md5",
                "value": "6ZrBBE+MJg7IvhMgyJLMmA=="
              }
            },
            "outputsLink": {
              "uri": "https://prod-08.centralindia.logic.azure.com:443/workflows/f7bf8f6b5c494e63bfc21b54087a596e/runs/08585329114263413369372385226CU17/contents/TriggerOutputs?api-version=2016-06-01&se=2022-11-17T18%3A00%3A00.0000000Z&sp=%2Fruns%2F08585329114263413369372385226CU17%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=0BlwV_RhWvTyjdcrHtlF-ETJHijX3aZiGdYCs17d6vU",
              "contentVersion": "9SJL6YJRooxAGionZENmtA==",
              "contentSize": 493,
              "contentHash": {
                "algorithm": "md5",
                "value": "9SJL6YJRooxAGionZENmtA=="
              }
            },
            "startTime": "2022-11-17T14:30:59.1292588Z",
            "endTime": "2022-11-17T14:30:59.1292588Z",
            "originHistoryName": "08585329114263413370372385226CU09",
            "correlation": {
              "clientTrackingId": "08585329114263413370372385226CU09"
            },
            "status": "Succeeded"
          }
        },
        "startTime": "2022-11-17T14:30:59.2453442Z",
        "status": "Failed"
      }
    ]
    ```

=== "Text"

    ``` text
    name                               startTime                     status
    ---------------------------------  ----------------------------  ---------
    08585329112602833828909892130CU17  2022-11-17T14:33:45.2763872Z  Running
    08585329113604313818196774173CU24  2022-11-17T14:32:05.132289Z   Cancelled
    08585329114263413369372385226CU17  2022-11-17T14:30:59.2453442Z  Failed
    ```

=== "CSV"

    ``` text
    name,startTime,status
    08585329112602833828909892130CU17,2022-11-17T14:33:45.2763872Z,Running
    08585329113604313818196774173CU24,2022-11-17T14:32:05.132289Z,Cancelled
    08585329114263413369372385226CU17,2022-11-17T14:30:59.2453442Z,Failed
    ```
