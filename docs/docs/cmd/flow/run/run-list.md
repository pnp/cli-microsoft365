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

`--status [status]` 
: Filter the results to only flow runs with a given status: `Succeeded`, `Running`, `Failed` or `Cancelled`. By default all flow runs are listed.

`--triggerStartTime [triggerStartTime]`
: Time indicating the inclusive start of a time range of flow runs to return. This should be defined as a valid ISO 8601 string (2021-12-16T18:28:48.6964197Z).

`--triggerEndTime [triggerEndTime]`
: Time indicating the exclusive end of a time range of flow runs to return. This should be defined as a valid ISO 8601 string (2021-12-16T18:28:48.6964197Z).

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

List runs of the specified Microsoft Flow with a specific status

```sh
m365 flow run list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --status Running
```

List runs of the specified Microsoft Flow between a specific time range

```sh
m365 flow run list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flowName 5923cb07-ce1a-4a5c-ab81-257ce820109a --triggerStartTime 2023-01-21T18:19:00Z --triggerEndTime 2023-01-22T00:00:00Z
```

## Response

### Standard response

=== "JSON"

    ```json
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
      }
    ]
    ```

=== "Text"

    ```text
    name                               startTime                     status
    ---------------------------------  ----------------------------  ---------
    08585329112602833828909892130CU17  2022-11-17T14:33:45.2763872Z  Running
    ```

=== "CSV"

    ```csv
    name,startTime,status
    08585329112602833828909892130CU17,2022-11-17T14:33:45.2763872Z,Running
    ```
