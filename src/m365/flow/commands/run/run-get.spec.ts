import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import command from './run-get.js';

describe(commands.RUN_GET, () => {
  const flowName = '396d5ec9-ae2d-4a84-967d-cd7f56cd8f30';
  const environmentName = 'Default-48595cc3-adce-4267-8e99-0c838923dbb9';
  const runName = '08586653536760200319026785874CU62';

  //#region Mocked Responses flow run get
  const flowResponse = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } } };
  const flowResponseNoEndTime = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } } };
  const triggerInformationResponse = { 'headers': { 'Pragma': 'no-cache', 'Transfer-Encoding': 'chunked', 'Retry-After': '3600', 'Vary': 'Accept-Encoding', 'Cache-Control': 'no-store, no-cache', 'Location': 'https://flow-apim-europe-002-westeurope-01.azure-apim.net/apim/office365/shared-office365-13750b71-ac69-403f-ad09-23e056037806/v4/Mail/OnFlaggedEmail?folderPath=Inbox&importance=Any&fetchOnlyWithAttachment=false&includeAttachments=false&LastPollInformation=eyJMYXN0UmVjZWl2ZWRNYWlsVGltZSI6IjIwMjItMTEtMTVUMTA6MTM6NTUuMDYyNjQ0MyswMDowMCIsIkxhc3RDcmVhdGVkTWFpbFRpbWUiOiIyMDIzLTAzLTA0VDA5OjA1OjIxKzAwOjAwIiwiTGFzdE1lc3NhZ2VJZCI6IkFBTWtBRGd6TjJRMU5UaGlMVEkwTmpZdE5HSXhZUzA1TURkakxUZzFPV1F4Tnpnd1pHTTJaZ0JHQUFBQUFBQzZqUWZVemFjVFNJSHFNdzJ5YWNuVUJ3QmlPQzh4dlltZFQ2RzJFX2hMTUs1a0FBQUFBQUVNQUFCaU9DOHh2WW1kVDZHMkVfaExNSzVrQUFMVXF5ODFBQUE9IiwiTGFzdEludGVybmV0TWVzc2FnZUlkIjoiPERCN1BSMDNNQjUwMTg4Nzk5MTQzMjRGQzY1Njk1ODA5RkUxQUQ5QERCN1BSMDNNQjUwMTguZXVycHJkMDMucHJvZC5vdXRsb29rLmNvbT4ifQ%3d%3d', 'Set-Cookie': '', 'x-ms-request-id': 'ca65b907-e292-41ff-b76e-f3fe60e87c8c;2192534a-f34f-4920-88c9-72b1b1bcf4b4', 'Strict-Transport-Security': 'max-age=31536000; includeSubDomains', 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'DENY', 'Timing-Allow-Origin': '*', 'x-ms-apihub-cached-response': 'true', 'x-ms-apihub-obo': 'false', 'Date': 'Sat, 04 Mar 2023 09:05:21 GMT', 'Content-Type': 'application/json', 'Expires': '-1', 'Content-Length': '1831' }, 'body': { 'from': 'john@contoso.com', 'toRecipients': 'doe@contoso.com', 'subject': 'This is a test mail', 'body': '<html><head>\r\n<meta http-equiv=\'Content- Type\' content=\'text / html; charset=utf - 8\'></head><body><p>This is dummy content</p></body></html>', 'importance': 'normal', 'bodyPreview': 'This is dummy content', 'hasAttachments': false, 'id': 'AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E_hLMK5kAALUqy81AAA=', 'internetMessageId': '<DB7PR03MB5018879914324FC65695809FE1AD9@DB7PR03MB5018.eurprd03.prod.outlook.com>', 'conversationId': 'AAQkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgAQAMqP9zsK8a1CnIYEgHclLTk=', 'receivedDateTime': '2023-03-01T15:06:57+00:00', 'isRead': false, 'attachments': [], 'isHtml': true } };

  const flowResponseFormatted = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual' };
  const flowResponseFormattedNoEndTime = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '', status: 'Running', triggerName: 'manual' };
  const flowResponseFormattedIncludingInformation = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual', triggerInformation: triggerInformationResponse.body };

  const basicActionsInformation = {
    'Compose': {
      'inputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionInputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionInputs%2Fread&sv=1.0&sig=wYD6bMHjdoOawYAMTHqpmyILowLKRCkFEfsNxi1NFzw",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'outputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionOutputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionOutputs%2Fread&sv=1.0&sig=KoI4_RAEVNFUgymktOFxvfb5mCoIKR2WCYGwexjD7kY",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'startTime': "2023-11-17T21:06:14.466889Z",
      'endTime': "2023-11-17T21:06:14.4676104Z",
      'correlation': {
        'actionTrackingId': "dabed750-0ec4-4c34-98f5-cc3695e0811e",
        'clientTrackingId': "08585013686846794051719443325CU251",
        'clientKeywords': [
          "resubmitFlow"
        ]
      },
      'status': "Succeeded",
      'code': "OK"
    },
    'Get_items': {
      'inputsLink': {
        'uri': "https://prod-01.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Get_items/contents/ActionInputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FGet_items%2Fcontents%2FActionInputs%2Fread&sv=1.0&",
        'contentVersion': "test==",
        'contentSize': 356,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'outputsLink': {
        'uri': "https://prod-01.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Get_items/contents/ActionOutputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FGet_items%2Fcontents%2FActionOutputs%2Fread&sv=1.0&",
        'contentVersion': "test==",
        'contentSize': 125480,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'startTime': "2023-11-17T21:06:14.4688325Z",
      'endTime': "2023-11-17T21:06:14.7534248Z",
      'correlation': {
        'actionTrackingId': "af8e70eb-274e-425f-83de-651211e7f6b8",
        'clientTrackingId': "08585013686846794051719443325CU251",
        'clientKeywords': [
          "resubmitFlow"
        ]
      },
      'status': "Succeeded",
      'code': "OK"
    }
  };

  const actionInputStringExample = 'Test';
  const actionInputComplexExample = {
    'host': { 'apiId': "subscriptions/12345678901-1234-1234-1234-12345678901/providers/Microsoft.Web/locations/westeurope/runtimes/europe-001/apis/sharepointonline", 'connectionReferenceName': "shared_sharepointonline", 'operationId': "GetItems" }, 'parameters': { 'dataset': "https://contoso.sharepoint.com/sites/Example", 'table': "a38bbda1-15c8-4479-b9af-4f6e8fe8d26c" }
  };

  const actionOutputStringExample = 'Test';
  const actionOutputComplexExample = { 'value': [{ '@odata.etag': '"1"', '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference', 'Id': 1, 'Title': 'Test', 'Value': 'Test' }] };

  const actionsWithInputOutputInformation = {
    'Compose': {
      'inputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionInputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionInputs%2Fread&sv=1.0&sig=wYD6bMHjdoOawYAMTHqpmyILowLKRCkFEfsNxi1NFzw",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'outputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionOutputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionOutputs%2Fread&sv=1.0&sig=KoI4_RAEVNFUgymktOFxvfb5mCoIKR2WCYGwexjD7kY",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'startTime': "2023-11-17T21:06:14.466889Z",
      'endTime': "2023-11-17T21:06:14.4676104Z",
      'correlation': {
        'actionTrackingId': "dabed750-0ec4-4c34-98f5-cc3695e0811e",
        'clientTrackingId': "08585013686846794051719443325CU251",
        'clientKeywords': [
          "resubmitFlow"
        ]
      },
      'status': "Succeeded",
      'code': "OK",
      'input': actionInputStringExample,
      'output': actionOutputStringExample
    },
    'Get_items': {
      'inputsLink': {
        'uri': "https://prod-01.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Get_items/contents/ActionInputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FGet_items%2Fcontents%2FActionInputs%2Fread&sv=1.0&",
        'contentVersion': "test==",
        'contentSize': 356,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'outputsLink': {
        'uri': "https://prod-01.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Get_items/contents/ActionOutputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FGet_items%2Fcontents%2FActionOutputs%2Fread&sv=1.0&",
        'contentVersion': "test==",
        'contentSize': 125480,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'startTime': "2023-11-17T21:06:14.4688325Z",
      'endTime': "2023-11-17T21:06:14.7534248Z",
      'correlation': {
        'actionTrackingId': "af8e70eb-274e-425f-83de-651211e7f6b8",
        'clientTrackingId': "08585013686846794051719443325CU251",
        'clientKeywords': [
          "resubmitFlow"
        ]
      },
      'status': "Succeeded",
      'code': "OK",
      'input': actionInputComplexExample,
      'output': actionOutputComplexExample
    }
  };

  const actionsWithInputOutputInformationOfSelectedAction = {
    'Compose': {
      'inputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionInputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionInputs%2Fread&sv=1.0&sig=wYD6bMHjdoOawYAMTHqpmyILowLKRCkFEfsNxi1NFzw",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'outputsLink': {
        'uri': "https://prod-255.westeurope.logic.azure.com:443/workflows/88b5146587d144ea85efb15683c4e58f/runs/08586652586741142222645090602CU35/actions/Compose/contents/ActionOutputs?api-version=2016-06-01&se=2023-12-14T19%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Factions%2FCompose%2Fcontents%2FActionOutputs%2Fread&sv=1.0&sig=KoI4_RAEVNFUgymktOFxvfb5mCoIKR2WCYGwexjD7kY",
        'contentVersion': "test==",
        'contentSize': 6,
        'contentHash': {
          'algorithm': "md5",
          'value': "test=="
        }
      },
      'startTime': "2023-11-17T21:06:14.466889Z",
      'endTime': "2023-11-17T21:06:14.4676104Z",
      'correlation': {
        'actionTrackingId': "dabed750-0ec4-4c34-98f5-cc3695e0811e",
        'clientTrackingId': "08585013686846794051719443325CU251",
        'clientKeywords': [
          "resubmitFlow"
        ]
      },
      'status': "Succeeded",
      'code': "OK",
      'input': actionInputStringExample,
      'output': actionOutputStringExample
    }
  };

  const flowResponseFormattedIncludingActionsBasicInformation = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { actions: basicActionsInformation, 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual' };

  const commandResultIncluddingAllActions = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { actions: basicActionsInformation, 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual', actions: actionsWithInputOutputInformation };

  const commandResultIncludingSelectedAction = { 'name': runName, 'id': '/providers/Microsoft.ProcessSimple/environments/Default-48595cc3-adce-4267-8e99-0c838923dbb9/flows/edf73e7e-9928-4cb9-8eb2-fc263f375ada/runs/08586652586741142222645090602CU35', 'type': 'Microsoft.ProcessSimple/environments/flows/runs', 'properties': { actions: basicActionsInformation, 'startTime': '2018-09-07T19:23:31.3640166Z', 'endTime': '2018-09-07T19:33:31.3640166Z', 'status': 'Running', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'trigger': { 'name': 'manual', 'inputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerInputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerInputs%2Fread&sv=1.0&sig=yjCLU5P9pqCoDPHmRFq-oLKGhcFNnGUXH6ojpQz9z6Q', 'contentVersion': '1UI/8pYQdWDVSsijF+0l2Q==', 'contentSize': 58, 'contentHash': { 'algorithm': 'md5', 'value': '1UI/8pYQdWDVSsijF+0l2Q==' } }, 'outputsLink': { 'uri': 'https://prod-09.westeurope.logic.azure.com:443/workflows/8c76aebc46484a29889c426f55a52f55/runs/08586652586741142222645090602CU35/contents/TriggerOutputs?api-version=2016-06-01&se=2018-09-07T23%3A00%3A00.0000000Z&sp=%2Fruns%2F08586652586741142222645090602CU35%2Fcontents%2FTriggerOutputs%2Fread&sv=1.0&sig=SLd0zBScyF6F6eMTBIROalK32e2t0od2SDETe9X9SMM', 'contentVersion': 'YgV2ecynizFKxzT8yiNtpA==', 'contentSize': 4244, 'contentHash': { 'algorithm': 'md5', 'value': 'YgV2ecynizFKxzT8yiNtpA==' } }, 'startTime': '2018-09-07T19:23:31.3482269Z', 'endTime': '2018-09-07T19:23:31.3482269Z', 'correlation': { 'clientTrackingId': '08586652586741142222645090602CU35', 'clientKeywords': ['testFlow'] }, 'status': 'Succeeded' } }, startTime: '2018-09-07T19:23:31.3640166Z', endTime: '2018-09-07T19:33:31.3640166Z', status: 'Running', triggerName: 'manual', actions: actionsWithInputOutputInformationOfSelectedAction };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RUN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified run', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormatted));
  });

  it('renders empty string for endTime, if the run specified is still running', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponseNoEndTime;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedNoEndTime));
  });

  it('retrieves information about the specified run including trigger information using includeTriggerInformation parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponse;
      }

      if (opts.url === flowResponse.properties.trigger.outputsLink.uri) {
        return triggerInformationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, includeTriggerInformation: true, verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedIncludingInformation));
  });

  it('retrieves information about the specified run including trigger information using withTrigger parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?api-version=2016-11-01`) {
        return flowResponse;
      }

      if (opts.url === flowResponse.properties.trigger.outputsLink.uri) {
        return triggerInformationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, withTrigger: true, verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedIncludingInformation));
  });

  it('retrieves information about the specified run including all actions details information using withActions parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?$expand=properties%2Factions&api-version=2016-11-01`) {
        return flowResponseFormattedIncludingActionsBasicInformation;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.inputsLink.uri) {
        return actionInputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.outputsLink.uri) {
        return actionOutputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.inputsLink.uri) {
        return actionInputComplexExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.outputsLink.uri) {
        return actionOutputComplexExample;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, withActions: true, verbose: true } });
    assert(loggerLogSpy.calledWith(commandResultIncluddingAllActions));
  });

  it('retrieves information about the specified run including details information of selected actions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?$expand=properties%2Factions&api-version=2016-11-01`) {
        return flowResponseFormattedIncludingActionsBasicInformation;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.inputsLink.uri) {
        return actionInputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.outputsLink.uri) {
        return actionOutputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.inputsLink.uri) {
        return actionInputComplexExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.outputsLink.uri) {
        return actionOutputComplexExample;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, withActions: "Compose", verbose: true } });
    assert(loggerLogSpy.calledWith(commandResultIncludingSelectedAction));
  });

  it('retrieves information of actions without details even when name of the selected action is incorrect', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentName}/flows/${flowName}/runs/${runName}?$expand=properties%2Factions&api-version=2016-11-01`) {
        return flowResponseFormattedIncludingActionsBasicInformation;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.inputsLink.uri) {
        return actionInputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Compose.outputsLink.uri) {
        return actionOutputStringExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.inputsLink.uri) {
        return actionInputComplexExample;
      }

      if (opts.url === flowResponseFormattedIncludingActionsBasicInformation.properties.actions.Get_items.outputsLink.uri) {
        return actionOutputComplexExample;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName, withActions: "Wrong,Wrong2", verbose: true } });
    assert(loggerLogSpy.calledWith(flowResponseFormattedIncludingActionsBasicInformation));
  });

  it('correctly handles Flow not found', async () => {
    sinon.stub(request, 'get').rejects({
      'error': {
        'code': 'FlowNotFound',
        'message': `Could not find flow '${flowName}'.`
      }
    });

    await assert.rejects(command.action(logger, { options: { flowName: flowName, environmentName: environmentName, name: runName } } as any),
      new CommandError(`Could not find flow '${flowName}'.`));
  });

  it('fails validation if the flowName is not valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: 'invalid', name: runName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the withActions parameter is not valid boolean or string', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, name: runName, withActions: -1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the flowName is not valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, name: runName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});