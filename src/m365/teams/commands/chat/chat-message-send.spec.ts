import * as assert from 'assert';
import * as os from 'os';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./chat-message-send');

describe(commands.CHAT_MESSAGE_SEND, () => {  
  //#region Mocked Responses  
  const findGroupChatsByMembersResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats(id,topic,createdDateTime,members,members())", "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'group'&$expand=members&$select=id,topic,createdDateTime,members&$skiptoken=eyJDb250aW51YXRpb25Ub2tlbiI6Ilczc2ljM1JoY25RaU9pSXlNREl5TFRBeExUSXdWREE1T2pRME9qVXhMakl5Tnlzd01Eb3dNQ0lzSW1WdVpDSTZJakl3TWpJdE1ERXRNakJVTURrNk5EUTZOVEV1TWpJM0t6QXdPakF3SWl3aWMyOXlkRTl5WkdWeUlqb3dmU3g3SW5OMFlYSjBJam9pTVRrM01DMHdNUzB3TVZRd01Eb3dNRG93TUNzd01Eb3dNQ0lzSW1WdVpDSTZJakU1TnpBdE1ERXRNREZVTURBNk1EQTZNREF1TURBeEt6QXdPakF3SWl3aWMyOXlkRTl5WkdWeUlqb3dmVjA9IiwiQ2hhdFR5cGUiOiJjaGF0fG1lZXRpbmd8c2ZiaW50ZXJvcGNoYXR8cGhvbmVjaGF0In0%3d", "value": [{ "id": "19:35bd5bc75e604da8a64e6cba7cfcf175@thread.v2", "topic": "Megan Bowen_Alex Wilber_Sundar Ganesan_ArchivedChat", "createdDateTime": "2021-12-22T13:13:11.023Z", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A35bd5bc75e604da8a64e6cba7cfcf175%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozNWJkNWJjNzVlNjA0ZGE4YTY0ZTZjYmE3Y2ZjZjE3NUB0aHJlYWQudjIjI2ExZDY1Nzk0LWMyYmUtNGEzMy04MTExLWVlY2Y2OGZlOWYzNQ==", "roles": ["Owner"], "displayName": "Alex Wilber", "visibleHistoryStartDateTime": "2021-12-22T13:13:11.023Z", "userId": "a1d65794-c2be-4a33-8111-eecf68fe9f35", "email": "AlexW@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozNWJkNWJjNzVlNjA0ZGE4YTY0ZTZjYmE3Y2ZjZjE3NUB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "2021-12-22T13:13:11.023Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozNWJkNWJjNzVlNjA0ZGE4YTY0ZTZjYmE3Y2ZjZjE3NUB0aHJlYWQudjIjIzcwZTM3MDExLWY2ZjItNDAwMi04MDU2LWQ2MDg1YjQ5N2E2ZA==", "roles": ["Owner"], "displayName": "Nate Grecian", "visibleHistoryStartDateTime": "2021-12-22T13:13:11.023Z", "userId": "70e37011-f6f2-4002-8056-d6085b497a6d", "email": "NateG@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] }, { "id": "19:c03b5a8f9a2e42788561a89d055e6de5@thread.v2", "topic": null, "createdDateTime": "2021-12-09T08:22:07.845Z", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3Ac03b5a8f9a2e42788561a89d055e6de5%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTpjMDNiNWE4ZjlhMmU0Mjc4ODU2MWE4OWQwNTVlNmRlNUB0aHJlYWQudjIjI2FkN2MxYTU5LWM0N2ItNDdmYi1hNDgxLTM1NWI0ZmM5YzEzNA==", "roles": ["Owner"], "displayName": "Andrew Konnelli", "visibleHistoryStartDateTime": "2021-12-09T08:22:07.845Z", "userId": "ad7c1a59-c47b-47fb-a481-355b4fc9c134", "email": "AndrewK@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTpjMDNiNWE4ZjlhMmU0Mjc4ODU2MWE4OWQwNTVlNmRlNUB0aHJlYWQudjIjIzI4YzRlMTdkLWI1NzktNGUxZS1iMDNjLWEyOTZkMTJjYmZiOQ==", "roles": ["Owner"], "displayName": "Dave Kapowski", "visibleHistoryStartDateTime": "2021-12-09T08:22:07.845Z", "userId": "28c4e17d-b579-4e1e-b03c-a296d12cbfb9", "email": "DaveK@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTpjMDNiNWE4ZjlhMmU0Mjc4ODU2MWE4OWQwNTVlNmRlNUB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "2021-12-09T08:22:07.845Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] }, { "id": "19:8a2a5e0f94f74346908e6a7b0023da50@thread.v2", "topic": null, "createdDateTime": "2021-12-09T07:17:25.322Z", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A8a2a5e0f94f74346908e6a7b0023da50%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo4YTJhNWUwZjk0Zjc0MzQ2OTA4ZTZhN2IwMDIzZGE1MEB0aHJlYWQudjIjI2FkN2MxYTU5LWM0N2ItNDdmYi1hNDgxLTM1NWI0ZmM5YzEzNA==", "roles": ["Owner"], "displayName": "Andrew Konnelli", "visibleHistoryStartDateTime": "2021-12-09T07:17:25.322Z", "userId": "ad7c1a59-c47b-47fb-a481-355b4fc9c134", "email": "AndrewK@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo4YTJhNWUwZjk0Zjc0MzQ2OTA4ZTZhN2IwMDIzZGE1MEB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "2021-12-09T07:17:25.322Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] } ] };

  const findGroupChatsByMembersResponseWithNextLink: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats(id,topic,createdDateTime,members,members())", "value": [{ "id": "19:5fb8d18dd38b40a4ae0209888adf5c38@thread.v2", "topic": "CC Call v3", "createdDateTime": "2021-10-18T16:56:30.205Z", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A5fb8d18dd38b40a4ae0209888adf5c38%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjI2M2MWYxYjc5LTA4MmItNDRhYy05YzFiLTZhZTFhZjc5NTMzNg==", "roles": ["Owner"], "displayName": "Alex Wilber", "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "c61f1b79-082b-44ac-9c1b-6ae1af795336", "email": "AlexW@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjI2Y2ZjVlNzA2LWY5NjYtNDlkYi1iNWEwLTEyMGFiZGE3MTc2Nw==", "roles": ["Owner"], "displayName": "Nate Grecian", "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "f6f5e706-f966-49db-b5a0-120abda71767", "email": "NateG@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] }, { "id": "19:d08368bb4b5c4a70a4606b8914108327@thread.v2", "topic": null, "createdDateTime": "2020-06-26T14:20:19.997Z", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3Ad08368bb4b5c4a70a4606b8914108327%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjI2M2MWYxYjc5LTA4MmItNDRhYy05YzFiLTZhZTFhZjc5NTMzNg==", "roles": ["Owner"], "displayName": null, "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "d3d5g634-082b-44ac-9c1b-6ae1af795336", "email": null, "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjI2Y2ZjVlNzA2LWY5NjYtNDlkYi1iNWEwLTEyMGFiZGE3MTc2Nw==", "roles": ["Owner"], "displayName": "Nate Grecian", "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "f6f5e706-f966-49db-b5a0-120abda71767", "email": "NateG@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo1ZmI4ZDE4ZGQzOGI0MGE0YWUwMjA5ODg4YWRmNWMzOEB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "2021-10-18T16:56:30.205Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] } ] };
  
  const singleChatByNameResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats(id,topic,createdDateTime,chatType,members())", "value": [{ "id": "19:98a7bf5fe7884694b8078541c5eb6e56@thread.v2", "topic": "Just a conversation", "createdDateTime": "2021-04-30T07:36:14.152Z", "chatType": "group", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A98a7bf5fe7884694b8078541c5eb6e56%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo5OGE3YmY1ZmU3ODg0Njk0YjgwNzg1NDFjNWViNmU1NkB0aHJlYWQudjIjIzY4Nzg2ZjQ5LTMzN2ItNGFmNy05Y2E5LTQ1Y2NmNDM2MTE5Yw==", "roles": ["Owner"], "displayName": "Dave Kapowski", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "68786f49-337b-4af7-9ca9-45ccf436119c", "email": "DaveK@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo5OGE3YmY1ZmU3ODg0Njk0YjgwNzg1NDFjNWViNmU1NkB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo5OGE3YmY1ZmU3ODg0Njk0YjgwNzg1NDFjNWViNmU1NkB0aHJlYWQudjIjI2EwNjA0NjIzLTdjMjgtNDk2Zi05ZDNjLWU4N2RmNjE4YjMxMA==", "roles": ["Owner"], "displayName": "Andrew Konelli", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "a0604623-7c28-496f-9d3c-e87df618b310", "email": "AndrewK@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] } ] };

  const multipleChatByNameResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats(id,topic,createdDateTime,chatType,members())", "value": [{ "id": "19:309128478c1743b19bebd08efc390efb@thread.v2", "topic": "Just a conversation with same name", "createdDateTime": "2021-09-14T07:44:11.5Z", "chatType": "group", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A309128478c1743b19bebd08efc390efb%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozMDkxMjg0NzhjMTc0M2IxOWJlYmQwOGVmYzM5MGVmYkB0aHJlYWQudjIjI2ExZDY1Nzk0LWMyYmUtNGEzMy04MTExLWVlY2Y2OGZlOWYzNQ==", "roles": ["Owner"], "displayName": "Alex Wilber", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "a1d65794-c2be-4a33-8111-eecf68fe9f35", "email": "AlexW@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozMDkxMjg0NzhjMTc0M2IxOWJlYmQwOGVmYzM5MGVmYkB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTozMDkxMjg0NzhjMTc0M2IxOWJlYmQwOGVmYzM5MGVmYkB0aHJlYWQudjIjIzg4MGRlNWIyLTk5MjQtNGViMS1iZjdhLWVlZDhkNzFiNjExNg==", "roles": ["Guest"], "displayName": "Nate Grecian", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "880de5b2-9924-4eb1-bf7a-eed8d71b6116", "email": "NateG@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] }, { "id": "19:650081f4700a4414ac15cd7993129f80@thread.v2", "topic": "Just a conversation with same name", "createdDateTime": "2020-06-26T08:27:55.154Z", "chatType": "group", "members@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A650081f4700a4414ac15cd7993129f80%40thread.v2')/members", "members": [{ "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo2NTAwODFmNDcwMGE0NDE0YWMxNWNkNzk5MzEyOWY4MEB0aHJlYWQudjIjIzQ4ZDMxODg3LTVmYWQtNGQ3My1hOWY1LTNjMzU2ZTY4YTAzOA==", "roles": ["Owner"], "displayName": "Megan Bowen", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "48d31887-5fad-4d73-a9f5-3c356e68a038", "email": "MeganB@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo2NTAwODFmNDcwMGE0NDE0YWMxNWNkNzk5MzEyOWY4MEB0aHJlYWQudjIjIzU4YmE0NzI3LTVjOWYtNDllOS04NDM4LTFiMzI4NTM1Mzc5MQ==", "roles": ["Owner"], "displayName": "Alex Wilber", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "58ba4727-5c9f-49e9-8438-1b3285353791", "email": "AlexW@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" }, { "@odata.type": "#microsoft.graph.aadUserConversationMember", "id": "MCMjMCMjZGNkMjE5ZGQtYmM2OC00YjliLWJmMGItNGEzM2E3OTZiZTM1IyMxOTo2NTAwODFmNDcwMGE0NDE0YWMxNWNkNzk5MzEyOWY4MEB0aHJlYWQudjIjIzE5ODA2ZjhjLTAxNTMtNDQxNC05MzA1LWRjNDgwMTJiNDc4OQ==", "roles": ["Owner"], "displayName": "Nate Grecian", "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z", "userId": "19806f8c-0153-4414-9305-dc48012b4789", "email": "NateG@M365x214355.onmicrosoft.com", "tenantId": "dcd219dd-bc68-4b9b-bf0b-4a33a796be35" } ] } ] };

  const noChatByNameResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats(id,topic,createdDateTime,chatType,members())", "value": [] };
  
  const chatCreatedResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats/$entity", "id": "19:82fe7758-5bb3-4f0d-a43f-e555fd399c6f_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:28.51Z", "chatType": "oneOnOne" };

  const groupChatCreatedResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats/$entity", "id": "19:650081f4700a4414ac15cd7993129f80@thread.v2", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:28.51Z", "chatType": "group" };

  const messageSentResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#chats('19%3A2da4c29f6d7041eca70b638b43d45437%40thread.v2')/messages/$entity", "id": "1616991463150", "replyToId": null, "etag": "1616991463150", "messageType": "message", "createdDateTime": "2021-03-29T04:17:43.15Z", "lastModifiedDateTime": "2021-03-29T04:17:43.15Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "conversation": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello World" }, "attachments": [], "mentions": [], "reactions": [] };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
 
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
  });

  beforeEach(() => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => {
      return 'MeganB@M365x214355.onmicrosoft.com';
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:82fe7758-5bb3-4f0d-a43f-e555fd399c6f_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces/messages`
        || opts.url === `https://graph.microsoft.com/v1.0/chats/19:98a7bf5fe7884694b8078541c5eb6e56@thread.v2/messages`
        || opts.url === `https://graph.microsoft.com/v1.0/chats/19:c03b5a8f9a2e42788561a89d055e6de5@thread.v2/messages`) {
        return Promise.resolve(messageSentResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'group'&$expand=members&$select=id,topic,createdDateTime,members`) {
        return Promise.resolve(findGroupChatsByMembersResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'group'&$expand=members&$select=id,topic,createdDateTime,members&$skiptoken=eyJDb250aW51YXRpb25Ub2tlbiI6Ilczc2ljM1JoY25RaU9pSXlNREl5TFRBeExUSXdWREE1T2pRME9qVXhMakl5Tnlzd01Eb3dNQ0lzSW1WdVpDSTZJakl3TWpJdE1ERXRNakJVTURrNk5EUTZOVEV1TWpJM0t6QXdPakF3SWl3aWMyOXlkRTl5WkdWeUlqb3dmU3g3SW5OMFlYSjBJam9pTVRrM01DMHdNUzB3TVZRd01Eb3dNRG93TUNzd01Eb3dNQ0lzSW1WdVpDSTZJakU1TnpBdE1ERXRNREZVTURBNk1EQTZNREF1TURBeEt6QXdPakF3SWl3aWMyOXlkRTl5WkdWeUlqb3dmVjA9IiwiQ2hhdFR5cGUiOiJjaGF0fG1lZXRpbmd8c2ZiaW50ZXJvcGNoYXR8cGhvbmVjaGF0In0%3d`) {
        return Promise.resolve(findGroupChatsByMembersResponseWithNextLink);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=topic eq '${encodeURIComponent('Just a conversation')}'&$expand=members&$select=id,topic,createdDateTime,chatType`) {
        return Promise.resolve(singleChatByNameResponse);
      }      
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=topic eq '${encodeURIComponent('Just a conversation with same name')}'&$expand=members&$select=id,topic,createdDateTime,chatType`) {
        return Promise.resolve(multipleChatByNameResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=topic eq '${encodeURIComponent('Nonexistent conversation name')}'&$expand=members&$select=id,topic,createdDateTime,chatType`) {
        return Promise.resolve(noChatByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      accessToken.getUserNameFromAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      accessToken.getUserNameFromAccessToken
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHAT_MESSAGE_SEND), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });  

  it('fails validation if chatId and chatName and userEmails are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false,
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the chatId is not valid', () => {
    const actual = command.validate({
      options: {
        chatId: "8b081ef6",
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for an incorrect chatId missing leading 19:.', (done) => {
    const actual = command.validate({
      options: {
        chatId: '8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d@unq.gbl.spaces',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for an incorrect chatId missing trailing @thread.v2 or @unq.gbl.spaces', (done) => {
    const actual = command.validate({
      options: {
        chatId: '19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });
  
  it('fails validation for an invalid email address (single)', (done) => {
    const actual = command.validate({
      options: {
        userEmails: 'alexwcontoso.com',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for invalid email addresses (multiple)', (done) => {
    const actual = command.validate({
      options: {
        userEmails: 'alexw@contoso.com,natecontoso.com',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if chatId and chatName properties are both defined', (done) => {
    const actual = command.validate({
      options: {
        chatId: '19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d',
        chatName: 'test',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if chatId and userEmails properties are both defined', (done) => {
    const actual = command.validate({
      options: {
        chatId: '19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d',
        userEmails: 'alexw@contoso.com',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });
  
  it('fails validation if chatName and userEmails properties are both defined', (done) => {
    const actual = command.validate({
      options: {
        chatName: 'test',
        userEmails: 'alexw@contoso.com',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if all three mutually exclusive properties are defined', (done) => {
    const actual = command.validate({
      options: {
        chatId: '19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d',
        chatName: 'test',
        userEmails: 'alexw@contoso.com',
        message: "Hello World"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });
  
  it('fails validation if message is not specified', () => {
    const actual = command.validate({
      options: {
        debug: false,
        chatId: "19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d@unq.gbl.spaces"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct chatId input', () => {
    const actual = command.validate({
      options: {
        chatId: "19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d@unq.gbl.spaces",
        message: "Hello World"
      }
    });
    assert.strictEqual(actual, true);
  });
  
  it('validates for a correct chatName input', () => {
    const actual = command.validate({
      options: {
        chatName: 'test',
        message: "Hello World"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct userEmails input', () => {
    const actual = command.validate({
      options: {
        userEmails: 'alexw@contoso.com',
        message: "Hello World"
      }
    });
    assert.strictEqual(actual, true);
  });
  
  it('validates for a correct userEmails (array) input', () => {
    const actual = command.validate({
      options: {
        userEmails: 'alexw@contoso.com,nateg@contoso.com',
        message: "Hello World"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
  
  it('sends chat message using chatId', (done) => {
    command.action(logger, {
      options: {
        chatId: "19:82fe7758-5bb3-4f0d-a43f-e555fd399c6f_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('sends chat message using chatName', (done) => {    
    command.action(logger, {
      options: {
        chatName: "Just a conversation",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('sends chat message using userEmails (single)', (done) => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.resolve(chatCreatedResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:82fe7758-5bb3-4f0d-a43f-e555fd399c6f_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces/messages`) {
        return Promise.resolve(messageSentResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        userEmails: "AlexW@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('sends chat message to existing conversation using userEmails (multiple)', (done) => {
    command.action(logger, {
      options: {
        userEmails: "AndrewK@M365x214355.onmicrosoft.com,DaveK@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('sends chat message to new conversation using userEmails (multiple)', (done) => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.resolve(groupChatCreatedResponse);
      }      
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:650081f4700a4414ac15cd7993129f80@thread.v2/messages`) {
        return Promise.resolve(messageSentResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        userEmails: "AlexW@M365x214355.onmicrosoft.com,DaveK@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('fails sending message with nonexistent chatName', (done) => {    
    command.action(logger, {
      options: {
        chatName: "Nonexistent conversation name",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`No chat conversation was found with this name.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
    
  it('fails sending message with multiple found chat conversations by chatName', (done) => {    
    command.action(logger, {
      options: {
        chatName: "Just a conversation with same name",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`Multiple chat conversations with this topic found. Please disambiguate:${os.EOL}${[
            `- 19:309128478c1743b19bebd08efc390efb@thread.v2 - ${new Date("2021-09-14T07:44:11.5Z").toLocaleString()} - AlexW@M365x214355.onmicrosoft.com, MeganB@M365x214355.onmicrosoft.com, NateG@M365x214355.onmicrosoft.com`,
            `- 19:650081f4700a4414ac15cd7993129f80@thread.v2 - ${new Date("2020-06-26T08:27:55.154Z").toLocaleString()} - MeganB@M365x214355.onmicrosoft.com, AlexW@M365x214355.onmicrosoft.com, NateG@M365x214355.onmicrosoft.com`
          ].join(os.EOL)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails sending message with multiple found chat conversations by userEmails', (done) => {    
    command.action(logger, {
      options: {
        userEmails: "AlexW@M365x214355.onmicrosoft.com,NateG@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`Multiple chat conversations with this topic found. Please disambiguate:${os.EOL}${[
            `- 19:35bd5bc75e604da8a64e6cba7cfcf175@thread.v2 - Megan Bowen_Alex Wilber_Sundar Ganesan_ArchivedChat - ${new Date("2021-12-22T13:13:11.023Z").toLocaleString()}`,
            `- 19:5fb8d18dd38b40a4ae0209888adf5c38@thread.v2 - CC Call v3 - ${new Date("2021-10-18T16:56:30.205Z").toLocaleString()}`
          ].join(os.EOL)}`)));                              
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  // The following test is used to test the retry mechanism in use because of an intermittent Graph issue.
  it('sends chat message using userEmails with single retry because of 404 intermittent failure', (done) => {
    sinonUtil.restore(request.post);
    let retries: number = 0;
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        if (retries === 0) {
          retries++;
          return Promise.reject({ message: "Request failed with status code 404" });
        } 
        else {
          return Promise.resolve(chatCreatedResponse);
        }
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:82fe7758-5bb3-4f0d-a43f-e555fd399c6f_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces/messages`) {
        return Promise.resolve(messageSentResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        userEmails: "AlexW@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  // The following test is used to test the retry mechanism in use because of an intermittent Graph issue.
  it('fails sending chat message when maximum of 3 retries with 404 intermittent failure have occurred', (done) => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.reject("Request failed with status code 404");
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        userEmails: "AlexW@M365x214355.onmicrosoft.com",
        message: "Hello World"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`Request failed with status code 404`)));                              
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });  
});