import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./term-list');

describe(commands.TERM_LIST, () => {
  const csomDefaultResponseJson = JSON.stringify([
    {
      "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
    }, 70, {
      "IsNull": false
    }, 71, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
    }, 73, {
      "IsNull": false
    }, 74, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
    }, 76, {
      "IsNull": false
    }, 78, {
      "IsNull": false
    }, 79, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
    }, 81, {
      "IsNull": false
    }, 83, {
      "IsNull": false
    }, 84, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
    }, 86, {
      "IsNull": false
    }, 87, {
      "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
        {
          "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

          }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

          }, "MergedTermIds": [

          ], "PathOfTerm": "HR", "TermsCount": 0
        }, {
          "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

          }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

          }, "MergedTermIds": [

          ], "PathOfTerm": "IT", "TermsCount": 0
        }, {
          "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

          }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

          }, "MergedTermIds": [

          ], "PathOfTerm": "Leadership", "TermsCount": 1
        }
      ]
    }
  ]);

  const csomFirstChildTermResponseJson = JSON.stringify([
    {
      "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
    }, 70, {
      "IsNull": false
    }, 71, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
    }, 73, {
      "IsNull": false
    }, 74, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
    }, 76, {
      "IsNull": false
    }, 78, {
      "IsNull": false
    }, 79, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
    }, 81, {
      "IsNull": false
    }, 83, {
      "IsNull": false
    }, 84, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
    }, 86, {
      "IsNull": false
    }, 87, {
      "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
        {
          "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "6ede85a0-70c8-6000-02cc-23cfae5cac8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNHOmHw1O1kkuIa5r3F81zsA==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(c387e91c-b553-4b92-886b-9af717cd73b0)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "Child 1", "CustomProperties": {

          }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

          }, "MergedTermIds": [

          ], "PathOfTerm": "Leadership;Child 1", "TermsCount": 1
        }
      ]
    }
  ]);

  const csomSecondChildTermResponseJson = JSON.stringify([
    {
      "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
    }, 70, {
      "IsNull": false
    }, 71, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
    }, 73, {
      "IsNull": false
    }, 74, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
    }, 76, {
      "IsNull": false
    }, 78, {
      "IsNull": false
    }, 79, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
    }, 81, {
      "IsNull": false
    }, 83, {
      "IsNull": false
    }, 84, {
      "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
    }, 86, {
      "IsNull": false
    }, 87, {
      "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
        {
          "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "94ff85a0-d037-6000-02cc-2f3cb55cae2a|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSN+zepP1QYfkGbYiwj/kPrgw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(3fa937fb-1854-417e-9b62-2c23fe43eb83)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "Child 2", "CustomProperties": {

          }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

          }, "MergedTermIds": [

          ], "PathOfTerm": "Leadership;Child 1;Child 2", "TermsCount": 0
        }
      ]
    }
  ]);

  const csomDefaultResponseFormatted = [{
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==",
    "CreatedDate": "2018-09-13T11:52:55.320Z",
    "Id": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
    "LastModifiedDate": "2018-09-13T11:52:55.337Z",
    "Name": "HR",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "HR",
    "TermsCount": 0
  },
  {
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==",
    "CreatedDate": "2018-09-13T11:52:55.477Z",
    "Id": "247543b6-45f2-4232-b9e8-66c5bf53c31e",
    "LastModifiedDate": "2018-09-13T11:52:55.490Z",
    "Name": "IT",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "IT",
    "TermsCount": 0
  },
  {
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD/1ASKE2ziDgfrY1GAg==",
    "CreatedDate": "2018-09-13T11:52:55.600Z",
    "Id": "ffc3608f-1250-4d28-b388-381fad8d4602",
    "LastModifiedDate": "2018-09-13T11:52:55.617Z",
    "Name": "Leadership",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "Leadership",
    "TermsCount": 1
  }];

  const csomChildResponseFormatted = [{
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==",
    "CreatedDate": "2018-09-13T11:52:55.320Z",
    "Id": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
    "LastModifiedDate": "2018-09-13T11:52:55.337Z",
    "Name": "HR",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "HR",
    "TermsCount": 0
  },
  {
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==",
    "CreatedDate": "2018-09-13T11:52:55.477Z",
    "Id": "247543b6-45f2-4232-b9e8-66c5bf53c31e",
    "LastModifiedDate": "2018-09-13T11:52:55.490Z",
    "Name": "IT",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "IT",
    "TermsCount": 0
  },
  {
    "_ObjectType_": "SP.Taxonomy.Term",
    "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD/1ASKE2ziDgfrY1GAg==",
    "CreatedDate": "2018-09-13T11:52:55.600Z",
    "Id": "ffc3608f-1250-4d28-b388-381fad8d4602",
    "LastModifiedDate": "2018-09-13T11:52:55.617Z",
    "Name": "Leadership",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "Leadership",
    "TermsCount": 1,
    "Children": [{
      "_ObjectType_": "SP.Taxonomy.Term",
      "_ObjectIdentity_": "6ede85a0-70c8-6000-02cc-23cfae5cac8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNHOmHw1O1kkuIa5r3F81zsA==",
      "CreatedDate": "2018-09-13T11:52:55.320Z",
      "Id": "c387e91c-b553-4b92-886b-9af717cd73b0",
      "LastModifiedDate": "2018-09-13T11:52:55.337Z",
      "Name": "Child 1",
      "CustomProperties": {},
      "CustomSortOrder": null,
      "IsAvailableForTagging": true,
      "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
      "Description": "",
      "IsDeprecated": false,
      "IsKeyword": false,
      "IsPinned": false,
      "IsPinnedRoot": false,
      "IsReused": false,
      "IsRoot": true,
      "IsSourceTerm": true,
      "LocalCustomProperties": {},
      "MergedTermIds": [],
      "PathOfTerm": "Leadership;Child 1",
      "TermsCount": 1,
      "Children": [{
        "_ObjectType_": "SP.Taxonomy.Term",
        "_ObjectIdentity_": "94ff85a0-d037-6000-02cc-2f3cb55cae2a|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSN+zepP1QYfkGbYiwj/kPrgw==",
        "CreatedDate": "2018-09-13T11:52:55.320Z",
        "Id": "3fa937fb-1854-417e-9b62-2c23fe43eb83",
        "LastModifiedDate": "2018-09-13T11:52:55.337Z",
        "Name": "Child 2",
        "CustomProperties": {},
        "CustomSortOrder": null,
        "IsAvailableForTagging": true,
        "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
        "Description": "",
        "IsDeprecated": false,
        "IsKeyword": false,
        "IsPinned": false,
        "IsPinnedRoot": false,
        "IsReused": false,
        "IsRoot": true,
        "IsSourceTerm": true,
        "LocalCustomProperties": {},
        "MergedTermIds": [],
        "PathOfTerm": "Leadership;Child 1;Child 2",
        "TermsCount": 0
      }]
    }]
  }];

  const csomDefaultResponseFormattedText = [
    {
      "Id": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
      "Name": "HR"
    },
    {
      "Id": "247543b6-45f2-4232-b9e8-66c5bf53c31e",
      "Name": "IT"
    },
    {
      "Id": "ffc3608f-1250-4d28-b388-381fad8d4602",
      "Name": "Leadership"
    }
  ];

  const csomChildResponseFormattedText = [
    {
      "Id": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
      "Name": "HR",
      "ParentTermId": ""
    },
    {
      "Id": "247543b6-45f2-4232-b9e8-66c5bf53c31e",
      "Name": "IT",
      "ParentTermId": ""
    },
    {
      "Id": "ffc3608f-1250-4d28-b388-381fad8d4602",
      "Name": "Leadership",
      "ParentTermId": ""
    },
    {
      "Id": "c387e91c-b553-4b92-886b-9af717cd73b0",
      "Name": "Child 1",
      "ParentTermId": "ffc3608f-1250-4d28-b388-381fad8d4602"
    },
    {
      "Id": "3fa937fb-1854-417e-9b62-2c23fe43eb83",
      "Name": "Child 2",
      "ParentTermId": "c387e91c-b553-4b92-886b-9af717cd73b0"
    }
  ];

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => {
      if (settingName === "prompt") { return false; }
      else {
        return defaultValue;
      }
    }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TERM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when webUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc', termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/project-x', termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Name', 'ParentTermId']);
  });

  it('gets taxonomy terms including child terms from term set by id, term group by id', async () => {
    const termGroupId = '0e8f395e-ff58-4d45-9ff7-e331ab728beb';
    const termSetId = '7a167c47-2b37-41d0-94d0-e962c1a4f2ed';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' && opts.headers && opts.headers['X-RequestDigest']) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{${termGroupId}}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{${termSetId}}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
          return csomDefaultResponseJson;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="${csomDefaultResponseFormatted.filter(y => y.TermsCount > 0)[0]._ObjectIdentity_}" /></ObjectPaths></Request>`) {
          return csomFirstChildTermResponseJson;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="6ede85a0-70c8-6000-02cc-23cfae5cac8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNHOmHw1O1kkuIa5r3F81zsA==" /></ObjectPaths></Request>`) {
          return csomSecondChildTermResponseJson;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { termSetId: termSetId, termGroupId: termGroupId, includeChildTerms: true } });
    assert(loggerLogSpy.calledWith(csomChildResponseFormatted));
  });

  it('gets taxonomy terms including child terms from term set by id, term group by id in a friendly formatted way', async () => {
    const termGroupId = '0e8f395e-ff58-4d45-9ff7-e331ab728beb';
    const termSetId = '7a167c47-2b37-41d0-94d0-e962c1a4f2ed';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' && opts.headers && opts.headers['X-RequestDigest']) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{${termGroupId}}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{${termSetId}}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
          return csomDefaultResponseJson;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="${csomDefaultResponseFormatted.filter(y => y.TermsCount > 0)[0]._ObjectIdentity_}" /></ObjectPaths></Request>`) {
          return csomFirstChildTermResponseJson;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="6ede85a0-70c8-6000-02cc-23cfae5cac8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNHOmHw1O1kkuIa5r3F81zsA==" /></ObjectPaths></Request>`) {
          return csomSecondChildTermResponseJson;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { termSetId: termSetId, termGroupId: termGroupId, includeChildTerms: true, output: 'text' } });
    assert(loggerLogSpy.calledWith(csomChildResponseFormattedText));
  });

  it('gets taxonomy terms from the specified sitecollection, term set by id, term group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/project-x/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/project-x', termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('gets taxonomy terms from term set by id, term group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('gets taxonomy term set by name, term group by id in a friendly way', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { termSetName: 'PnPTermSets', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', output: 'text' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormattedText));
  });

  it('gets taxonomy term set by id, term group by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupName: 'PnPTermSets' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('gets taxonomy term set by name, term group by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('returns all properties for output json', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets', output: 'json' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('escapes XML in term group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets>' } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('escapes XML in term set name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations&gt;</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return csomDefaultResponseJson;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { termSetName: 'PnP-Organizations>', termGroupName: 'PnPTermSets', verbose: true } });
    assert(loggerLogSpy.calledWith(csomDefaultResponseFormatted));
  });

  it('correctly handles term group not found via id', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e"
        }
      ]);
    });

    await assert.rejects(command.action(logger, {
      options: {
        termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed',
        termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term group not found via name', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
        }
      ]);
    });

    await assert.rejects(command.action(logger, {
      options: {
        termSetName: 'PnP-CollabFooter-SharedLinks',
        termGroupName: 'PnPTermSets'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term set not found via id', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d"
        }
      ]);
    });
    await assert.rejects(command.action(logger, {
      options: {
        termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed',
        termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term set not found via name', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
        }
      ]);
    });

    await assert.rejects(command.action(logger, {
      options: {
        termSetName: 'PnP-CollabFooter-SharedLinks',
        termGroupName: 'PnPTermSets'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles error when retrieving taxonomy terms', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]);
    });

    await assert.rejects(command.action(logger, {
      options: {
        termSetName: 'PnP-Organizations',
        termGroupName: 'PnPTermSets'
      }
    } as any), new CommandError('File Not Found.'));
  });

  it('correctly handles no terms found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "10ca969e-3062-0000-2cdb-e38e5b6fba03" }, 70, { "IsNull": false }, 71, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 73, { "IsNull": false }, 74, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh/fzgFZGpUQ==" }, 76, { "IsNull": false }, 78, { "IsNull": false }, 79, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+s=" }, 81, { "IsNull": false }, 83, { "IsNull": false }, 84, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7" }, 86, { "IsNull": false }, 87, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [] }]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets', output: 'json' } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('fails validation if neither termSetId nor termSetName specified', async () => {
    const actual = await command.validate({ options: { termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termSetId and termSetName specified', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termSetName: 'PnP-CollabFooter-SharedLinks', termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termSetId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { termSetId: 'invalid', termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither termGroupId nor termGroupName specified', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termGroupId and termGroupName specified', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and termGroupName specified', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetName and termGroupName specified', async () => {
    const actual = await command.validate({ options: { termSetName: 'People', termGroupName: 'PnPTermSets' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetId and termGroupId specified', async () => {
    const actual = await command.validate({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetName and termGroupId specified', async () => {
    const actual = await command.validate({ options: { termSetName: 'PnP-CollabFooter-SharedLinks', termGroupId: '9e54299e-208a-4000-8546-cc4139091b26' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.reject('getRequestDigest error'));

    await assert.rejects(command.action(logger, {
      options: {
        termSetName: 'PnP-Organizations',
        termGroupName: 'PnPTermSets',
        output: 'json'
      }
    } as any), new CommandError('getRequestDigest error'));
  });
});
