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
import command from './dataverse-webresource-list.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.DATAVERSE_WEBRESOURCE_LIST, () => {
  const envResponse: any = { 'properties': { 'linkedEnvironmentMetadata': { 'instanceApiUrl': 'https://contoso-dev.api.crm4.dynamics.com' } } };
  const solutionResponse: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#solutions',
    '@odata.count': 1,
    'value': [
      {
        'solutionid': '00000001-0000-0000-0001-00000000009b',
        'uniquename': 'Crc00f1',
        'version': '1.0.0.0',
        'installedon': '2021-10-01T21:54:14Z',
        'solutionpackageversion': null,
        'friendlyname': 'Common Data Services Default Solution',
        'versionnumber': 860052,
        'publisherid': {
          'friendlyname': 'CDS Default Publisher',
          'publisherid': '00000001-0000-0000-0000-00000000005a'
        }
      }
    ]
  };
  const emptySolutionResponse: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#solutions',
    '@odata.count': 0,
    'value': []
  };
  const solutionComponentResponse: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#solutioncomponents',
    'value': [
      {
        '@odata.etag': 'W/\'251809901\'',
        'objectid': 'b6a1791b-7706-e711-aff7-000c29480724',
        'solutioncomponentid': '4c515176-f689-ee11-8179-000d3adf73a3'
      },
      {
        '@odata.etag': 'W/\'251809902\'',
        'objectid': '0301083b-b5d0-e911-a978-000d3ab5a0d7',
        'solutioncomponentid': '4d515176-f689-ee11-8179-000d3adf73a3'
      },
      {
        '@odata.etag': 'W/\'225188450\'',
        'objectid': 'f519e0a0-c5d3-e911-a973-000d3ab5a6ae',
        'solutioncomponentid': '1bb35189-aede-ed11-a7c7-000d3adf73a3'
      }
    ]
  };
  const webResourceResponseAll: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#webresourceset',
    'value': [
      {
        'webresourceid': 'b6a1791b-7706-e711-aff7-000c29480724',
        'name': 'mock_myJSresource.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'content': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'content_binary': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': false,
        'isAvailableForMobileOfflineLabel': 'No',
        'isenabledformobileclient': false,
        'isEnabledForMobileClientLabel': 'No',
        'ismanaged': false,
        'isManagedLabel': 'Unmanaged',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      },
      {
        'webresourceid': '0301083b-b5d0-e911-a978-000d3ab5a0d7',
        'name': 'mock_myJSresource1.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'content': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'content_binary': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource1',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': false,
        'isAvailableForMobileOfflineLabel': 'No',
        'isenabledformobileclient': false,
        'isEnabledForMobileClientLabel': 'No',
        'ismanaged': true,
        'isManagedLabel': 'Managed',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      },
      {
        'webresourceid': 'f519e0a0-c5d3-e911-a973-000d3ab5a6ae',
        'name': 'mock_myJSresource2.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'content': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'content_binary': 'Ww0KICAgIHsNCiAgICAgICAgInBhcmVudCI6ICJtb2NrX0luc3RhbGxhdGlvbkRlcGVuZGVudE9wdGlvblNldCIsDQogICAgICAgICJjaGlsZCI6ICJtb2NrX3NpemVjbGFzcyIsDQogICAgICAgICJvcHRpb25zIjogew0KICAgICAgICAgICAiNzUwNjEwMDAwIjogWw0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDAiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDEiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDIiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDMiLA0KICAgICAgICAgICAgICAgICI3NTA2MTAwMDQiDQogICAgICAgICAgICBdLA0KICAgICAgICAgICAgIjc1MDYxMDAwMSI6IFsNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA1IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA2IiwNCiAgICAgICAgICAgICAgICAiNzUwNjEwMDA3Ig0KICAgICAgICAgICAgXQ0KICAgICAgICB9DQogICAgfQ0KXQ==',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource2',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': true,
        'isAvailableForMobileOfflineLabel': 'Yes',
        'isenabledformobileclient': true,
        'isEnabledForMobileClientLabel': 'Yes',
        'ismanaged': false,
        'isManagedLabel': 'Unmanaged',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      }
    ]
  };
  const webResourceResponseWithoutContent: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#webresourceset',
    'value': [
      {
        'webresourceid': 'b6a1791b-7706-e711-aff7-000c29480724',
        'name': 'mock_myJSresource.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': false,
        'isAvailableForMobileOfflineLabel': 'No',
        'isenabledformobileclient': false,
        'isEnabledForMobileClientLabel': 'No',
        'ismanaged': false,
        'isManagedLabel': 'Unmanaged',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      },
      {
        'webresourceid': '0301083b-b5d0-e911-a978-000d3ab5a0d7',
        'name': 'mock_myJSresource1.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource1',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': false,
        'isAvailableForMobileOfflineLabel': 'No',
        'isenabledformobileclient': false,
        'isEnabledForMobileClientLabel': 'No',
        'ismanaged': false,
        'isManagedLabel': 'Unmanaged',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      },
      {
        'webresourceid': 'f519e0a0-c5d3-e911-a973-000d3ab5a6ae',
        'name': 'mock_myJSresource2.js',
        'componentstate': 0,
        'componentStateLabel': 'Published',
        'contentfileref': null,
        'contentjson': null,
        'contentjsonfileref': null,
        'createdon': '2024-02-05T09:30:15Z',
        'dependencyxml': null,
        'description': null,
        'displayname': 'Mock My JS Resource2',
        'introducedversion': '1.0.0.60',
        'isavailableformobileoffline': false,
        'isAvailableForMobileOfflineLabel': 'No',
        'isenabledformobileclient': false,
        'isEnabledForMobileClientLabel': 'No',
        'ismanaged': false,
        'isManagedLabel': 'Unmanaged',
        'languagecode': 0,
        'modifiedon': '2024-02-05T10:15:22Z',
        'overwritetime': '1900-01-01T00:00:00Z',
        'silverlightversion': null,
        'solutionid': 'f62e2c8d-1eac-4d98-9a13-8d0e63a2d28c',
        'versionnumber': 123456789,
        'webresourceidunique': 'a1b2c3d4-e5f6-4a5b-8c9d-1e2f3a4b5c6d',
        'webresourcetype': 3,
        'webresourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        },
        'iscustomizable': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'iscustomizableanddeletable'
        },
        'ishidden': {
          'Value': false,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'ishidden'
        }
      }
    ]
  };
  const webResourceResponseTrimmed: any = {
    '@odata.context': 'https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/$metadata#webresourceset',
    'value': [
      {
        'displayname': 'Mock My JS Resource',
        'ismanaged': false,
        'webresourcetype': 3,
        'webResourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        }
      },
      {
        'displayname': 'Mock My JS Resource1',
        'ismanaged': false,
        'webresourcetype': 3,
        'webResourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        }
      },
      {
        'displayname': 'Mock My JS Resource2',
        'ismanaged': true,
        'webresourcetype': 3,
        'webResourceTypeLabel': 'Script (JScript)',
        'canbedeleted': {
          'Value': true,
          'CanBeChanged': true,
          'ManagedPropertyLogicalName': 'canbedeleted'
        }
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  beforeEach(() => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/solutions?$filter=solutionid eq 00000001-0000-0000-0001-00000000009b&$count=true`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/solutions?$filter=solutionid eq 00000001-0000-0000-0001-00000000009c&$count=true`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return emptySolutionResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/solutioncomponents?$filter=_solutionid_value eq 00000001-0000-0000-0001-00000000009b and componenttype eq 61&$select=objectid`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionComponentResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/webresourceset?$select=webresourceid,name,canbedeleted,componentstate,content,content_binary,contentfileref,contentjson,contentjsonfileref,createdon,dependencyxml,description,displayname,introducedversion,isavailableformobileoffline,iscustomizable,isenabledformobileclient,ishidden,ismanaged,languagecode,modifiedon,overwritetime,silverlightversion,solutionid,versionnumber,webresourceidunique,webresourcetype&$filter=webresourceid eq b6a1791b-7706-e711-aff7-000c29480724 or webresourceid eq 0301083b-b5d0-e911-a978-000d3ab5a0d7 or webresourceid eq f519e0a0-c5d3-e911-a973-000d3ab5a6ae`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webResourceResponseAll;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/webresourceset?$select=webresourceid,name,canbedeleted,componentstate,contentfileref,contentjson,contentjsonfileref,createdon,dependencyxml,description,displayname,introducedversion,isavailableformobileoffline,iscustomizable,isenabledformobileclient,ishidden,ismanaged,languagecode,modifiedon,overwritetime,silverlightversion,solutionid,versionnumber,webresourceidunique,webresourcetype&$filter=webresourceid eq b6a1791b-7706-e711-aff7-000c29480724 or webresourceid eq 0301083b-b5d0-e911-a978-000d3ab5a0d7 or webresourceid eq f519e0a0-c5d3-e911-a973-000d3ab5a6ae`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webResourceResponseWithoutContent;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/webresourceset?$select=displayname,webresourcetype,ismanaged,canbedeleted&$filter=webresourceid eq b6a1791b-7706-e711-aff7-000c29480724 or webresourceid eq 0301083b-b5d0-e911-a978-000d3ab5a0d7 or webresourceid eq f519e0a0-c5d3-e911-a973-000d3ab5a6ae`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webResourceResponseTrimmed;
        }
      }

      throw {
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      };
    });

    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;

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
    sinon.restore();
    auth.service.connected = false;
    sinonUtil.restore([
      request.get
    ]);
  });

  describe('metadata', () => {
    it('has correct name', () => {
      assert.strictEqual(command.name.startsWith(commands.DATAVERSE_WEBRESOURCE_LIST), true);
    });

    it('has a description', () => {
      assert.notStrictEqual(command.description, null);
    });

    it('defines correct properties for the default output', () => {
      assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'webresourceType', 'webresourceTypeLabel', 'isManaged', 'isManagedLabel', 'canBeDeleted']);
    });
  });

  describe('list webresources', () => {
    it('it retrieves list of web resources for solution in environment if format json', async () => {
      await command.action(logger, {
        options: {
          environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36',
          solutionId: '00000001-0000-0000-0001-00000000009b',
          output: 'json'
        }
      });

      assert.strictEqual(loggerLogSpy.callCount, 1);
      assert(loggerLogSpy.calledWithExactly(webResourceResponseAll.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseWithoutContent.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseTrimmed.value));
    });

    it('it retrieves list of web resources for solution in environment without the file content', async () => {
      await command.action(logger, {
        options: {
          environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36',
          solutionId: '00000001-0000-0000-0001-00000000009b',
          output: 'json',
          excludeContent: true
        }
      });

      assert.strictEqual(loggerLogSpy.callCount, 1);
      assert(loggerLogSpy.calledWithExactly(webResourceResponseWithoutContent.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseAll.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseTrimmed.value));
    });

    it('it retrieves list of web resources for solution in environment in format text', async () => {
      await command.action(logger, {
        options: {
          environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36',
          solutionId: '00000001-0000-0000-0001-00000000009b',
          output: 'text'
        }
      });
      assert.strictEqual(loggerLogSpy.callCount, 1);
      assert(loggerLogSpy.neverCalledWith(webResourceResponseWithoutContent.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseTrimmed.value));
      assert(loggerLogSpy.neverCalledWith(webResourceResponseAll.value));

      const output = [
        {
          displayName: 'Mock My JS Resource',
          webresourceType: 3,
          webresourceTypeLabel: 'Script (JScript)',
          isManaged: false,
          isManagedLabel: 'Unmanaged',
          canBeDeleted: true
        },
        {
          displayName: 'Mock My JS Resource1',
          webresourceType: 3,
          webresourceTypeLabel: 'Script (JScript)',
          isManaged: false,
          isManagedLabel: 'Unmanaged',
          canBeDeleted: true
        },
        {
          displayName: 'Mock My JS Resource2',
          webresourceType: 3,
          webresourceTypeLabel: 'Script (JScript)',
          isManaged: true,
          isManagedLabel: 'Managed',
          canBeDeleted: true
        }
      ];
      assert(loggerLogSpy.calledWithExactly(output));
    });
  });

  describe('validation', () => {
    let commandInfo: CommandInfo;
    const incorrectGuid = '4be50206-9576-4237-8b17-';
    const correctGuid = '4be50206-9576-4237-8b17-38d8aadfaa41';


    before(() => {
      commandInfo = cli.getCommandInfo(command);
    });

    it('fails validation if environmentName is not a valid guid.', async () => {
      const actual = await command.validate({
        options: {
          environmentName: incorrectGuid,
          solutionId: correctGuid
        }
      }, commandInfo);
      assert.strictEqual(actual, `The value provided as environmentName \'${incorrectGuid}\' is not a valid GUID`);
    });

    it('passes validation if both solutionId and environmentName is valid guid.', async () => {
      const actual = await command.validate({
        options: {
          environmentName: correctGuid,
          solutionId: correctGuid
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    });

    it('fails validation if solutionId is not a valid guid.', async () => {
      const actual = await command.validate({
        options: {
          environmentName: correctGuid,
          solutionId: incorrectGuid
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    });

  });

  describe('error handling', () => {
    it('it correctly handles no environments', async () => {
      try {
        await command.action(logger, {
          options: {
            debug: true,
            environmentName: '4be50206-9576-4237-8b17-38d8aadfaa41',
            solutionId: '00000001-0000-0000-0001-00000000009b'
          }
        });
      }
      catch (err: any) {
        assert.match(err.message, /The environment \'4be50206-9576-4237-8b17-38d8aadfaa41\' could not be retrieved./);
      }
      assert(loggerLogSpy.notCalled);
    });

    it('it correctly handles no solutions', async () => {
      try {
        await command.action(logger, {
          options: {
            debug: true,
            environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36',
            solutionId: '00000001-0000-0000-0001-00000000009c'
          }
        });
      }
      catch (err: any) {
        assert.strictEqual(err.message, 'Solution with ID \'00000001-0000-0000-0001-00000000009c\' not found in environment \'4be50206-9576-4237-8b17-38d8aadfaa36\'');
      }
      assert(loggerLogSpy.notCalled);
    });

    it('it correctly handles API OData error', async () => {
      await assert.rejects(command.action(logger, {
        options: {
        }
      }));
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`);
    });
  });
});
