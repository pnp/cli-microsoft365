import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './contenttype-field-set.js';

describe(commands.CONTENTTYPE_FIELD_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).requestDigest = '';
    (command as any).siteId = '';
    (command as any).webId = '';
    (command as any).fieldLink = null;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPE_FIELD_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a field reference to content type updating field schema', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return;
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds a field reference to content type updating field schema (debug)', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return;
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`;
        }

        throw `[
  {
    "SchemaVersion": "15.0.0.0",
    "LibraryVersion": "16.0.7911.1206",
    "ErrorInfo": {
      "ErrorMessage": "Invalid request",
      "ErrorValue": null,
      "TraceCorrelationId": "59577d9e-70af-0000-22fb-870cf639feff",
      "ErrorCode": -2130575252,
      "ErrorTypeName": "InvalidRequest"
    },
    "TraceCorrelationId": "59577d9e-70af-0000-22fb-870cf639feff"
  }
]`;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('adds a field reference to content type updating field schema from AllowDeletion=FALSE', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"FALSE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return;
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds a field reference to content type without updating field schema', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds a field reference to content type without updating field schema (debug)', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('handles error while updating field schema', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('handles error while adding field reference', async () => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
  {
    "SchemaVersion": "15.0.0.0",
    "LibraryVersion": "16.0.7911.1206",
    "ErrorInfo": {
      "ErrorMessage": "An error has occurred",
      "ErrorValue": null,
      "TraceCorrelationId": "1e5a7d9e-9047-0000-22fb-8361c9a5b96e",
      "ErrorCode": -2130575252,
      "ErrorTypeName": "Error"
    },
    "TraceCorrelationId": "1e5a7d9e-9047-0000-22fb-8361c9a5b96e"
  }
]`;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('updates existing field link', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates existing field link (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('updates existing field link (hidden)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: false, hidden: true } }));
    assert(loggerLogSpy.notCalled);
  });

  it('updates existing field link (required)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: false } }));
    assert(loggerLogSpy.notCalled);
  });

  it('handles error while trying to retrieve field link', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          'odata.null': true
        };
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } } as any),
      new CommandError(`Couldn't find field link for field 5ee2dd25-d941-455a-9bdb-7f2c54aed11b`));
  });

  it('skips updating when existing field link is up-to-date (no values specified)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert(loggerLogSpy.notCalled);
  });

  it('skips updating when existing field link is up-to-date (no values specified; debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('handles error while updating the field link', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return {
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        };
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        };
      }

      if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
        return {
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        };
      }

      if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
        return {
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return `[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": {
        "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
      },
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    }
  ]`;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true } } as any),
      new CommandError(`Unknown Error`));
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified field ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the required option is set to true', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the required option is set to false', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the hidden option is set to true', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', hidden: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the hidden option is set to false', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', hidden: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures contentTypeId as string option', () => {
    const types = command.types;
    ['contentTypeId', 'c'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('handles error while retrieving request digest', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } }));
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return {
              'odata.null': true
            };
          case 2:
            return {
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            };
        }
      }

      if ((opts.url as string).indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return {
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', id: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: true, hidden: true } } as any),
      new CommandError('An error has occurred'));
  });
});