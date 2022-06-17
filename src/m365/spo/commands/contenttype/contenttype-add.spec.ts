import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
import * as SpoContentTypeGetCommand from './contenttype-get';
const command: Command = require('./contenttype-add');

describe(commands.CONTENTTYPE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates site content type with minimal properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"><Property Name="Description" Type="Null" /><Property Name="Group" Type="Null" /><Property Name="Id" Type="String">0x0100FF0B2E33A3718B46A3909298D240FD93</Property><Property Name="Name" Type="String">PnP Tile</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": null, "TraceCorrelationId": "2846869e-a0d0-0000-2105-47de3b2952e7"
          }, 13, {
            "IsNull": false
          }, 14, {
            "_ObjectIdentity_": "2846869e-a0d0-0000-2105-47de3b2952e7|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:contenttype:0x0100FF0B2E33A3718B46A3909298D240FD93"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoContentTypeGetCommand) {
        return Promise.resolve({
          stdout: '{"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"","Scope":"/sites/portal","Sealed":false,"StringId":"0x0100FF0B2E33A3718B46A3909298D240FD93"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { output: "json",  debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } }, () => {
      try {
        assert(loggerLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates site content type with description and group (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"><Property Name="Description" Type="String">A tile</Property><Property Name="Group" Type="String">PnP Content Types</Property><Property Name="Id" Type="String">0x0100FF0B2E33A3718B46A3909298D240FD93</Property><Property Name="Name" Type="String">PnP Tile</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": null, "TraceCorrelationId": "2846869e-a0d0-0000-2105-47de3b2952e7"
          }, 13, {
            "IsNull": false
          }, 14, {
            "_ObjectIdentity_": "2846869e-a0d0-0000-2105-47de3b2952e7|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:contenttype:0x0100FF0B2E33A3718B46A3909298D240FD93"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoContentTypeGetCommand) {
        return Promise.resolve({
          stdout: '{"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"Scope":"/sites/portal","Sealed":false,"StringId":"0x0100FF0B2E33A3718B46A3909298D240FD93"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', description: 'A tile', group: 'PnP Content Types' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('creates list content type with minimal properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('My%20list')?$select=Id`) > -1) {
        return Promise.resolve({
          Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99'
        });
      }

      if ((opts.url as string).indexOf('/_api/site?$select=Id') > -1) {
        return Promise.resolve({
          Id: '276f6d32-f43b-4b26-ada6-7aa9d5bcab6a'
        });
      }

      if ((opts.url as string).indexOf('/_api/web?$select=Id') > -1) {
        return Promise.resolve({
          Id: '942595c1-6100-4ad0-9dd4-19743732ffdc'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"><Property Name="Description" Type="Null" /><Property Name="Group" Type="Null" /><Property Name="Id" Type="String">0x0100FF0B2E33A3718B46A3909298D240FD93</Property><Property Name="Name" Type="String">PnP Tile</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method><Identity Id="5" Name="1a48869e-c092-0000-1f61-81ec89809537|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:list:81f0ecee-75a8-46f0-b384-c8f4f9f31d99" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": null, "TraceCorrelationId": "2846869e-a0d0-0000-2105-47de3b2952e7"
          }, 13, {
            "IsNull": false
          }, 14, {
            "_ObjectIdentity_": "2846869e-a0d0-0000-2105-47de3b2952e7|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:list:81f0ecee-75a8-46f0-b384-c8f4f9f31d99:contenttype:0x0100FF0B2E33A3718B46A3909298D240FD93"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', listTitle: 'My list' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates list content type with description', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('My%20list')?$select=Id`) > -1) {
        return Promise.resolve({
          Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99'
        });
      }

      if ((opts.url as string).indexOf('/_api/site?$select=Id') > -1) {
        return Promise.resolve({
          Id: '276f6d32-f43b-4b26-ada6-7aa9d5bcab6a'
        });
      }

      if ((opts.url as string).indexOf('/_api/web?$select=Id') > -1) {
        return Promise.resolve({
          Id: '942595c1-6100-4ad0-9dd4-19743732ffdc'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"><Property Name="Description" Type="String">A tile</Property><Property Name="Group" Type="Null" /><Property Name="Id" Type="String">0x0100FF0B2E33A3718B46A3909298D240FD93</Property><Property Name="Name" Type="String">PnP Tile</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method><Identity Id="5" Name="1a48869e-c092-0000-1f61-81ec89809537|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:list:81f0ecee-75a8-46f0-b384-c8f4f9f31d99" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": null, "TraceCorrelationId": "2846869e-a0d0-0000-2105-47de3b2952e7"
          }, 13, {
            "IsNull": false
          }, 14, {
            "_ObjectIdentity_": "2846869e-a0d0-0000-2105-47de3b2952e7|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:list:81f0ecee-75a8-46f0-b384-c8f4f9f31d99:contenttype:0x0100FF0B2E33A3718B46A3909298D240FD93"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoContentTypeGetCommand) {
        return Promise.resolve({
          stdout: '{"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"Scope":"/sites/portal","Sealed":false,"StringId":"0x0100FF0B2E33A3718B46A3909298D240FD93"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { debug: false, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', listTitle: 'My list', description: 'A tile' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"><Property Name="Description" Type="String">&lt;A tile</Property><Property Name="Group" Type="String">&lt;PnP Content Types</Property><Property Name="Id" Type="String">&lt;0x0100FF0B2E33A3718B46A3909298D240FD93</Property><Property Name="Name" Type="String">&lt;PnP Tile</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": null, "TraceCorrelationId": "2846869e-a0d0-0000-2105-47de3b2952e7"
          }, 13, {
            "IsNull": false
          }, 14, {
            "_ObjectIdentity_": "2846869e-a0d0-0000-2105-47de3b2952e7|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:276f6d32-f43b-4b26-ada6-7aa9d5bcab6a:web:942595c1-6100-4ad0-9dd4-19743732ffdc:contenttype:0x0100FF0B2E33A3718B46A3909298D240FD93"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoContentTypeGetCommand) {
        return Promise.resolve({
          stdout: '{"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"Scope":"/sites/portal","Sealed":false,"StringId":"0x0100FF0B2E33A3718B46A3909298D240FD93"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: '<PnP Tile', id: '<0x0100FF0B2E33A3718B46A3909298D240FD93', description: '<A tile', group: '<PnP Content Types' } }, () => {
      try {
        assert(loggerLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when a content with the specified name already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8008.1219", "ErrorInfo": {
              "ErrorMessage": "A duplicate content type \"PnP Tile\" was found.", "ErrorValue": null, "TraceCorrelationId": "0e46869e-2024-0000-1f04-7f2be163c9c0", "ErrorCode": 183, "ErrorTypeName": "Microsoft.SharePoint.SPException"
            }, "TraceCorrelationId": "0e46869e-2024-0000-1f04-7f2be163c9c0"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } } as any, (err?: any) => {
      try {
        assert(JSON.stringify(err), JSON.stringify(new CommandError("A duplicate content type \"PnP Tile\" was found.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified list doesn\'t exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('My%20list')?$select=Id`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: "List 'My list' does not exist at site with URL 'https://contoso.sharepoint.com/sites/sales'." } } } });
      }

      if ((opts.url as string).indexOf('/_api/site?$select=Id') > -1) {
        return Promise.resolve({
          Id: '276f6d32-f43b-4b26-ada6-7aa9d5bcab6a'
        });
      }

      if ((opts.url as string).indexOf('/_api/web?$select=Id') > -1) {
        return Promise.resolve({
          Id: '942595c1-6100-4ad0-9dd4-19743732ffdc'
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', listTitle: 'My list' } } as any, (err?: any) => {
      try {
        assert(JSON.stringify(err), JSON.stringify(new CommandError("List 'My list' does not exist at site with URL 'https://contoso.sharepoint.com/sites/sales'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = command.types;
    ['i', 'id'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });
});