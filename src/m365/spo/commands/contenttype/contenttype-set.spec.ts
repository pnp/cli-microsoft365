import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './contenttype-set.js';

describe(commands.CONTENTTYPE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const siteId = 'c119e182-eabc-4454-8f1e-6b39551586a7';
  const webId = '5d50a096-7973-4838-85bd-ead8e9a75f2f';
  const listId = '00000000-0000-0000-0000-000000000000';
  const listTitle = 'Assets';
  const listUrl = '/sites/project-x/Lists/Assets';
  const id = '0x0101';
  const name = 'Asset';
  const newName = 'New asset name';

  const contentTypesResponse = {
    value: [
      {
        Name: name,
        Group: 'Custom group',
        Id: {
          StringValue: id
        }
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown options', () => {
    assert.strictEqual(command.allowUnknownOptions(), true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is specified and is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listTitle is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: listId, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listUrl is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: listId, listUrl: listUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle and listUrl is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listTitle: listTitle, listUrl: listUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId, listUrl and is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: listId, listUrl: listUrl, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId is specified together with updateChildren', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: listId, updateChildren: true });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle is specified together with updateChildren', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listTitle: listTitle, updateChildren: true });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listUrl is specified together with updateChildren', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listUrl: listUrl, updateChildren: true });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when webUrl, id and listId are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, listId: listId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when webUrl, name are specified together with updateChildren', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, name: name, updateChildren: true });
    assert.strictEqual(actual.success, true);
  });

  it('correctly updates content type with id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return `[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`;
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, Name: newName, updateChildren: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates content type with name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }
      else if (opts.url === `${webUrl}/_api/Web/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return contentTypesResponse;
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return (`[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`);
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, name: name, Name: newName, updateChildren: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates content type with name and listId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }
      else if (opts.url === `${webUrl}/_api/Web/Lists/GetById('${formatting.encodeQueryParameter(listId)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return contentTypesResponse;
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return `[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`;
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, name: name, listId: listId, Name: newName, updateChildren: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates content type with name and listTitle', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }
      else if (opts.url === `${webUrl}/_api/Web/Lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return contentTypesResponse;
      }
      else if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$select=Id`) {
        return { Id: listId };
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return `[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`;
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, name: name, listTitle: listTitle, Name: newName, updateChildren: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates content type with name and listUrl', async () => {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }
      else if (opts.url === `${webUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return contentTypesResponse;
      }
      else if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
        return { Id: listId };
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return `[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`;
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, name: name, listUrl: listUrl, Name: newName, updateChildren: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates content type with id and pushing changes to children', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
        return `[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`;
      }

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, Name: newName, updateChildren: true }) });
    assert(loggerLogSpy.notCalled);
  });

  it('fails to update content type with name and listUrl when content type does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return { value: [] };
      }

      throw 'Invalid request url: ' + opts.url;
    });
    const patchStub = sinon.stub(request, 'patch').resolves();

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, name: name, Name: newName }) }), new CommandError(`The specified content type '${name}' does not exist`));
    assert(patchStub.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/site?$select=Id`) {
        return { Id: siteId };
      }
      else if (opts.url === `${webUrl}/_api/web?$select=Id`) {
        return { Id: webId };
      }

      throw `Invalid GET-request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`
        && opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="12" ObjectPathId="9" Name="Name"><Parameter Type="String">${newName}</Parameter></SetProperty><Method Name="Update" Id="13" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(id)}" /></ObjectPaths></Request>`) {
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

      throw `Invalid POST-request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, Name: newName }) }), new CommandError('Unknown Error'));
  });
});