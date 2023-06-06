import Command, { CommandError } from '../../../../Command';
import commands from '../../commands';
import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
import { sinonUtil } from '../../../../utils/sinonUtil';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { Logger } from '../../../../cli/Logger';
import * as os from 'os';

const command: Command = require('./applicationcustomizer-remove');

describe(commands.APPLICATIONCUSTOMIZER_REMOVE, () => {
  let commandInfo: CommandInfo;
  const webUrl = 'https://contoso.sharepoint.com';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  const clientSideComponentId = '015e0fcf-fe9d-4037-95af-0a4776cdfbb4';
  const title = 'SiteGuidedTour';
  let promptOptions: any;
  let log: any[];
  let logger: Logger;
  let requests: any[];
  const singleResponse = {
    value: [
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "Id": id,
        "ImageUrl": null,
        "Location": "ClientSideExtension.ApplicationCustomizer",
        "Name": title,
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": { "High": 0, "Low": 0 },
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 65536,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": "1.0.1.0"
      }
    ]
  };

  const multipleResponse = {
    value: [
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "'{testMessage:Test message}'",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": '',
        "Id": 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
        "ImageUrl": null,
        "Location": 'ClientSideExtension.ApplicationCustomizer',
        "Name": 'YourName',
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": [Object],
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": '16.0.1.0'
      },
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "'{testMessage:Test message}'",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": '',
        "Id": '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
        "ImageUrl": null,
        "Location": 'ClientSideExtension.ApplicationCustomizer',
        "Name": 'YourName',
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": [Object],
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": '16.0.1.0'
      }
    ]
  };

  const defaultDeleteCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return undefined;
      }
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return undefined;
      }
      return new Error('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.APPLICATIONCUSTOMIZER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if at least one of the parameters has a value', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when all parameters are empty', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: null, clientSideComponentId: null, title: '' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, clientSideComponentId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope option is not a valid scope', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, scope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should prompt before removing application customizer when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, id: id } });
    let promptIssued = false;
    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('aborts removing application customizer when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });
    await command.action(logger, { options: { webUrl: webUrl, id: id } });
    assert(requests.length === 0);
  });

  it('handles error when no user application customizer with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        return { "odata.null": true };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { id: id, webUrl: webUrl, confirm: true }
      }
      ), new CommandError(`No application customizer with id '${id}' found`));
  });

  it('handles error when no user application customizer with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { title: title, webUrl: webUrl, confirm: true }
      }
      ), new CommandError(`No application customizer with title '${title}' found`));
  });

  it('handles error when no user application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, confirm: true }
      }
      ), new CommandError(`No application customizer with ClientSideComponentId '${clientSideComponentId}' found`));
  });

  it('handles error when multiple user application customizer with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { title: title, webUrl: webUrl, scope: 'Site', confirm: true }
      }
      ), new CommandError(`Multiple application customizer with title '${title}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponse.value.map(a => `- ${a.Id}`).join(os.EOL)}`));
  });

  it('handles error when multiple user application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Site', confirm: true }
      }
      ), new CommandError(`Multiple application customizer with ClientSideComponentId '${clientSideComponentId}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponse.value.map(a => `- ${a.Id}`).join(os.EOL)}`));
  });

  it('should remove the application customizer from the site by its ID when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        return singleResponse.value[0];
      }
      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { verbose: true, id: id, webUrl: webUrl, scope: 'Web' } } as any);
    assert(deleteCallsSpy.calledOnce);
  });

  it('should remove the application customizer from the site collection by its ID when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        const response = singleResponse.value[0];
        response.Scope = 2;
        return response;
      }
      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    await command.action(logger, { options: { verbose: true, id: id, webUrl: webUrl, scope: 'Site', confirm: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });

  it('should remove the application customizer from the site by its clientSideComponentId when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return singleResponse;
      }
      else if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    await command.action(logger, { options: { verbose: true, clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Web', confirm: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });
});