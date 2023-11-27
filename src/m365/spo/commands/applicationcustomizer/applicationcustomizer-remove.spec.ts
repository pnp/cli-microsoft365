import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './applicationcustomizer-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APPLICATIONCUSTOMIZER_REMOVE, () => {
  let commandInfo: CommandInfo;
  const webUrl = 'https://contoso.sharepoint.com';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  const clientSideComponentId = '015e0fcf-fe9d-4037-95af-0a4776cdfbb4';
  const title = 'SiteGuidedTour';
  let cli: Cli;
  let promptIssued: boolean = false;
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
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
    requests = [];
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound,
      Cli.promptForConfirmation,
      Cli.handleMultipleResultsFound
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    assert(promptIssued);
  });

  it('aborts removing application customizer when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
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
        options: { id: id, webUrl: webUrl, force: true }
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
        options: { title: title, webUrl: webUrl, force: true }
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
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, force: true }
      }
      ), new CommandError(`No application customizer with ClientSideComponentId '${clientSideComponentId}' found`));
  });

  it('handles error when multiple user application customizer with the specified title found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { title: title, webUrl: webUrl, scope: 'Site', force: true }
      }
      ), new CommandError("Multiple application customizer with title 'SiteGuidedTour' found. Found: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59."));
  });

  it('handles error when multiple user application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Site', force: true }
      }
      ), new CommandError("Multiple application customizer with ClientSideComponentId '015e0fcf-fe9d-4037-95af-0a4776cdfbb4' found. Found: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59."));
  });

  it('handles selecting single result when multiple application customizers with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return multipleResponse;
      }
      else if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleResponse.value[0]);

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    await command.action(logger, { options: { verbose: true, title: title, webUrl: webUrl, scope: 'Web', force: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });

  it('should remove the application customizer from the site by its ID when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        return singleResponse.value[0];
      }
      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
    await command.action(logger, { options: { verbose: true, id: id, webUrl: webUrl, scope: 'Site', force: true } } as any);
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
    await command.action(logger, { options: { verbose: true, clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Web', force: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });
});