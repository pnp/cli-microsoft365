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
import { Logger } from '../../../../cli/Logger';

const command: Command = require('./applicationcustomizer-remove');

describe(commands.APPLICATIONCUSTOMIZER_REMOVE, () => {
  let commandInfo: CommandInfo;
  const webUrl = 'https://contoso.sharepoint.com';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  let promptOptions: any;
  let log: any[];
  let logger: Logger;
  let requests: any[];

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
    assert.strictEqual(command.name.startsWith(commands.APPLICATIONCUSTOMIZER_REMOVE), true);
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
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, clientSideComponentId: null, title: '' } }, commandInfo);
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
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: webUrl, id: id } });
    assert(requests.length === 0);
  });

  it('should remove the application customizer from the site by its ID when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Location eq ') > -1) {
        return {
          value: [
            {
              "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
              "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
              "CommandUIExtension": null,
              "Description": null,
              "Group": null,
              "Id": id,
              "ImageUrl": null,
              "Location": "ClientSideExtension.ApplicationCustomizer",
              "Name": "{b2307a39-e878-458b-bc90-03bc578531d6}",
              "RegistrationId": null,
              "RegistrationType": 0,
              "Rights": { "High": 0, "Low": 0 },
              "Scope": "1",
              "ScriptBlock": null,
              "ScriptSrc": null,
              "Sequence": 65536,
              "Title": "Places",
              "Url": null,
              "VersionOfUserCustomAction": "1.0.1.0"
            }
          ]
        };
      }
      throw new Error('Invalid request');
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { id: id, webUrl: webUrl, scope: 'Web', confirm: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });


  it('should remove the application customizer from the site collection by its ID when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Location eq ') > -1) {
        return {
          value: [
            {
              "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
              "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
              "CommandUIExtension": null,
              "Description": null,
              "Group": null,
              "Id": "015e0fcf-fe9d-4037-95af-0a4776cdfbb5",
              "ImageUrl": null,
              "Location": "ClientSideExtension.ApplicationCustomizer",
              "Name": "{b2307a39-e878-458b-bc90-03bc578531d6}",
              "RegistrationId": null,
              "RegistrationType": 0,
              "Rights": { "High": 0, "Low": 0 },
              "Scope": "2",
              "ScriptBlock": null,
              "ScriptSrc": null,
              "Sequence": 65536,
              "Title": "Places",
              "Url": null,
              "VersionOfUserCustomAction": "1.0.1.0"
            }
          ]
        };
      }
      return new Error('Invalid request');
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { id: "015e0fcf-fe9d-4037-95af-0a4776cdfbb5", webUrl: webUrl, scope: 'Site' } } as any);
    assert(deleteCallsSpy.calledOnce);
  });

  it('handles error when no user application customizer with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Location eq ') > -1) {
        return {
          value: [
          ]
        };
      }
      return new Error('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(
      command.action(logger, {
        options: { id: id, webUrl: webUrl, scope: 'Site' }
      }
      ), new CommandError(`No application customizer found`));
  });

  it('handles error when multiple user application customizer with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Location eq ') > -1) {
        return {
          value: [
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        };
      }
      return new Error('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(
      command.action(logger, {
        options: { title: 'YourAppCustomizer', webUrl: webUrl, scope: 'Site' }
      }
      ), new CommandError(`Multiple application customizer found. Please disambiguate using IDs: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59`));
  });

  it('should remove the application customizer from the site by its clientSideComponentId when the prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Location eq ') > -1) {
        return {
          value: [
            {
              "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
              "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
              "CommandUIExtension": null,
              "Description": null,
              "Group": null,
              "ImageUrl": null,
              "Location": "ClientSideExtension.ApplicationCustomizer",
              "Name": "{b2307a39-e878-458b-bc90-03bc578531d6}",
              "RegistrationId": null,
              "RegistrationType": 0,
              "Rights": { "High": 0, "Low": 0 },
              "Scope": "1",
              "ScriptBlock": null,
              "ScriptSrc": null,
              "Sequence": 65536,
              "Title": "Places",
              "Url": null,
              "VersionOfUserCustomAction": "1.0.1.0"
            }
          ]
        };
      }
      return new Error('Invalid request');
    });

    const deleteCallsSpy: sinon.SinonStub = defaultDeleteCallsStub();
    await command.action(logger, { options: { clientSideComponentId: "015e0fcf-fe9d-4037-95af-0a4776cdfbb4", webUrl: webUrl, scope: 'Web', confirm: true } } as any);
    assert(deleteCallsSpy.calledOnce);
  });
});