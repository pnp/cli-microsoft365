import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./applicationcustomizer-add');

describe(commands.APPLICATIONCUSTOMIZER_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const title = 'PageFooter';
  const clientSideComponentId = '76d5f8c8-6228-4df8-a2da-b94cbc8115bc';
  const clientSideComponentProperties = '{"testMessage":"Test message"}';
  const customActionError = {
    "url": "https://contoso.sharepoint.com/_api/Web/UserCustomActions",
    "status": 400,
    "statusText": "Bad Request"
  };
  const customActionAddResponse = {
    ClientSideComponentId: '799883f5-7962-4384-a10a-105adaec6ffc',
    ClientSideComponentProperties: '',
    CommandUIExtension: null,
    Description: null,
    Group: null,
    Id: 'bdcea35f-d5d9-45a2-a075-4d1e2f519e74',
    ImageUrl: null,
    Location: 'ClientSideExtension.ApplicationCustomizer',
    Name: 'Some customizer',
    RegistrationId: null,
    RegistrationType: 0,
    Rights: '{"High":"0","Low":"0"}',
    Scope: 'Web',
    ScriptBlock: null,
    ScriptSrc: null,
    Sequence: 0,
    Title: 'Some customizer',
    Url: null,
    VersionOfUserCustomAction: '16.0.1.0'
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APPLICATIONCUSTOMIZER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds the application customizer to a specific site without specifying clientSideComponentProperties', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/UserCustomActions'
        && opts.data['Location'] === 'ClientSideExtension.ApplicationCustomizer'
        && opts.data['ClientSideComponentId'] === clientSideComponentId
        && opts.data['Name'] === title
        && opts.data['ClientSideComponentProperties'] === undefined) {
        return;
      }

      throw customActionError;
    });

    await command.action(logger, { options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, scope: 'Web' } } as any);
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('adds the application customizer to a specific site while specifying clientSideComponentProperties', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Site/UserCustomActions'
        && opts.data['Location'] === 'ClientSideExtension.ApplicationCustomizer'
        && opts.data['ClientSideComponentId'] === clientSideComponentId
        && opts.data['ClientSideComponentProperties'] === clientSideComponentProperties
        && opts.data['Name'] === title) {
        return customActionAddResponse;
      }

      throw customActionError;
    });

    await command.action(logger, { options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: clientSideComponentProperties, verbose: true } } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', title: title, clientSideComponentId: clientSideComponentId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope option is not a valid scope', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, scope: 'Invalid scope' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentProperties option is not a valid json string', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: 'invalid json string' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: clientSideComponentProperties, scope: 'Site' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 
