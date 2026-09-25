import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './applicationcustomizer-add.js';

describe(commands.APPLICATIONCUSTOMIZER_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const title = 'PageFooter';
  const description = 'Page footer customizer';
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
    Description: description,
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
  let commandOptionsSchema: typeof options;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APPLICATIONCUSTOMIZER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds the application customizer to a specific site without specifying clientSideComponentProperties', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/UserCustomActions') {
        return;
      }

      throw customActionError;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, scope: 'Web' }) });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      Title: title,
      Name: title,
      Description: undefined,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: clientSideComponentId,
      HostProperties: ''
    });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('adds the application customizer to a specific site while specifying clientSideComponentProperties', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Site/UserCustomActions') {
        return customActionAddResponse;
      }

      throw customActionError;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, description: description, clientSideComponentProperties: clientSideComponentProperties, verbose: true }) });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      Title: title,
      Name: title,
      Description: description,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: clientSideComponentId,
      ClientSideComponentProperties: clientSideComponentProperties,
      HostProperties: ''
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('adds the application customizer to a specific site while specifying hostProperties', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Site/UserCustomActions') {
        return customActionAddResponse;
      }

      throw customActionError;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, description: description, hostProperties: clientSideComponentProperties, verbose: true }) });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      Title: title,
      Name: title,
      Description: description,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: clientSideComponentId,
      HostProperties: clientSideComponentProperties
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', title: title, clientSideComponentId: clientSideComponentId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the scope option is not a valid scope', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, scope: 'Invalid scope' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the clientSideComponentProperties option is not a valid json string', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: 'invalid json string' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the hostProperties option is not a valid json string', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, hostProperties: 'invalid json string' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if all options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: clientSideComponentProperties, hostProperties: clientSideComponentProperties, scope: 'Site' });
    assert.strictEqual(actual.success, true);
  });
}); 
