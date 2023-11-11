import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.GROUP_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes the group when id is passed', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, force: true } });
    assert(requestPostSpy.called);
  });

  it('deletes the group when name is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/mysite/_api/web/sitegroups/GetByName('Team Site Owners')?$select=Id`) {
        return {
          Id: 7
        };
      }
      throw 'Invalid request';
    });

    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', name: 'Team Site Owners', debug: true, force: true } });
    assert(requestPostSpy.called);
  });

  it('aborts deleting the group when prompt is not continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } });
    assert(requestPostSpy.notCalled);
  });

  it('deletes the group when prompt is continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } });
    assert(requestPostSpy.called);
  });

  it('correctly handles group remove reject request', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, force: true } } as any),
      new CommandError(error.error['odata.error'].message.value));
  });

  it('prompts before removing group when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: { id: 7, webUrl: 'https://contoso.sharepoint.com/mysite' } });

    assert(promptIssued);
  });

  it('prompts before removing group when confirmation argument not passed (name)', async () => {
    await command.action(logger, { options: { name: 'Team Site Owners', webUrl: 'https://contoso.sharepoint.com/mysite' } });

    assert(promptIssued);
  });

  it('fails validation if both id and name options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 7 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 'Hi' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both id and name options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, name: 'Team Site Members' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
