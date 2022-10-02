import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./group-remove');

describe(commands.GROUP_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let trackEvent: any;
  let telemetry: any;
  let promptOptions: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, confirm: true } });
    assert(trackEvent.called);
  });

  it('logs correct telemetry event', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, confirm: true } });
    assert.strictEqual(telemetry.name, commands.GROUP_REMOVE);
  });

  it('deletes the group when id is passed', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, confirm: true } });
    assert(requestPostSpy.called);
  });

  it('deletes the group when name is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/mysite/_api/web/sitegroups/GetByName('Team Site Owners')?$select=Id`) {
        return Promise.resolve({
          Id: 7
        });
      }
      return Promise.reject('Invalid Request');
    });

    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', name: 'Team Site Owners', debug: true, confirm: true } });
    assert(requestPostSpy.called);
  });

  it('aborts deleting the group when prompt is not continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } });
    assert(requestPostSpy.notCalled);
  });

  it('deletes the group when prompt is continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } });
    assert(requestPostSpy.called);
  });

  it('correctly handles group remove reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.reject(err);
      }
      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, confirm: true } } as any),
      new CommandError(err));
  });

  it('prompts before removing group when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: { debug: false, id: 7, webUrl: 'https://contoso.sharepoint.com/mysite' } });
    let promptIssued = false;
    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing group when confirmation argument not passed (name)', async () => {
    await command.action(logger, { options: { debug: false, name: 'Team Site Owners', webUrl: 'https://contoso.sharepoint.com/mysite' } });
    let promptIssued = false;
    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both id and name options are not passed', async () => {
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
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, name: 'Team Site Members' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});