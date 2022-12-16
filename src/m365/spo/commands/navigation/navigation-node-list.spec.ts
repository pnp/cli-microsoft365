import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./navigation-node-list');

describe(commands.NAVIGATION_NODE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.NAVIGATION_NODE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets nodes from the top navigation', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.resolve({ value: [{ "Id": 2003, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert(loggerLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
  });

  it('gets nodes from the quick launch', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/quicklaunch`) > -1) {
        return Promise.resolve({ value: [{ "Id": 2003, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "IsDocLib": true, "IsExternal": false, "IsVisible": true, "ListTemplateType": 0, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } });
    assert(loggerLogSpy.calledWith([{ "Id": 2003, "Title": "Node 1", "Url": "/sites/team-a/SitePages/page1.aspx" }, { "Id": 2004, "Title": "Node 2", "Url": "/sites/team-a/SitePages/page2.aspx" }]));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject({ error: 'An error has occurred' });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
