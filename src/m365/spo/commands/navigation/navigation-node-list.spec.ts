import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './navigation-node-list.js';

describe(commands.NAVIGATION_NODE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const navigationNodeResponse = {
    value: [
      {
        "Id": 2003,
        "IsDocLib": true,
        "IsExternal": false,
        "IsVisible": true,
        "ListTemplateType": 0,
        "Title": "Node 1",
        "Url": "/sites/team-a/SitePages/page1.aspx",
        "Children": [
          {
            "AudienceIds": null,
            "CurrentLCID": 1033,
            "Id": 2005,
            "IsDocLib": true,
            "IsExternal": true,
            "IsVisible": true,
            "ListTemplateType": 0,
            "Title": "External site",
            "Url": "https://externalsite.com",
            "Children": []
          }
        ]
      },
      {
        "Id": 2004,
        "IsDocLib": true,
        "IsExternal": false,
        "IsVisible": true,
        "ListTemplateType": 0,
        "Title": "Node 2",
        "Url": "/sites/team-a/SitePages/page2.aspx"
      }
    ]
  };

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets nodes from the top navigation', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/topnavigationbar?$expand=Children,Children/Children,Children/Children/Children') {
        return navigationNodeResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert(loggerLogSpy.calledOnceWith(navigationNodeResponse.value));
  });

  it('gets nodes from the quick launch', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/quicklaunch?$expand=Children,Children/Children,Children/Children/Children') {
        return navigationNodeResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } });
    assert(loggerLogSpy.calledOnceWith(navigationNodeResponse.value));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        code: "-2147024891, System.UnauthorizedAccessException",
        message: "Attempted to perform an unauthorized operation."
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any),
      new CommandError('Attempted to perform an unauthorized operation.'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', location: 'TopNavigationBar' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' });
    assert.strictEqual(actual.success, true);
  });
});
