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

  const menuStateResponse = {
    AudienceIds: [],
    FriendlyUrlPrefix: '',
    IsAudienceTargetEnabledForGlobalNav: false,
    Nodes: [
      {
        AudienceIds: [],
        CurrentLCID: 1033,
        CustomProperties: [],
        FriendlyUrlSegment: '',
        IsDeleted: false,
        IsHidden: false,
        IsTitleForExistingLanguage: false,
        Key: '2003',
        Nodes: [
          {
            AudienceIds: [],
            CurrentLCID: 1033,
            CustomProperties: [],
            FriendlyUrlSegment: '',
            IsDeleted: false,
            IsHidden: false,
            IsTitleForExistingLanguage: false,
            Key: '2005',
            Nodes: [],
            NodeType: 0,
            OpenInNewWindow: true,
            SimpleUrl: 'https://externalsite.com',
            Title: 'External site',
            Translations: []
          }
        ],
        NodeType: 0,
        OpenInNewWindow: null,
        SimpleUrl: '/sites/team-a/SitePages/page1.aspx',
        Title: 'Node 1',
        Translations: []
      },
      {
        AudienceIds: [],
        CurrentLCID: 1033,
        CustomProperties: [],
        FriendlyUrlSegment: '',
        IsDeleted: false,
        IsHidden: false,
        IsTitleForExistingLanguage: false,
        Key: '2004',
        Nodes: [],
        NodeType: 0,
        OpenInNewWindow: false,
        SimpleUrl: '/sites/team-a/SitePages/page2.aspx',
        Title: 'Node 2',
        Translations: []
      }
    ],
    SimpleUrl: '',
    SPSitePrefix: '/sites/team-a',
    SPWebPrefix: '/sites/team-a',
    StartingNodeKey: '1025',
    StartingNodeTitle: 'Quick launch',
    Version: '2026-02-21T00:12:03.6643426Z'
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
      request.post
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

  it('has correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Key', 'Title', 'SimpleUrl']);
  });

  it('gets nodes from the top navigation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/navigation/MenuState' &&
        opts.data.menuNodeKey === '1002') {
        return menuStateResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert(loggerLogSpy.calledOnceWith(menuStateResponse.Nodes));
  });

  it('gets nodes from the quick launch', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/navigation/MenuState' &&
        opts.data.menuNodeKey === null) {
        return menuStateResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } });
    assert(loggerLogSpy.calledOnceWith(menuStateResponse.Nodes));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects({
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
    assert.strictEqual(actual.error?.issues[0].message, `'invalid' is not a valid SharePoint Online site URL.`);
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
