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
import command, { options } from './navigation-node-get.js';

describe(commands.NAVIGATION_NODE_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = 2209;
  const childNode = {
    AudienceIds: [],
    CurrentLCID: 1033,
    CustomProperties: [],
    FriendlyUrlSegment: '',
    IsDeleted: false,
    IsHidden: false,
    IsTitleForExistingLanguage: false,
    Key: '2209',
    Nodes: [
      {
        AudienceIds: [],
        CurrentLCID: 1033,
        CustomProperties: [],
        FriendlyUrlSegment: '',
        IsDeleted: false,
        IsHidden: false,
        IsTitleForExistingLanguage: false,
        Key: '2210',
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
    SimpleUrl: '/sites/team-a/Lists/Work Status/AllItems.aspx',
    Title: 'Work Status',
    Translations: []
  };
  const quickLaunchMenuState = {
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
        Nodes: [childNode],
        NodeType: 0,
        OpenInNewWindow: null,
        SimpleUrl: '',
        Title: 'Lists',
        Translations: []
      }
    ],
    SimpleUrl: '',
    SPSitePrefix: '/sites/team-a',
    SPWebPrefix: '/sites/team-a',
    StartingNodeKey: '1025',
    StartingNodeTitle: 'Quick launch',
    Version: '2026-10-07T19:21:10.213646Z'
  };
  const topNavigationMenuState = {
    AudienceIds: [],
    FriendlyUrlPrefix: '',
    IsAudienceTargetEnabledForGlobalNav: false,
    Nodes: [childNode],
    SimpleUrl: '',
    SPSitePrefix: '/sites/team-a',
    SPWebPrefix: '/sites/team-a',
    StartingNodeKey: '1002',
    StartingNodeTitle: 'SharePoint Top Navigation Bar',
    Version: '2026-10-07T19:21:10.213646Z'
  };
  const emptyQuickLaunchMenuState = { ...quickLaunchMenuState, Nodes: [] };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', id: id });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id is not a valid number', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: 12.48 });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when webUrl and id are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves navigation node from the quick launch', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState` && opts.data.menuNodeKey === null) {
        return quickLaunchMenuState;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: id, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(childNode));
  });

  it('retrieves navigation node from the top navigation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        if (opts.data.menuNodeKey === null) {
          return emptyQuickLaunchMenuState;
        }

        if (opts.data.menuNodeKey === '1002') {
          return topNavigationMenuState;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: id } });
    assert(loggerLogSpy.calledOnceWith(childNode));
  });

  it('throws an error when navigation node is not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return emptyQuickLaunchMenuState;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id } }),
      new CommandError(`Navigation node with id '${id}' not found.`));
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        code: "-2147024891, System.UnauthorizedAccessException",
        message: "Attempted to perform an unauthorized operation."
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id } }),
      new CommandError("Attempted to perform an unauthorized operation."));
  });
});
