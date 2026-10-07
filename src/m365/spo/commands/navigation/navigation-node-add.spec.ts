import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './navigation-node-add.js';

describe(commands.NAVIGATION_NODE_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';
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
        Nodes: [
          {
            AudienceIds: [],
            CurrentLCID: 1033,
            CustomProperties: [],
            FriendlyUrlSegment: '',
            IsDeleted: false,
            IsHidden: false,
            IsTitleForExistingLanguage: false,
            Key: '2006',
            Nodes: [],
            NodeType: 0,
            OpenInNewWindow: null,
            SimpleUrl: '/sites/team-a/SitePages/Home.aspx',
            Title: 'Sub Item',
            Translations: []
          }
        ],
        NodeType: 0,
        OpenInNewWindow: true,
        SimpleUrl: 'https://contoso.com',
        Title: 'Site A',
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
    Nodes: [
      {
        AudienceIds: [],
        CurrentLCID: 1033,
        CustomProperties: [],
        FriendlyUrlSegment: '',
        IsDeleted: false,
        IsHidden: false,
        IsTitleForExistingLanguage: false,
        Key: '2039',
        Nodes: [],
        NodeType: 0,
        OpenInNewWindow: null,
        SimpleUrl: '/sites/team-a',
        Title: 'Home',
        Translations: []
      }
    ],
    SimpleUrl: '',
    SPSitePrefix: '/sites/team-a',
    SPWebPrefix: '/sites/team-a',
    StartingNodeKey: '1002',
    StartingNodeTitle: 'SharePoint Top Navigation Bar',
    Version: '2026-10-07T19:21:10.213646Z'
  };
  const addedNode = {
    AudienceIds: [],
    CurrentLCID: 1033,
    CustomProperties: [],
    FriendlyUrlSegment: '',
    IsDeleted: false,
    IsHidden: false,
    IsTitleForExistingLanguage: false,
    Key: '3000',
    Nodes: [],
    NodeType: 0,
    OpenInNewWindow: null,
    SimpleUrl: nodeUrl,
    Title: title,
    Translations: []
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const stubMenuStateRequests = (persistChanges: boolean = true): sinon.SinonStub => {
    const menuStates: { [startingNodeKey: string]: any } = {
      '1025': structuredClone(quickLaunchMenuState),
      '1002': structuredClone(topNavigationMenuState)
    };
    let nextKey = 3000;
    const assignKeys = (nodes: any[]): void => {
      nodes.forEach(node => {
        if (node.Key === null) {
          Object.assign(node, {
            CurrentLCID: 1033,
            CustomProperties: [],
            FriendlyUrlSegment: '',
            IsTitleForExistingLanguage: false,
            Key: (nextKey++).toString(),
            Translations: []
          });
        }
        assignKeys(node.Nodes);
      });
    };

    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return structuredClone(menuStates[opts.data.menuNodeKey ?? '1025']);
      }

      if (opts.url === `${webUrl}/_api/navigation/SaveMenuState`) {
        if (persistChanges) {
          const menuState = structuredClone(opts.data.menuState);
          assignKeys(menuState.Nodes);
          menuStates[menuState.StartingNodeKey] = menuState;
        }
        return { value: 200 };
      }

      throw 'Invalid request';
    });
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
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('adds new navigation node to the top navigation', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl, verbose: true } });
    const saveRequest = postStub.getCalls().find(call => call.args[0].url === `${webUrl}/_api/navigation/SaveMenuState`)!;
    assert.deepStrictEqual(saveRequest.args[0].data.menuState.Nodes.at(-1), {
      AudienceIds: [],
      IsDeleted: false,
      IsHidden: false,
      Key: null,
      Nodes: [],
      NodeType: 0,
      OpenInNewWindow: null,
      SimpleUrl: nodeUrl,
      Title: title
    });
    assert.strictEqual(saveRequest.args[0].data.menuState.StartingNodeKey, '1002');
    assert(loggerLogSpy.calledOnceWith(addedNode));
  });

  it('adds new navigation node to the quick launch with audience targeting and opens it in new window', async () => {
    stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl, audienceIds: audienceIds, openInNewWindow: true } });
    assert(loggerLogSpy.calledOnceWith({ ...addedNode, AudienceIds: audienceIds.split(','), OpenInNewWindow: true }));
  });

  it('adds new linkless navigation node to the quick launch', async () => {
    stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, location: 'QuickLaunch', title: title } });
    assert(loggerLogSpy.calledOnceWith({ ...addedNode, SimpleUrl: '' }));
  });

  it('adds new navigation node below an existing node in the quick launch', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, parentNodeId: 2006, title: title, url: nodeUrl } });
    const saveRequest = postStub.getCalls().find(call => call.args[0].url === `${webUrl}/_api/navigation/SaveMenuState`)!;
    assert.strictEqual(saveRequest.args[0].data.menuState.Nodes[0].Nodes[0].Nodes[0].Title, title);
    assert(loggerLogSpy.calledOnceWith(addedNode));
  });

  it('adds new navigation node below an existing node in the top navigation', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, parentNodeId: 2039, title: title, url: nodeUrl } });
    const saveRequest = postStub.getCalls().find(call => call.args[0].url === `${webUrl}/_api/navigation/SaveMenuState`)!;
    assert.strictEqual(saveRequest.args[0].data.menuState.StartingNodeKey, '1002');
    assert.strictEqual(saveRequest.args[0].data.menuState.Nodes[0].Nodes[0].Title, title);
    assert(loggerLogSpy.calledOnceWith(addedNode));
  });

  it('throws an error when the parent node does not exist', async () => {
    stubMenuStateRequests();

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, parentNodeId: 9999, title: title, url: nodeUrl } }),
      new CommandError(`Navigation node with id '9999' not found.`));
  });

  it('throws an error when the added node cannot be retrieved', async () => {
    stubMenuStateRequests(false);

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl } }),
      new CommandError(`The navigation node was added, but it couldn't be retrieved.`));
  });

  it('throws an error when the parent node cannot be retrieved after adding the node', async () => {
    let menuStateRequests = 0;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return ++menuStateRequests === 1 ? structuredClone(quickLaunchMenuState) : { ...quickLaunchMenuState, Nodes: [] };
      }

      if (opts.url === `${webUrl}/_api/navigation/SaveMenuState`) {
        return { value: 200 };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, parentNodeId: 2003, title: title, url: nodeUrl } }),
      new CommandError(`The navigation node was added, but it couldn't be retrieved.`));
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl } }),
      new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', location: 'TopNavigationBar', title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the specified parentNodeId is not a number', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, parentNodeId: 'invalid', title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'invalid', title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both location and parentNodeId are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'QuickLaunch', parentNodeId: 2003, title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither location nor parentNodeId is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if title is not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'QuickLaunch' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if audienceIds contains an invalid audienceId', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'TopNavigationBar', title: title, audienceIds: `${audienceIds},invalid` });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if audienceIds contains more than 10 guids', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'TopNavigationBar', title: title, audienceIds: Array(11).fill('7aa4a1ca-4035-4f2f-bac7-7beada59b5ba').join(',') });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the removed isExternal option is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'QuickLaunch', title: title, isExternal: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'TopNavigationBar', title: title });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when location is QuickLaunch and all properties are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, location: 'QuickLaunch', title: title, url: 'https://contoso.com', audienceIds: audienceIds, openInNewWindow: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when location is not specified but parentNodeId is', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, parentNodeId: 2003, title: title });
    assert.strictEqual(actual.success, true);
  });
});
