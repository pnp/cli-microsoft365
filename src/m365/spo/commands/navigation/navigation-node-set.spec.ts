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
import command, { options } from './navigation-node-set.js';

describe(commands.NAVIGATION_NODE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = 2006;
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';
  const node = {
    AudienceIds: ['0d718612-8407-4d6b-833c-6891a553354f'],
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
        Nodes: [node],
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
    ...quickLaunchMenuState,
    Nodes: [node],
    StartingNodeKey: '1002',
    StartingNodeTitle: 'SharePoint Top Navigation Bar'
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const stubMenuStateRequests = (nodeInTopNavigation: boolean = false): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        if (opts.data.menuNodeKey === null) {
          return nodeInTopNavigation ? { ...quickLaunchMenuState, Nodes: [] } : structuredClone(quickLaunchMenuState);
        }

        if (opts.data.menuNodeKey === '1002') {
          return nodeInTopNavigation ? structuredClone(topNavigationMenuState) : { ...topNavigationMenuState, Nodes: [] };
        }
      }

      if (opts.url === `${webUrl}/_api/navigation/SaveMenuState`) {
        return { value: 200 };
      }

      throw 'Invalid request';
    });
  };

  const getSavedMenuState = (postStub: sinon.SinonStub): any => {
    return postStub.getCalls().find(call => call.args[0].url === `${webUrl}/_api/navigation/SaveMenuState`)!.args[0].data.menuState;
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
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('updates all properties of a navigation node in the quick launch', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, audienceIds: audienceIds, openInNewWindow: true, verbose: true } });
    assert.deepStrictEqual(getSavedMenuState(postStub).Nodes[0].Nodes[0], {
      ...node,
      AudienceIds: audienceIds.split(','),
      OpenInNewWindow: true,
      SimpleUrl: nodeUrl,
      Title: title
    });
  });

  it('updates the title of a navigation node in the top navigation', async () => {
    const postStub = stubMenuStateRequests(true);

    await command.action(logger, { options: { webUrl: webUrl, id: id, title: title } });
    const savedMenuState = getSavedMenuState(postStub);
    assert.strictEqual(savedMenuState.StartingNodeKey, '1002');
    assert.deepStrictEqual(savedMenuState.Nodes[0], { ...node, Title: title });
  });

  it('sets a navigation node not to open in a new window', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, id: id, openInNewWindow: false } });
    assert.deepStrictEqual(getSavedMenuState(postStub).Nodes[0].Nodes[0], { ...node, OpenInNewWindow: false });
  });

  it('clears audienceIds of a navigation node', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, id: id, audienceIds: '' } });
    assert.deepStrictEqual(getSavedMenuState(postStub).Nodes[0].Nodes[0], { ...node, AudienceIds: [] });
  });

  it('sets a navigation node as linkless', async () => {
    const postStub = stubMenuStateRequests();

    await command.action(logger, { options: { webUrl: webUrl, id: id, url: '' } });
    assert.deepStrictEqual(getSavedMenuState(postStub).Nodes[0].Nodes[0], { ...node, SimpleUrl: '' });
  });

  it('throws an error when the navigation node does not exist', async () => {
    stubMenuStateRequests();

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: 9999, title: title } }),
      new CommandError(`Navigation node with id '9999' not found.`));
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

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id, title: title } }),
      new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', id: id, title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id is not a valid number', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: 'invalid', title: title });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no options are set to be changed', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if audienceIds contains more than 10 guids', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, audienceIds: Array(11).fill('7aa4a1ca-4035-4f2f-bac7-7beada59b5ba').join(',') });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if audienceIds contains invalid guid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, audienceIds: `${audienceIds},invalid` });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if the removed isExternal option is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, isExternal: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if all options are set properly', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, title: title, url: nodeUrl, audienceIds: audienceIds, openInNewWindow: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if audienceIds is an empty string', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, audienceIds: '' });
    assert.strictEqual(actual.success, true);
  });
});
