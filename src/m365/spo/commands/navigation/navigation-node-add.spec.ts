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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { spo } from '../../../../utils/spo';
const command: Command = require('./navigation-node-add');

describe(commands.NAVIGATION_NODE_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';
  const topNavigationResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2039', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2041', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Sub level 1', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2040', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site B', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2001', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/team-a/sitepages/about.aspx', 'Title': 'About', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1025', 'StartingNodeTitle': 'Quick launch', 'Version': '2023-03-09T18:33:53.5468097Z' };
  const quickLaunchResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2003', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2006', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/SharePointDemoSite#/', 'Title': 'Sub Item', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': 'http://google.be', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2018', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Site B', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1002', 'StartingNodeTitle': 'SharePoint Top Navigation Bar', 'Version': '2023-03-09T18:34:53.650545Z' };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.getTopNavigationMenuState,
      spo.getQuickLaunchMenuState,
      spo.saveMenuState
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.NAVIGATION_NODE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('adds new navigation node to the top navigation', async () => {
    const nodeAddResponse = {
      "AudienceIds": null,
      "CurrentLCID": 1033,
      "Id": 2001,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About",
      "Url": "/sites/team-a/sitepages/about.aspx"
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar` &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: title,
          Url: nodeUrl,
          IsExternal: false
        })) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
  });

  it('adds new navigation node to the quick launch navigation and opens it in new window', async () => {
    let saveCalled = false;
    const nodeAddResponse = {
      "AudienceIds": [],
      "CurrentLCID": 1033,
      "Id": 2003,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About",
      "Url": "/sites/team-a/sitepages/about.aspx"
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/quicklaunch`) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getQuickLaunchMenuState').callsFake(async () => {
      return quickLaunchResponse;
    });

    sinon.stub(spo, 'saveMenuState').callsFake(async () => {
      saveCalled = true;
      return;
    });

    await command.action(logger, { options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl, openInNewWindow: true, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
    assert.strictEqual(saveCalled, true);
  });

  it('adds new navigation node with a parent id and opens it in new window', async () => {
    const parentNodeId = 2039;
    let saveCalled = false;
    const nodeAddResponse = {
      "AudienceIds": audienceIds.split(','),
      "CurrentLCID": 1033,
      "Id": 2041,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About",
      "Url": "/sites/team-a/sitepages/about.aspx"
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${parentNodeId})/Children`) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getQuickLaunchMenuState').callsFake(async () => {
      return quickLaunchResponse;
    });

    sinon.stub(spo, 'getTopNavigationMenuState').callsFake(async () => {
      return topNavigationResponse;
    });

    sinon.stub(spo, 'saveMenuState').callsFake(async () => {
      saveCalled = true;
      return;
    });

    await command.action(logger, { options: { webUrl: webUrl, parentNodeId: parentNodeId, title: title, url: nodeUrl, audienceIds: audienceIds, openInNewWindow: true, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
    assert.strictEqual(saveCalled, true);
  });

  it('adds new navigation node below an existing node', async () => {
    const nodeAddResponse = {
      "AudienceIds": null,
      "CurrentLCID": 1033,
      "Id": 2001,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About",
      "Url": "/sites/team-a/sitepages/about.aspx"
    };
    const parentNodeId = 1000;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${parentNodeId})/Children` &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: title,
          Url: nodeUrl,
          IsExternal: false
        })) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, parentNodeId: 1000, title: title, url: nodeUrl, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
  });

  it('adds new navigation node to the top navigation with audience targetting and opens it in new window', async () => {
    let saveCalled = false;
    const nodeAddResponse = {
      "AudienceIds": audienceIds.split(','),
      "CurrentLCID": 1033,
      "Id": 2001,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About",
      "Url": "/sites/team-a/sitepages/about.aspx"
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar`) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getTopNavigationMenuState').callsFake(async () => {
      return topNavigationResponse;
    });

    sinon.stub(spo, 'saveMenuState').callsFake(async () => {
      saveCalled = true;
      return;
    });


    await command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl, audienceIds: audienceIds, openInNewWindow: true, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
    assert.strictEqual(saveCalled, true);
  });

  it('adds new linkless navigation node to the top navigation with', async () => {
    const requestBody = {
      AudienceIds: undefined,
      Title: title,
      Url: 'http://linkless.header/',
      IsExternal: false
    };
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar`) {
        return {
          "AudienceIds": null,
          "CurrentLCID": 1033,
          "Id": 2001,
          "IsDocLib": true,
          "IsExternal": false,
          "IsVisible": true,
          "ListTemplateType": 0,
          "Title": title,
          "Url": "http://linkless.header/"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar`) {
        throw { error: 'An error has occurred' };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar`) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar', title: title, url: nodeUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified parentNodeId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, url: nodeUrl, parentNodeId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, location: 'invalid', title: title, url: nodeUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains an invalid audienceId', async () => {
    const invalidAudienceIds = `${audienceIds},invalid`;
    const actual = await command.validate({ options: { webUrl: webUrl, parentNodeId: 2000, title: title, url: nodeUrl, audienceIds: invalidAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains more than 10 guids', async () => {
    const invalidAudienceIds = `${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds}`;
    const actual = await command.validate({ options: { webUrl: webUrl, parentNodeId: 2000, title: title, url: nodeUrl, audienceIds: invalidAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and the link is external', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl, isExternal: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is QuickLaunch and the link is external', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl, isExternal: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when location is not specified but parentNodeId is', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentNodeId: 2000, title: title, url: nodeUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when audienceIds contains less than 10 ids and all are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentNodeId: 2000, title: title, url: nodeUrl, audienceIds: audienceIds } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
