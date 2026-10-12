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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command, { options } from './navigation-node-set.js';

describe(commands.NAVIGATION_NODE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = 2000;
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';

  const topNavigationResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2039', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2041', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Sub level 1', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2040', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site B', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2001', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/team-a/sitepages/about.aspx', 'Title': 'About', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1025', 'StartingNodeTitle': 'Quick launch', 'Version': '2023-03-09T18:33:53.5468097Z' };
  const quickLaunchResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2003', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2006', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/SharePointDemoSite#/', 'Title': 'Sub Item', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': 'http://google.be', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2018', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Site B', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1002', 'StartingNodeTitle': 'SharePoint Top Navigation Bar', 'Version': '2023-03-09T18:34:53.650545Z' };


  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
      request.patch,
      spo.getQuickLaunchMenuState,
      spo.getTopNavigationMenuState,
      spo.saveMenuState]);
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

  it('correctly updates existing navigation node and make sure it opens link in new tab', async () => {
    let saveCalled = false;
    const id = 2041;
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return '';
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


    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: false, audienceIds: audienceIds, openInNewWindow: true, verbose: true }) });
    assert.strictEqual(saveCalled, true);
  });

  it('correctly clears audienceIds from existing navigation node', async () => {
    const requestBody = {
      AudienceIds: [],
      IsExternal: undefined,
      Title: undefined,
      Url: undefined
    };
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return '';
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, audienceIds: "" }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly sets navigation node as linkless', async () => {
    const requestBody = {
      IsExternal: undefined,
      Title: undefined,
      Url: 'http://linkless.header/'
    };
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return '';
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, url: "" }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles navigation node that does not exist', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, id: id, title: title, verbose: true }) }), new CommandError('Navigation node does not exist.'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if no options are set to be changed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if audienceIds contains more than 10 guids', () => {
    const manyAudienceIds = `${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds}`;
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, audienceIds: manyAudienceIds });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if audienceIds contains invalid guid', () => {
    const invalidAudienceIds = `${audienceIds},invalid`;
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, audienceIds: invalidAudienceIds });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if all options are set properly', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: true, audienceIds: audienceIds });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: id,
      title: title,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('returns a validation error when audienceIds contains more than 10 GUIDs', () => {
    const manyAudienceIds = `${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds}`;
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: id,
      audienceIds: manyAudienceIds
    });

    assert.strictEqual(actual.success, false);
    if (!actual.success) {
      assert.strictEqual(
        actual.error.issues[0].message,
        'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.'
      );
    }
  });

  it('returns a validation error when audienceIds contains an invalid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: id,
      audienceIds: `${audienceIds},invalid`
    });

    assert.strictEqual(actual.success, false);
    if (!actual.success) {
      assert.strictEqual(
        actual.error.issues[0].message,
        'The following GUIDs are invalid for the option \'audienceIds\': invalid.'
      );
    }
  });
});
