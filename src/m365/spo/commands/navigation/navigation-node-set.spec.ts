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
import { MenuState } from './NavigationNode';
const command: Command = require('./navigation-node-set');

describe(commands.NAVIGATION_NODE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = 2000;
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';
  const menuState: MenuState = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2560', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Lists/TestLiiiist/AllItems.aspx', 'Title': 'TestLiiiist', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2565', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Lists/TTTT/AllItems.aspx', 'Title': 'TTTT', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2587', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2588', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2589', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub 2', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub1', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2590', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2591', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub 1', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2592', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub2', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2593', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub 1', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2594', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub 1', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Sub1', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'NavLink', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '1033', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2572', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Lists/PNP_SDK_TEST_HandleMaxRequestsInCsomBatch/AllItems.aspx', 'Title': 'PNP_SDK_TEST_HandleMaxRequestsInCsomBatch', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2563', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Lists/aaaaaa/AllItems.aspx', 'Title': 'aaaaaa', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2527', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Teams Wiki Data/Forms/AllItems.aspx', 'Title': 'Teams Wiki Data', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2526', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/pnpcoresdktestgroup/Lists/19gTfvczSQL0KAVOEGY3jvpEN8wm7aBF_kmmftUcAjbk1threa/AllItems.aspx', 'Title': '19:gTfvczSQL0KAVOEGY3jvpEN8wm7-aBF_kmmftUcAjbk1@thread.tacv2_wiki', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '', 'Title': 'Recent', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 0, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '-1', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://mathijsdev2.sharepoint.com/sites/pnpcoresdktestgroup/_layouts/15/AdminRecycleBin.aspx?ql=1', 'Title': 'Recycle Bin', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/pnpcoresdktestgroup', 'SPWebPrefix': '/sites/pnpcoresdktestgroup', 'StartingNodeKey': '1025', 'StartingNodeTitle': 'Quick launch', 'Version': '2023-03-01T21:20:41.3422485Z' };

  let log: string[];
  let logger: Logger;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      spo.getMenuState
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly updates existing navigation node', async () => {
    const requestBody = {
      Title: title,
      Url: nodeUrl,
      IsExternal: false,
      AudienceIds: audienceIds.split(',')
    };
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getMenuState').callsFake(async (webUrl: string) => {
      if (webUrl) { }
      return menuState;
    });

    await command.action(logger, { options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: false, audienceIds: audienceIds } } as any);
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
  });

  it('correctly updates existing navigation node and make sure it opens link in new tab', async () => {
    const id = 2588;
    const requestBody = {
      Title: title,
      Url: nodeUrl,
      IsExternal: false,
      AudienceIds: audienceIds.split(',')
    };
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return '';
      }

      throw 'Invalid request';
    });
    sinon.stub(spo, 'getMenuState').callsFake(async () => { return menuState; });
    await command.action(logger, { options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: false, audienceIds: audienceIds, openInNewWindow: true } } as any);
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, requestBody);
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
    await command.action(logger, { options: { webUrl: webUrl, id: id, audienceIds: "" } } as any);
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
    await command.action(logger, { options: { webUrl: webUrl, id: id, url: "" } } as any);
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

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id, title: title, verbose: true } } as any), new CommandError('Navigation node does not exist.'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is no options are set to be changed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains more than 10 guids', async () => {
    const manyAudienceIds = `${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds}`;
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, audienceIds: manyAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains invalid guid', async () => {
    const invalidAudienceIds = `${audienceIds},invalid`;
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, audienceIds: invalidAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are set properly', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: true, audienceIds: audienceIds } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
