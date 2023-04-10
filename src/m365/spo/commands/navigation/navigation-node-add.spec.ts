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
const command: Command = require('./navigation-node-add');

describe(commands.NAVIGATION_NODE_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';

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
      request.post
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

  it('adds new navigation node pointing to an external URL to the quick launch', async () => {
    const nodeAddResponse = {
      "AudienceIds": null,
      "CurrentLCID": 1033,
      "Id": 2001,
      "IsDocLib": true,
      "IsExternal": true,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "About us",
      "Url": "https://contoso.com/about-us"
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/quicklaunch` &&
        JSON.stringify(opts.data) === JSON.stringify({
          Title: title,
          Url: nodeUrl,
          IsExternal: true
        })) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, location: 'QuickLaunch', title: title, url: nodeUrl, isExternal: true, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
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

  it('adds new navigation node to the top navigation with audience targetting', async () => {
    const nodeAddResponse = {
      "AudienceIds": audienceIds.split(','),
      "CurrentLCID": 1033,
      "Id": 2001,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": title,
      "Url": nodeUrl
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/topnavigationbar`) {
        return nodeAddResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, location: 'TopNavigationBar', title: title, url: nodeUrl, audienceIds: audienceIds, verbose: true } });
    assert(loggerLogSpy.calledWith(nodeAddResponse));
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
