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
import command from './navigation-node-get.js';

describe(commands.NAVIGATION_NODE_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = '2209';
  const navigationNodeGetResponse = {
    "AudienceIds": null,
    "CurrentLCID": 1033,
    "Id": id,
    "IsDocLib": true,
    "IsExternal": false,
    "IsVisible": true,
    "ListTemplateType": 100,
    "Title": "Work Status",
    "Url": "/sites/team-a/Lists/Work Status/AllItems.aspx"
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid number', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl and id are specified', async () => {
    const actual = await command.validate({
      options: {
        webUrl: webUrl,
        id: id
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves navigation node by specified webUrl and id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        return navigationNodeGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: id, verbose: true } });
    assert(loggerLogSpy.calledWith(navigationNodeGetResponse));
  });

  it('command correctly handles navigation node get reject request', async () => {
    const errorMessage = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        id: id
      }
    }), new CommandError(errorMessage));
  });
});