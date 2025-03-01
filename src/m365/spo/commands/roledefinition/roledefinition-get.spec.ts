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
import command from './roledefinition-get.js';

describe(commands.ROLEDEFINITION_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    assert.strictEqual(command.name, commands.ROLEDEFINITION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'aaa' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions(1)') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    }), new CommandError(err));
  });

  it('gets role definition from web by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions(1)') > -1) {
        return {
          "BasePermissions": {
            "High": "432",
            "Low": "1012866047"
          },
          "Description": "Can view, add, update, delete, approve, and customize.",
          "Hidden": false,
          "Id": 1073741828,
          "Name": "Design",
          "Order": 32,
          "RoleTypeKind": 4
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    });
    assert(loggerLogSpy.calledWith(
      {
        "BasePermissions": {
          "High": "432",
          "Low": "1012866047"
        },
        "Description": "Can view, add, update, delete, approve, and customize.",
        "Hidden": false,
        "Id": 1073741828,
        "Name": "Design",
        "Order": 32,
        "RoleTypeKind": 4,
        "BasePermissionsValue": [
          "ViewListItems",
          "AddListItems",
          "EditListItems",
          "DeleteListItems",
          "ApproveItems",
          "OpenItems",
          "ViewVersions",
          "DeleteVersions",
          "CancelCheckout",
          "ManagePersonalViews",
          "ManageLists",
          "ViewFormPages",
          "Open",
          "ViewPages",
          "AddAndCustomizePages",
          "ApplyThemeAndBorder",
          "ApplyStyleSheets",
          "CreateSSCSite",
          "BrowseDirectories",
          "BrowseUserInfo",
          "AddDelPrivateWebParts",
          "UpdatePersonalWebParts",
          "UseClientIntegration",
          "UseRemoteAPIs",
          "CreateAlerts",
          "EditMyUserInfo"
        ],
        "RoleTypeKindValue": "WebDesigner"
      }
    ));
  });
});
