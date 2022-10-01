import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./roledefinition-get');

describe(commands.ROLEDEFINITION_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ROLEDEFINITION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
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
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions(1)') > -1) {
        return Promise.resolve(
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
            "RoleTypeKind": 4
          });
      }
      return Promise.reject('Invalid request');
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