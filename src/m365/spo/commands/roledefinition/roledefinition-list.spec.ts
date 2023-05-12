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
const command: Command = require('./roledefinition-list');

describe(commands.ROLEDEFINITION_LIST, () => {
  let log: any[];
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ROLEDEFINITION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Name']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('list role definitions handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
  });

  it('lists all role definitions from web', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/cli/_api/web/roledefinitions') {
        return ({
          value:
            [
              {
                "BasePermissions": {
                  "High": "2147483647",
                  "Low": "4294967295"
                },
                "Description": "Has full control.",
                "Hidden": false,
                "Id": 1073741829,
                "Name": "Full Control",
                "Order": 1,
                "RoleTypeKind": 5
              },
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
              }
            ]
        });
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/cli'
      }
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "BasePermissions": {
            "High": "2147483647",
            "Low": "4294967295"
          },
          "Description": "Has full control.",
          "Hidden": false,
          "Id": 1073741829,
          "Name": "Full Control",
          "Order": 1,
          "RoleTypeKind": 5,
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
            "AnonymousSearchAccessList",
            "Open",
            "ViewPages",
            "AddAndCustomizePages",
            "ApplyThemeAndBorder",
            "ApplyStyleSheets",
            "ViewUsageData",
            "CreateSSCSite",
            "ManageSubwebs",
            "CreateGroups",
            "ManagePermissions",
            "BrowseDirectories",
            "BrowseUserInfo",
            "AddDelPrivateWebParts",
            "UpdatePersonalWebParts",
            "ManageWeb",
            "AnonymousSearchAccessWebLists",
            "UseClientIntegration",
            "UseRemoteAPIs",
            "ManageAlerts",
            "CreateAlerts",
            "EditMyUserInfo",
            "EnumeratePermissions"
          ],
          "RoleTypeKindValue": "Administrator"
        },
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
      ]
    ));
  });

  it('should return an empty array for BasePermissionValue & not return RoleTypeKindValue with unmappable data', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return ({
          value:
            [
              {
                "BasePermissions": {
                  "High": "0",
                  "Low": "0"
                },
                "Description": "Has no permissions.",
                "Hidden": false,
                "Id": 1073741822,
                "Name": "No Permissions",
                "Order": 1,
                "RoleTypeKind": 9
              }
            ]
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "BasePermissions": {
            "High": "0",
            "Low": "0"
          },
          "Description": "Has no permissions.",
          "Hidden": false,
          "Id": 1073741822,
          "Name": "No Permissions",
          "Order": 1,
          "RoleTypeKind": 9,
          "BasePermissionsValue": [],
          "RoleTypeKindValue": undefined
        }
      ]
    ));
  });
});
