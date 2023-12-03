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
import command from './group-member-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.GROUP_MEMBER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const JSONSPGroupMembersList =
    [
      {
        "Id": 6,
        "IsHiddenInUI": false,
        "LoginName": "i:0#.f|membership|Alex.Wilber@contoso.com.com",
        "Title": "Alex Wilber",
        "PrincipalType": 1,
        "Email": "Alex.Wilber@contoso.com",
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": true,
        "UserId": {
          "NameId": "10032000afc2e592",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": "Alex.Wilber@contoso.com"
      },
      {
        "Id": 18,
        "IsHiddenInUI": false,
        "LoginName": "i:0#.f|membership|AdeleV@contoso.com",
        "Title": "Adele Vance",
        "PrincipalType": 1,
        "Email": "AdeleV@contoso.com",
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": false,
        "UserId": {
          "NameId": "10032000b07c0d71",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": "AdeleV@contoso.com"
      }
    ];

  const groupMembersList = {
    value: JSONSPGroupMembersList
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
      request.get,
      cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'UserPrincipalName', 'Id', 'Email']);
  });

  it('Getting the members of a SharePoint Group using groupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return groupMembersList;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    });
    assert(loggerLogSpy.calledWith(JSONSPGroupMembersList));
  });

  it('Getting the members of a SharePoint Group using groupId (DEBUG)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return groupMembersList;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    });
    assert(loggerLogSpy.calledWith(JSONSPGroupMembersList));
  });

  it('Getting the members of a SharePoint Group using groupName (DEBUG)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return groupMembersList;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners"
      }
    });
    assert(loggerLogSpy.calledWith(JSONSPGroupMembersList));
  });

  it('Correctly Handles Error when listing members of the group', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupid and groupName is entered', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "4", groupName: "Contoso Site Owners" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither groupId nor groupName is entered', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "INVALIDGROUP" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 3 } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
