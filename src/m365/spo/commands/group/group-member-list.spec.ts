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
import command, { options } from './group-member-list.js';

describe(commands.GROUP_MEMBER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
      options: commandOptionsSchema.parse({
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      })
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
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      })
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
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners"
      })
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
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      })
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if webURL is Invalid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: "InvalidWEBURL", groupId: 3 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'InvalidWEBURL' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if groupid and groupName is entered', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 4, groupName: "Contoso Site Owners" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither groupId nor groupName is entered', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: "https://contoso.sharepoint.com/sites/SiteA" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if groupId is Invalid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "INVALIDGROUP" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com/sites/SiteA",
      groupId: 3,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if all the required options are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 3 });
    assert.strictEqual(actual.success, true);
  });
});
