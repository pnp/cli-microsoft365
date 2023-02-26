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
const command: Command = require('./group-member-list');

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
      request.get,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_MEMBER_LIST), true);
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        throw 'Invalid request';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 3
      }
    }), new CommandError('Invalid request'));
  });

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupid and groupName is entered', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "4", groupName: "Contoso Site Owners" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither groupId nor groupName is entered', async () => {
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
