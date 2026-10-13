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
import command, { options } from './group-get.js';

describe(commands.GROUP_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves group by id with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return {
          "value": [{
            "Id": 7,
            "IsHiddenInUI": false,
            "LoginName": "Team Site Members",
            "Title": "Team Site Members",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "Team Site Members",
            "RequestToJoinLeaveEmailSetting": ""
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 7
      })
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 7,
        IsHiddenInUI: false,
        LoginName: "Team Site Members",
        Title: "Team Site Members",
        PrincipalType: 8,
        AllowMembersEditMembership: false,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: false,
        Description: null,
        OnlyAllowMembersViewMembership: false,
        OwnerTitle: "Team Site Members",
        RequestToJoinLeaveEmailSetting: ""
      }]
    }));
  });

  it('retrieves group by name with output option json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return {
          "value": [{
            "Id": 7,
            "IsHiddenInUI": false,
            "LoginName": "Team Site Members",
            "Title": "Team Site Members",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "Team Site Members",
            "RequestToJoinLeaveEmailSetting": ""
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: "Team Site Members"
      })
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 7,
        IsHiddenInUI: false,
        LoginName: "Team Site Members",
        Title: "Team Site Members",
        PrincipalType: 8,
        AllowMembersEditMembership: false,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: false,
        Description: null,
        OnlyAllowMembersViewMembership: false,
        OwnerTitle: "Team Site Members",
        RequestToJoinLeaveEmailSetting: ""
      }]
    }));
  });

  it('correctly retrieves the associated owner group', async () => {
    const ownerGroupResponse = {
      Id: 3,
      IsHiddenInUI: false,
      LoginName: "Team Site Owners",
      Title: "Team Site Owners",
      PrincipalType: 8
    };

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedOwnerGroup')) {
        return ownerGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroup: 'Owner'
      })
    });
    assert(loggerLogSpy.calledWith(ownerGroupResponse));
  });

  it('correctly retrieves the associated member group', async () => {
    const memberGroupResponse = {
      Id: 3,
      IsHiddenInUI: false,
      LoginName: "Team Site Members",
      Title: "Team Site Members",
      PrincipalType: 8
    };

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedMemberGroup')) {
        return memberGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroup: 'Member'
      })
    });
    assert(loggerLogSpy.calledWith(memberGroupResponse));
  });

  it('correctly retrieves the associated visitor group', async () => {
    const visitorGroupResponse = {
      Id: 3,
      IsHiddenInUI: false,
      LoginName: "Team Site Visitors",
      Title: "Team Site Visitors",
      PrincipalType: 8
    };

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedVisitorGroup')) {
        return visitorGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroup: 'Visitor'
      })
    });
    assert(loggerLogSpy.calledWith(visitorGroupResponse));
  });

  it('handles error correctly', async () => {
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
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroup: 'Visitor'
      })
    }), new CommandError('An error has occurred'));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', id: 1 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if associatedGroup has an invalid value', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', associatedGroup: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both id and name options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id and name both are passed(multiple options)', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 7, name: "Team Site Members" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the specified ID is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 'a' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      id: 7,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation url is valid and id is passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 7 });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the url is valid and name is passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: "Team Site Members" });
    assert.strictEqual(actual.success, true);
  });
}); 
