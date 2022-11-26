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
const command: Command = require('./group-get');

describe(commands.GROUP_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name', 'associatedGroup'] }]);
  });

  it('retrieves group by id with output option json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetById') > -1) {
        return Promise.resolve(
          {
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
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 7
      }
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups/GetByName') > -1) {
        return Promise.resolve(
          {
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
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: "Team Site Members"
      }
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

    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedOwnerGroup')) {
        return Promise.resolve(ownerGroupResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        associatedGroup: 'Owner'
      }
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

    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedMemberGroup')) {
        return Promise.resolve(memberGroupResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        associatedGroup: 'Member'
      }
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

    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url!.endsWith('/_api/web/AssociatedVisitorGroup')) {
        return Promise.resolve(visitorGroupResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        associatedGroup: 'Visitor'
      }
    });
    assert(loggerLogSpy.calledWith(visitorGroupResponse));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        associatedGroup: 'Visitor'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if associatedGroup has an invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', associatedGroup: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and name both are passed(multiple options)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 7, name: "Team Site Members" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified ID is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 7 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and name is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: "Team Site Members" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 