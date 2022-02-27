import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-get');

describe(commands.GROUP_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves group by id with output option json', (done) => {
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

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 7
      }
    }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves group by name with output option json', (done) => {
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

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: "Team Site Members"
      }
    }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are not passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and name both are passed(multiple options)', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 7, name: "Team Site Members" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified ID is not a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'a' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 7 } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and name is passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: "Team Site Members" } });
    assert.strictEqual(actual, true);
  });
}); 