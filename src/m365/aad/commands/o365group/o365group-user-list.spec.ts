import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-user-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.O365GROUP_USER_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_USER_LIST), true);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(teamsCommands.TEAMS_USER_LIST) > -1), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the groupId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the groupId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {}
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both groupId and teamId are specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when invalid role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Invalid',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId and no role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId and Owner role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Owner'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId and Member role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Member'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId and Guest role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Guest'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly lists all users in a Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "00000000-0000-0000-0000-000000000000",
            "displayName": "Anne Matthews",
            "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
            "userType": "Owner"
          },
          {
            "id": "00000000-0000-0000-0000-000000000001",
            "displayName": "Karl Matteson",
            "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
            "userType": "Member"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists all owners in a Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", role: "Owner" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "00000000-0000-0000-0000-000000000000",
            "displayName": "Anne Matthews",
            "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
            "userType": "Owner"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists all members in a Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", role: "Member" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "00000000-0000-0000-0000-000000000001",
            "displayName": "Karl Matteson",
            "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
            "userType": "Member"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists all users in a Microsoft 365 group (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "00000000-0000-0000-0000-000000000000",
            "displayName": "Anne Matthews",
            "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
            "userType": "Owner"
          },
          {
            "id": "00000000-0000-0000-0000-000000000001",
            "displayName": "Karl Matteson",
            "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
            "userType": "Member"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing users', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000" } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});