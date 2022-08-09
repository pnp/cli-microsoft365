import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import teamsCommands from '../../../teams/commands';
import commands from '../../commands';
const command: Command = require('./o365group-user-set');

describe(commands.O365GROUP_USER_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.delete
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_USER_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(teamsCommands.USER_SET) > -1), true);
  });

  it('fails validation if the groupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        role: 'Member',
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        role: 'Member',
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the groupId nor the teamId are provided.', async () => {
    const actual = await command.validate({
      options: {
        role: 'Member',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and teamId are specified', async () => {
    const actual = await command.validate({
      options: {
        role: 'Member',
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid role is specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid teamId, userName and role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Member'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows error when the specified user is not present in specified O365 Group', (done) => {
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

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: 'notpresent.karl.matteson@contoso.onmicrosoft.com', role: 'Member' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The specified user does not belong to the given Microsoft 365 Group. Please use the 'o365group user add' command to add new users.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified user is not present in specified Microsoft Teams team', (done) => {
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

    command.action(logger, { options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: 'notpresent.karl.matteson@contoso.onmicrosoft.com', role: 'Member' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The specified user does not belong to the given Microsoft Teams team. Please use the 'graph teams user add' command to add new users.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified user is already a member in specified O365 Group', (done) => {
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

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: 'karl.matteson@contoso.onmicrosoft.com', role: 'Member' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The specified user is already a member in the specified Microsoft 365 group, and thus cannot be demoted.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified user is already a member in specified Microsoft Teams team', (done) => {
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

    command.action(logger, { options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: 'karl.matteson@contoso.onmicrosoft.com', role: 'Member' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The specified user is already a member in the specified Microsoft Teams team, and thus cannot be demoted.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified user is already a owner in specified Microsoft 365 Group', (done) => {
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

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: 'anne.matthews@contoso.onmicrosoft.com', role: 'Owner' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The specified user is already an owner in the specified Microsoft 365 group, and thus cannot be promoted.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified user is already a owner in specified Microsoft Teams team', (done) => {
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

    command.action(logger, { options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: 'anne.matthews@contoso.onmicrosoft.com', role: 'Owner' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The specified user is already an owner in the specified Microsoft Teams team, and thus cannot be promoted.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly promotes specified member to owner in specified Microsoft 365 Group', (done) => {
    let promoteMemberIssued = false;

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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.data) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        promoteMemberIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: 'karl.matteson@contoso.onmicrosoft.com', role: 'Owner' } } as any, () => {
      try {
        assert(promoteMemberIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly promotes specified member to owner in specified Microsoft Teams team (debug)', (done) => {
    let promoteMemberIssued = false;

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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.data) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        promoteMemberIssued = true;
        return Promise.resolve();
      }
    
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: 'karl.matteson@contoso.onmicrosoft.com', role: 'Owner' } } as any, () => {
      try {
        assert(promoteMemberIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly demote specified owner to member in specified Microsoft 365 Group', (done) => {
    let demoteOwnerIssued = false;

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

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000000/$ref`) {
        demoteOwnerIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: 'anne.matthews@contoso.onmicrosoft.com', role: 'Member' } } as any, () => {
      try {
        assert(demoteOwnerIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly demote specified owner to member in specified Microsoft 365 Group (debug)', (done) => {
    let demoteOwnerIssued = false;

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

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000000/$ref`) {
        demoteOwnerIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');      
    });

    command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: 'anne.matthews@contoso.onmicrosoft.com', role: 'Member' } } as any, () => {
      try {
        assert(demoteOwnerIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});