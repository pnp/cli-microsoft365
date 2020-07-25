import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-user-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.O365GROUP_USER_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.delete,
      global.setTimeout
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_USER_REMOVE), true);
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
    assert.strictEqual((alias && alias.indexOf(teamsCommands.TEAMS_USER_REMOVE) > -1), true);
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

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the groupId nor the teamID are provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
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

  it('passes validation when valid groupId and userName are specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('prompts before removing the specified user from the specified Microsoft 365 Group when confirm option not passed', (done) => {
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified user from the specified Team when confirm option not passed (debug)', (done) => {
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified user from the specified Microsoft 365 Group when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified owner from owners endpoint of the specified Microsoft 365 Group when prompt confirmed', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });


    sinon.stub(request, 'delete').callsFake((opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        });
      }

      return Promise.reject('Invalid request');

    });

    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified user from the members specified Microsoft 365 Group when prompt confirmed (verbose)', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/karl.matteson%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000002"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000002/$ref`) {
        memberDeleteCallIssued = true;

        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "karl.matteson@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly retrieves user and handle error removing member from specified Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'Invalid object identifier' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Invalid object identifier'))); done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly skips execution when specified user is not found', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews.not.found%40contoso.onmicrosoft.com/id`) {
        return Promise.reject({
          "message": "Resource 'anne.matthews.not.found%40contoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present."
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      memberDeleteCallIssued = true;
      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(memberDeleteCallIssued === false);
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