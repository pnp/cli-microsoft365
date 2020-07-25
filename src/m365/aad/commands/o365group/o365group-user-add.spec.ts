import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-user-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.O365GROUP_USER_ADD, () => {
  let log: string[];
  let cmdInstance: any;

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
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_USER_ADD), true);
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
    assert.strictEqual((alias && alias.indexOf(teamsCommands.TEAMS_USER_ADD) > -1), true);
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

  it('fails validation if neither the groupId nor teamId are provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        role: 'Member'
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

  it('fails validation when invalid role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Invalid',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and no role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and Owner role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Owner'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and Member role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Member'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly retrieves user and add member to specified Microsoft 365 group', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }

      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add member to specified Microsoft 365 group (debug)', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }

      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add owner to specified Microsoft 365 group', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }

      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", role: "Owner" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add owner to specified Microsoft Teams team (debug)', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }

      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", role: "Owner" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly skips adding member or owner when user is not found', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews.not.found%40contoso.onmicrosoft.com/id`) {
        return Promise.reject({
          "message": "Resource 'anne.matthews.not.found%40contoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present."
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref`) {
        addMemberRequestIssued = true;
      }

      return Promise.resolve();
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews.not.found@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued === false);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly handles error when user cannot be retrieved', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/doesnotexist.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'Resource \'doesnotexist.matthews@contoso.onmicrosoft.com\' does not exist or one of its queried reference-property objects are not present.' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "doesnotexist.matthews@contoso.onmicrosoft.com" } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Resource \'doesnotexist.matthews@contoso.onmicrosoft.com\' does not exist or one of its queried reference-property objects are not present.')));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and handle error adding member to specified Microsoft 365 group', (done) => {

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref`) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'Invalid object identifier' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Invalid object identifier'))); done();
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