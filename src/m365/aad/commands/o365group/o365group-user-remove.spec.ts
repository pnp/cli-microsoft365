import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import teamsCommands from '../../../teams/commands';
import commands from '../../commands';
const command: Command = require('./o365group-user-remove');

describe(commands.O365GROUP_USER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      global.setTimeout,
      Cli.prompt
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
    assert.strictEqual((alias && alias.indexOf(teamsCommands.USER_REMOVE) > -1), true);
  });

  it('fails validation if the groupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the groupId nor the teamID are provided.', (done) => {
    const actual = command.validate({
      options: {
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both groupId and teamId are specified', (done) => {
    const actual = command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId and userName are specified', (done) => {
    const actual = command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('prompts before removing the specified user from the specified Microsoft 365 Group when confirm option not passed', (done) => {
    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
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
    command.action(logger, { options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
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
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
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
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified owner from owners and members endpoint of the specified Microsoft 365 Group when prompt confirmed', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        });
      }

      return Promise.reject('Invalid request');

    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified member from members endpoint of the specified Microsoft 365 Group when prompt confirmed', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });


    sinon.stub(request, 'delete').callsFake((opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "value": [{ "id": "00000000-0000-0000-0000-000000000000" }]
        });
      }

      return Promise.reject('Invalid request');

    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified owners from owners endpoint of the specified Microsoft 365 Group when prompt confirmed', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      return Promise.reject('Invalid request');

    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('does not fail if the user is not owner or member of the specified Microsoft 365 Group when prompt confirmed', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });


    sinon.stub(request, 'delete').callsFake((opts) => {
      memberDeleteCallIssued = true;

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      return Promise.reject('Invalid request');

    });
    

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('stops removal if an unknown error message is thrown when deleting the owner', (done) => {
    let memberDeleteCallIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });


    sinon.stub(request, 'delete').callsFake((opts) => {
      memberDeleteCallIssued = true;

      // for example... you must have at least one owner
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "err": {
            "response": {
              "status": 400
            } 
          }
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      return Promise.reject('Invalid request');

    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", confirm: true } }, () => {
      try {
        assert(memberDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly retrieves user but does not find the Group Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.reject("Invalid object identifier");
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Invalid object identifier'))); done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and handle error removing owner from specified Microsoft 365 group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({ 
          response: { 
            status: 400,
            data: {
              error: { 'odata.error': { message: { value: 'Invalid object identifier' } } } }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Invalid object identifier'))); done();
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/id`) {
        return Promise.resolve({ 
          response: { 
            status: 200,
            data: {
              value: "00000000-0000-0000-0000-000000000000"
            }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.resolve({
          "err": {
            "response": {
              "status": 404
            } 
          }
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/00000000-0000-0000-0000-000000000001/$ref`) {
        return Promise.reject({ 
          response: { 
            status: 400,
            data: {
              error: { 'odata.error': { message: { value: 'Invalid object identifier' } } } }
          } 
        });
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } } as any, (err?: any) => {
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

    sinon.stub(request, 'delete').callsFake(() => {
      memberDeleteCallIssued = true;
      return Promise.resolve();
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});