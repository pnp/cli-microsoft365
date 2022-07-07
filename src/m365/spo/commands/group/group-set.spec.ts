import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-set');

const validId = 1;
const validName = "Project leaders";
const validWebUrl = 'https://contoso.sharepoint.com/sites/project-x';
const validOwnerEmail = 'john.doe@contoso.com';
const validOwnerUserName = 'john.doe@contoso.com';

const userInfoResponse = {
  userPrincipalName: validOwnerUserName
};

const ensureUserResponse = {
  Id: 3
};

describe(commands.GROUP_SET, () => {
  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.patch,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name, commands.GROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets();
    assert.deepStrictEqual(optionSets, [['id', 'name']]);
  });

  it('fails validation when group id is not a number', (done) => {
    const actual = command.validate({
      options: {
        webUrl: validWebUrl,
        id: 'invalid id'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerEmail and ownerUserName are specified', (done) => {
    const actual = command.validate({
      options: {
        webUrl: validWebUrl,
        ownerEmail: validOwnerEmail,
        ownerUserName: validOwnerUserName
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when invalid boolean is passed as option', (done) => {
    const actual = command.validate({
      options: {
        webUrl: validWebUrl,
        id: validId,
        autoAcceptRequestToJoinLeave: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when invalid web URL is passed', (done) => {
    const actual = command.validate({
      options: {
        webUrl: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid options specified', (done) => {
    const actual = command.validate({
      options: {
        webUrl: validWebUrl,
        id: validId,
        allowRequestToJoinLeave: 'true'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('successfully updates group settings by id', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        allowRequestToJoinLeave: 'true'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully updates group settings by name', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        allowRequestToJoinLeave: 'true'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully updates group owner by ownerEmail', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    }));
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return Promise.resolve(ensureUserResponse);
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        ownerEmail: validOwnerEmail
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully updates group owner by ownerEmail', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    }));
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return Promise.resolve(ensureUserResponse);
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        ownerUserName: validOwnerUserName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        autoAcceptRequestToJoinLeave: 'true'
      }
    }, (err?: any) => {
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