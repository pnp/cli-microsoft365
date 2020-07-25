import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as chalk from 'chalk';

describe(commands.O365GROUP_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
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
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.put,
      request.patch,
      request.get,
      global.setTimeout
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      fs.readFileSync,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates Microsoft 365 Group display name', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          displayName: 'My group'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group description (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          description: 'My group'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', description: 'My group' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group to public', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Public'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group to private', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Private'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group logo with a png image', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/png') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'logo.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group logo with a jpg image (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/jpeg') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.jpg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Group logo with a gif image', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/gif') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.gif' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when updating Microsoft 365 Group logo', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { debug: false, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.png' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when updating Microsoft 365 Group logo (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.png' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds owner to Microsoft 365 Group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', owners: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds owners to Microsoft 365 Group (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com' or userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            },
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds member to Microsoft 365 Group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', members: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds members to Microsoft 365 Group (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com' or userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid', description: 'My awesome group' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID and displayName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID and description specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', description: 'My awesome group' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if no property to update is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of the owners is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the owner is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated with an additional space', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if one of the members is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the member is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated with an additional space', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if isPrivate is invalid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if isPrivate is true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if isPrivate is false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if logoPath points to a non-existent file', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'invalid' } });
    Utils.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if logoPath points to a folder', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if logoPath points to an existing file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying owners', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--owners') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying members', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--members') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying group type', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isPrivate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying logo file path', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--logoPath') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});