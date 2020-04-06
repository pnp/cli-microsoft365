import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.TEAMS_TEAM_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.TEAMS_TEAM_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',

      }
    });
    assert.equal(actual, true);
    done();
  });

  it('sets the visibility settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Public'
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'Public' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the mailNickName correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          mailNickName: 'NewNickName'
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', mailNickName: 'NewNickName' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the description settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          description: 'desc'
        })) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', description: 'desc' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Team visibility to public through isPrivate option', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Public'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', isPrivate: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Team visibility to private through isPrivate option', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee') {
        if (JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Private'
        })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', isPrivate: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('updates Team image with a png image (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/photo/$value' &&
        opts.headers['content-type'] === 'image/png') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'logo.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('updates Team image with a jpg image (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/photo/$value' &&
        opts.headers['content-type'] === 'image/jpeg') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'logo.jpg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Team image with a gif image (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/photo/$value' &&
        opts.headers['content-type'] === 'image/gif') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'logo.gif' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when updating Team image', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'logo.png' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when updating Team image (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'logo.png' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('sets the classification settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          classification: 'MBI'
        })) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', classification: 'MBI' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee",
            "innerError": {
              "request-id": "27b49647-a335-48f8-9a7c-f1ed9b976aaa",
              "date": "2019-04-05T12:16:48"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', displayName: 'NewName' }
    }, (err?: any) => {
      try {
        assert.equal(err.message, 'No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the teamId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } });
    assert.equal(actual, true);
  });

  it('fails validation if visibility is not a valid visibility Private|Public', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'hidden' } });
    assert.notEqual(actual, false);
  });
  it('passes validation if visibility is Public', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'Public' } });
    assert.equal(actual, true);
  });

  it('passes validation if visibility is Private', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'Private' } });
    assert.equal(actual, true);
  });

  it('fails validation if isPrivate is invalid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', isPrivate: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if isPrivate is true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', isPrivate: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if isPrivate is false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', isPrivate: 'false' } });
    assert.equal(actual, true);
  });
  it('fails validation if logoPath points to a non-existent file', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'invalid' } });
    Utils.restore(fs.existsSync);
    assert.notEqual(actual, true);
  });

  it('fails validation if logoPath points to a folder', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notEqual(actual, true);
  });

  it('passes validation if logoPath points to an existing file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_TEAM_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});