import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate, CommandCancel } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.TEAMS_TEAM_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get,
      fs.existsSync,
      fs.readFileSync,
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
    assert.equal(command.name.startsWith(commands.TEAMS_TEAM_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('passes validation if name and description are passed when no template is passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('passes validation if name and description are not passed when a template is supplied', (done) => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({
      options: {
        templatePath: 'template.json'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('fails validation if description is not passed when no template is supplied', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if name is not passed when no template is supplied', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        description: 'Architecture Discussion'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if template not found', (done) => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = (command.validate() as CommandValidate)({
      options: {
        templatePath: 'abc'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('creates Microsoft Teams team in the tenant (verbose) when no template is supplied', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Engineering Team',
            description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template and name is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template and description is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Engineering Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied and waits for command to complete (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequest = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({ status: 'succeeded' });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(getRequest.called);
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when creating a Team', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles failed operation error when creating a Team when waiting for command to complete', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequest = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({ status: 'failed' });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, (err?: any) => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(getRequest.called);
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('failed')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles invalid operation error when creating a Team when waiting for command to complete', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequest = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({ status: 'invalid' });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, (err?: any) => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(getRequest.called);
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('invalid')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles InProgress operation status when creating a Team and waiting for the command to complete', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequest = sinon.stub(request, 'get')
    getRequest.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({ status: "InProgress" });
        }
        return Promise.reject('Invalid request');
      });
    getRequest.onCall(1).callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({ status: "succeeded" });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, (err?: any) => {
      try {
        assert.deepEqual(requestStub.getCall(0).args, [{
          url: `https://graph.microsoft.com/beta/teams`,
          resolveWithFullResponse: true,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            displayName: 'Sample Classroom Team',
            description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
          },
          json: true
        }]);
        assert(getRequest.calledTwice);
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('can be cancelled', () => {
    assert(command.cancel());
  });

  it('clears pending connection on cancel', () => {
    (command as any).timeout = {};
    const clearTimeoutSpy = sinon.spy(global, 'clearTimeout');
    (command.cancel() as CommandCancel)();
    Utils.restore(global.clearTimeout);
    assert(clearTimeoutSpy.called);
  });

  it('doesn\'t fail on cancel if no connection pending', () => {
    (command as any).timeout = undefined;
    (command.cancel() as CommandCancel)();
    assert(true);
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
    assert(find.calledWith(commands.TEAMS_TEAM_ADD));
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