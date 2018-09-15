import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./navigation-node-remove');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.NAVIGATION_NODE_REMOVE, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.delete
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      (command as any).getRequestDigestForSite,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.NAVIGATION_NODE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.NAVIGATION_NODE_REMOVE);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes navigation node from the top navigation', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', confirm: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes navigation node from the top navigation (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', confirm: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing navigation node when confirmation argument not passed', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } }, () => {
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

  it('aborts removing app when prompt not confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject('Invalid request');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the navigation node when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject({ error: 'An error has occurred' });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', confirm: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error (string error)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', confirm: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { location: 'TopNavigationBar', id: '2003' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', location: 'TopNavigationBar', id: '2003' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if location is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', id: '2003' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified location is not valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid', id: '2003' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } });
    assert.equal(actual, true);
  });

  it('passes validation when location is QuickLaunch and all required properties are present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', id: '2003' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.NAVIGATION_NODE_REMOVE));
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', confirm: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});