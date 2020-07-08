import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./apppage-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.APPPAGE_ADD, () => {
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.APPPAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('creates a single-part app page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.body.webPartDataAsJson ===
        "{}" && !opts.body.addToQuickLaunch) {
        return Promise.resolve("Done");
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Done"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a single-part app page showing on quicklaunch', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.body.webPartDataAsJson ===
        "{}" && opts.body.addToQuickLaunch) {
        return Promise.resolve("Done");
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, addToQuickLaunch: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Done"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to create a single-part app page if request is rejected', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.body.title === "failme") {
        return Promise.reject('Failed to create a single-part app page');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, title: 'failme', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Failed to create a single-part app page`)));
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

  it('supports specifying title', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webPartData', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webPartData') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if title not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webPartData: JSON.stringify({ "abc": "def" }), webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });
  it('fails validation if webPartData not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });
  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webPartData: JSON.stringify({ "abc": "def" }), title: 'page.aspx' } });
    assert.notEqual(actual, true);
  });
  it('fails validation if webPartData is not a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webUrl: 'https://contoso', webPartData: 'abc' } });
    assert.notEqual(actual, true);
  });
  it('validation passes on all required options', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webPartData: '{}', webUrl: 'https://contoso.sharepoint.com' } });
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
    assert(find.calledWith(commands.APPPAGE_ADD));
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