import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
const command: Command = require('./homesite-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import appInsights from '../../../../appInsights';

describe(commands.HOMESITE_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.HOMESITE_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('gets information about the Home Site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP.SPHSite/Details') {
        return Promise.resolve({
          "SiteId": "53ad95dc-5d2c-42a3-a63c-716f7b8014f5",
          "WebId": "288ce497-483c-4cd5-b8a2-27b726d002e2",
          "LogoUrl": "https://contoso.sharepoint.com/sites/Work/siteassets/work.png",
          "Title": "Work @ Contoso",
          "Url": "https://contoso.sharepoint.com/sites/Work"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "SiteId": "53ad95dc-5d2c-42a3-a63c-716f7b8014f5",
          "WebId": "288ce497-483c-4cd5-b8a2-27b726d002e2",
          "LogoUrl": "https://contoso.sharepoint.com/sites/Work/siteassets/work.png",
          "Title": "Work @ Contoso",
          "Url": "https://contoso.sharepoint.com/sites/Work"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the Home Site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP.SPHSite/Details') {
        return Promise.resolve({
          "SiteId": "53ad95dc-5d2c-42a3-a63c-716f7b8014f5",
          "WebId": "288ce497-483c-4cd5-b8a2-27b726d002e2",
          "LogoUrl": "https://contoso.sharepoint.com/sites/Work/siteassets/work.png",
          "Title": "Work @ Contoso",
          "Url": "https://contoso.sharepoint.com/sites/Work"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`doesn't output anything when information about the Home Site is not available`, (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_api/SP.SPHSite/Details') {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.HOMESITE_GET));
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