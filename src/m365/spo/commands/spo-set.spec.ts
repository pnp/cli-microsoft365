import commands from '../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
const command: Command = require('./spo-set');
import * as assert from 'assert';
import Utils from '../../../Utils';
import auth from '../../../Auth';

describe(commands.SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../vorpal-init');
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find
    ]);
    auth.service.spoUrl = undefined;
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      auth.storeConnectionInfo,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('sets SPO URL when no URL was set previously', (done) => {
    auth.service.spoUrl = undefined;

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert.equal(auth.service.spoUrl, 'https://contoso.sharepoint.com');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets SPO URL when other URL was set previously', (done) => {
    auth.service.spoUrl = 'https://northwind.sharepoint.com';

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert.equal(auth.service.spoUrl, 'https://contoso.sharepoint.com');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('throws error when trying to set SPO URL when not logged in to O365', (done) => {
    auth.service.connected = false;

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to Microsoft 365 first')));
        assert.equal(auth.service.spoUrl, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('throws error when setting the password fails', (done) => {
    auth.service.connected = true;
    Utils.restore(auth.storeConnectionInfo);
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred while setting the password'))

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred while setting the password')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying url', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if url not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/team-a' } });
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
    assert(find.calledWith(commands.SET));
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