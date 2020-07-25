import commands from '../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
const command: Command = require('./spo-set');
import * as assert from 'assert';
import Utils from '../../../Utils';
import auth from '../../../Auth';

describe(commands.SET, () => {
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
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
  });

  afterEach(() => {
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
    assert.strictEqual(command.name.startsWith(commands.SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets SPO URL when no URL was set previously', (done) => {
    auth.service.spoUrl = undefined;

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
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
        assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
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
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Log in to Microsoft 365 first')));
        assert.strictEqual(auth.service.spoUrl, undefined);
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
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred while setting the password')));
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

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert.strictEqual(actual, true);
  });
});