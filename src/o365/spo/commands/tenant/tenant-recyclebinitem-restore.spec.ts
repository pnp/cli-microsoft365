import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandCancel } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-recyclebinitem-restore');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_RECYCLEBINITEM_RESTORE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let requests: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    requests = [];
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
      request.post,
      global.setTimeout,
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
    assert.equal(command.name.startsWith(commands.TENANT_RECYCLEBINITEM_RESTORE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TENANT_RECYCLEBINITEM_RESTORE));
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

  it('handles REST API call error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('handles REST API call error with waiting', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('did not restore the deleted Site Collection from the Tenant Recycle Bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify([{"HasTimedout": true, "IsComplete": false, "PollingInterval": 15000}]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', wait: false } }, (err?: any) => {
      try {
        assert.equal(err, "Site Collection has not been restored");
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
          r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
          r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        assert(cmdInstanceLogSpy.calledWith('Site Collection restored'));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin, wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: {debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
          r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin, wait for completion (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: {verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
          r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
        r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('restores the deleted Site Collection from the Tenant Recycle Bin, wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers.accept && opts.headers.contenttype && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers.contenttype.indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    cmdInstance.action({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1 &&
        r.headers.accept && r.headers.contenttype && r.body &&
          r.headers.accept.indexOf('application/json') === 0 && r.headers.contenttype.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });
});