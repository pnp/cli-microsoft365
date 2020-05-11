import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandCancel, CommandValidate } from '../../../../Command';
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
  let maxAttempts: number = 5;

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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com' } });
    assert(actual);
  });

  it('handles REST API call error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
          return Promise.reject({ error: 'Invalid request' });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Invalid request')));
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

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            return Promise.reject({ error: 'Invalid request' });
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Invalid request')));
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

  it('restores the deleted site collection from the tenant recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

  it('restores the deleted site collection from the tenant recycle bin (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {debug: true, url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('site collection restored'));
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

  it('restores the deleted site collection from the tenant recycle bin (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
          return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {verbose: true, url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('site collection restored'));
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

  it('restores the deleted site collection from the tenant recycle bin, wait for completion', (done) => {
    let attempt: number = 0;
    const stub = sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            attempt++;
            if (attempt <= maxAttempts) {
              return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});    
            }
            else {
              return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
            }
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(stub.called);
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

  it('restores the deleted site collection from the tenant recycle bin, wait for completion (debug)', (done) => {
    let attempt: number = 0;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            attempt++;
            if (attempt <= maxAttempts) {
              return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});    
            }
            else {
              return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
            }
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: {debug: true, url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('site collection restored'));
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

  it('restores the deleted site collection from the tenant recycle bin, wait for completion (verbose)', (done) => {
    let attempt: number = 0;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            attempt++;
            if (attempt <= maxAttempts) {
              return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});    
            }
            else {
              return Promise.resolve({"HasTimedout": false, "IsComplete": true, "PollingInterval": 15000});
            }
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: {verbose: true, url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('site collection restored'));
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

  it('did not restore the deleted site collection from the tenant recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify([{"HasTimedout": true, "IsComplete": false, "PollingInterval": 15000}]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/hr', wait: false } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('site collection has not been restored')));
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

  it('did not restore the deleted site collection from the tenant recycle bin, after waiting for completion (timeout)', (done) => {
    let attempt: number = 0;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            attempt++;
            if (attempt <= maxAttempts) {
              return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});    
            }
            else {
              return Promise.resolve({"HasTimedout": true, "IsComplete": false, "PollingInterval": 15000});
            }
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: {url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Operation timeout')));
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

  it('did not restore the deleted site collection from the tenant recycle bin, after waiting for completion (error thrown)', (done) => {
    let attempt: number = 0;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers && opts.headers.accept && opts.headers['content-type'] && opts.body &&
          opts.headers.accept.indexOf('application/json') === 0 && opts.headers['content-type'].indexOf('application/json') === 0) {
            attempt++;
            if (attempt <= maxAttempts) {
              return Promise.resolve({"HasTimedout": false, "IsComplete": false, "PollingInterval": 15000});    
            }
            else {
              return Promise.reject('Operation timeout');
            }
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: {url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith('site collection has not been restored'));
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