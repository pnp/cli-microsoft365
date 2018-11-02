import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./tenant-deletedsite-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.TENANT_DELETEDSITE_LIST, () => {
  let vorpal: Vorpal;
  let log: any[];
  let requests: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc' }); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    requests = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
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
      auth.ensureAccessToken,
      auth.restoreAuth,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_DELETEDSITE_LIST), true);
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
        assert.equal(telemetry.name, commands.TENANT_DELETEDSITE_LIST);
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
    assert(find.calledWith(commands.TENANT_DELETEDSITE_LIST));
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

  it('handles promise error while retrieving deleted sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.reject('An error has occurred');
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {

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

  it('handles error while retrieving deleted sites', (done) => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1216", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "b6229e9e-10db-0000-3501-d4b481a9ed04", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "b6229e9e-10db-0000-3501-d4b481a9ed04"
          }
        ]));
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {

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

  it('retrieve the deleted sites (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8224.1216",
            "ErrorInfo": null,
            "TraceCorrelationId": "090c9e9e-4064-0000-3501-ded574252e4b"
          },
          32,
          {
            "IsNull": false
          },
          34,
          {
            "IsNull": false
          },
          35,
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable",
            "NextStartIndex": -1,
            "NextStartIndexFromSharePoint": null,
            "_Child_Items_": [
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fninja", "DaysRemaining": 81, "DeletionTime": "/Date(2018,9,21,15,1,1,530)/", "SiteId": "/Guid(442d44ee-c4f0-4455-ba5e-39c7cde66cb1)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/ninja", "UserCodeMaximumLevel": 300 },
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2ftokyo", "DaysRemaining": 84, "DeletionTime": "/Date(2018,9,24,8,46,2,167)/", "SiteId": "/Guid(641daa1a-c3bd-4880-97aa-64af3d02a330)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/tokyo", "UserCodeMaximumLevel": 300 }
            ]
          }
        ]));
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
          Url: 'https://contoso.sharepoint.com/sites/ninja',
          StorageMaximumLevel: 26214400,
          UserCodeMaximumLevel: 300,
          DeletionTime: '2018-10-21T13:01:01.530Z',
          DaysRemaining: 81
        }, {
          Url: 'https://contoso.sharepoint.com/sites/tokyo',
          StorageMaximumLevel: 26214400,
          UserCodeMaximumLevel: 300,
          DeletionTime: '2018-10-24T06:46:02.167Z',
          DaysRemaining: 84
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieve the deleted sites', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8224.1216",
            "ErrorInfo": null,
            "TraceCorrelationId": "090c9e9e-4064-0000-3501-ded574252e4b"
          },
          32,
          {
            "IsNull": false
          },
          34,
          {
            "IsNull": false
          },
          35,
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable",
            "NextStartIndex": -1,
            "NextStartIndexFromSharePoint": null,
            "_Child_Items_": [
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fninja", "DaysRemaining": 81, "DeletionTime": "/Date(2018,9,21,15,1,1,530)/", "SiteId": "/Guid(442d44ee-c4f0-4455-ba5e-39c7cde66cb1)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/ninja", "UserCodeMaximumLevel": 300 },
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2ftokyo", "DaysRemaining": 84, "DeletionTime": "/Date(2018,9,24,8,46,2,167)/", "SiteId": "/Guid(641daa1a-c3bd-4880-97aa-64af3d02a330)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/tokyo", "UserCodeMaximumLevel": 300 }
            ]
          }
        ]));
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
          Url: 'https://contoso.sharepoint.com/sites/ninja',
          StorageMaximumLevel: 26214400,
          UserCodeMaximumLevel: 300,
          DeletionTime: '2018-10-21T13:01:01.530Z',
          DaysRemaining: 81
        }, {
          Url: 'https://contoso.sharepoint.com/sites/tokyo',
          StorageMaximumLevel: 26214400,
          UserCodeMaximumLevel: 300,
          DeletionTime: '2018-10-24T06:46:02.167Z',
          DaysRemaining: 84
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no deleted sites in recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8224.1216",
            "ErrorInfo": null,
            "TraceCorrelationId": "090c9e9e-4064-0000-3501-ded574252e4b"
          },
          32,
          {
            "IsNull": false
          },
          34,
          {
            "IsNull": false
          },
          35,
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable",
            "NextStartIndex": -1,
            "NextStartIndexFromSharePoint": null,
            "_Child_Items_": null
          }
        ]));
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {

      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows message when there are no deleted sites in recycle bin (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8224.1216",
            "ErrorInfo": null,
            "TraceCorrelationId": "090c9e9e-4064-0000-3501-ded574252e4b"
          },
          32,
          {
            "IsNull": false
          },
          34,
          {
            "IsNull": false
          },
          35,
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable",
            "NextStartIndex": -1,
            "NextStartIndexFromSharePoint": null,
            "_Child_Items_": []
          }
        ]));
      }
      if (opts.url.indexOf('contextinfo') > -1) {
        return Promise.resolve('abc');
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No deleted site collections found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  /*
    it('handles tenant settings error', (done) => {
      // get tenant app catalog
      sinon.stub(request, 'post').callsFake((opts) => {
        requests.push(opts);
        if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
            }, 58, {
              "IsNull": false
            }, 59, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
            }
          ]));
        }
        if (opts.url.indexOf('contextinfo') > -1) {
          return Promise.resolve('abc');
        }
        return Promise.reject('Invalid request');
      });
  
      auth.site = new Site();
      auth.site.connected = true;
      auth.site.url = 'https://contoso-admin.sharepoint.com';
      cmdInstance.action = command.action();
      cmdInstance.action({
        options: {
  
        }
      }, (err?: any) => {
        try {
          assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Timed out')));
          done();
        }
        catch (e) {
          done(e);
        }
      });
    });
  */
  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8224.1216",
            "ErrorInfo": null,
            "TraceCorrelationId": "090c9e9e-4064-0000-3501-ded574252e4b"
          },
          32,
          {
            "IsNull": false
          },
          34,
          {
            "IsNull": false
          },
          35,
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable",
            "NextStartIndex": -1,
            "NextStartIndexFromSharePoint": null,
            "_Child_Items_": [
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fninja", "DaysRemaining": 81, "DeletionTime": "/Date(2018,9,21,15,1,1,530)/", "SiteId": "/Guid(442d44ee-c4f0-4455-ba5e-39c7cde66cb1)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/ninja", "UserCodeMaximumLevel": 300 },
              { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "090c9e9e-4064-0000-3501-ded574252e4b|908bed80-a04a-4433-b4a0-883d9847d110:0acf1d97-fe55-40a7-9b63-7aff8d24784e\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2ftokyo", "DaysRemaining": 84, "DeletionTime": "/Date(2018,9,24,8,46,2,167)/", "SiteId": "/Guid(641daa1a-c3bd-4880-97aa-64af3d02a330)/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https://contoso.sharepoint.com/sites/tokyo", "UserCodeMaximumLevel": 300 }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
          DaysRemaining: 81,
          DeletionTime: '2018-10-21T13:01:01.530Z',
          SiteId: "442d44ee-c4f0-4455-ba5e-39c7cde66cb1",
          Status: "Recycled",
          StorageMaximumLevel: 26214400,
          Url: 'https://contoso.sharepoint.com/sites/ninja',
          UserCodeMaximumLevel: 300,
        }, {
          DaysRemaining: 84,
          DeletionTime: '2018-10-24T06:46:02.167Z',
          SiteId: "641daa1a-c3bd-4880-97aa-64af3d02a330",
          Status: "Recycled",
          StorageMaximumLevel: 26214400,
          Url: 'https://contoso.sharepoint.com/sites/tokyo',
          UserCodeMaximumLevel: 300
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});