import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-recyclebinitem-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_RECYCLEBINITEM_LIST, () => {
  let log: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_RECYCLEBINITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('handles client.svc promise error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while getting tenant recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc"
          }
        ]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('includes all properties for json output', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19527.12077", "ErrorInfo": null, "TraceCorrelationId": "85bb2b9f-5099-2000-af64-2c100126d549"
          }, 54, {
            "IsNull": false
          }, 56, {
            "IsNull": false
          }, 57, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
          }, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists the tenant recyclebin items (debug)', (done) => {

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19527.12077", "ErrorInfo": null, "TraceCorrelationId": "85bb2b9f-5099-2000-af64-2c100126d549"
          }, 54, {
            "IsNull": false
          }, 56, {
            "IsNull": false
          }, 57, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }
            ]
          }
        ]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0]["DaysRemaining"], 92);
        assert.deepEqual(cmdInstanceLogSpy.lastCall.args[0][0]["DeletionTime"], new Date(2020, 0, 15, 11, 4, 3, 893));
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0]["Url"], 'https://contoso.sharepoint.com/sites/ClassicThrowaway');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][1].DaysRemaining, 92);
        assert.deepEqual(cmdInstanceLogSpy.lastCall.args[0][1].DeletionTime, new Date(2020, 0, 15, 11, 40, 58, 90));
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][1].Url, 'https://contoso.sharepoint.com/sites/ModernThrowaway');

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Orders retrieved sites by url ascending', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19527.12077", "ErrorInfo": null, "TraceCorrelationId": "85bb2b9f-5099-2000-af64-2c100126d549"
          }, 54, {
            "IsNull": false
          }, 56, {
            "IsNull": false
          }, 57, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DaysRemaining: 92,
            DeletionTime: new Date(2020, 0, 15, 11, 4, 3, 893),
            Url: 'https://contoso.sharepoint.com/sites/ClassicThrowaway'
          },
          {
            DaysRemaining: 92,
            DeletionTime: new Date(2020, 0, 15, 11, 40, 58, 90),
            Url: 'https://contoso.sharepoint.com/sites/ModernThrowaway'
          },
          {
            DaysRemaining: 92,
            DeletionTime: new Date(2020, 0, 15, 11, 40, 58, 90),
            Url: 'https://contoso.sharepoint.com/sites/ModernThrowaway'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles tenant recyclebin timeout', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
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
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Timed out')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});