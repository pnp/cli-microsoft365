import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./hubsite-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.HUBSITE_SET), true);
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
        assert.equal(telemetry.name, commands.HUBSITE_SET);
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

  it('aborts when not logged in to a SharePoint tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`https://contoso.sharepoint.com is not a tenant admin site. Log in to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the title of the specified hub site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">Sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "Description", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "Description",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the description of the specified hub site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">All things sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, description: 'All things sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "All things sales",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the logo URL of the specified hub site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">https://contoso.com/logo.png</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, logoUrl: 'https://contoso.com/logo.png', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "All things sales",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the title, description and logo URL of the specified hub site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">Sales</Parameter></SetProperty><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">https://contoso.com/logo.png</Parameter></SetProperty><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">All things sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, title: 'Sales', description: 'All things sales', logoUrl: 'https://contoso.com/logo.png', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "All things sales",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">&lt;Sales&gt;</Parameter></SetProperty><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">&lt;https://contoso.com/logo.png&gt;</Parameter></SetProperty><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">&lt;All things sales&gt;</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "<All things sales>", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "<https:\u002f\u002fcontoso.com\u002flogo.png>", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "<Sales>"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, title: '<Sales>', description: '<All things sales>', logoUrl: '<https://contoso.com/logo.png>', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "<All things sales>",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "<https://contoso.com/logo.png>",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "<Sales>"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('allows resetting hub site title', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "Description", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": ""
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "Description",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: ""
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('allows resetting hub site description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, description: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "https://contoso.com/logo.png",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('allows resetting hub site logo URL', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, logoUrl: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Description: "All things sales",
          ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
          LogoUrl: "",
          SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
          SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
          Title: "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "Invalid URL: Logo.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPFieldValidationException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, logoUrl: 'Logo', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Invalid URL: Logo.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures title as string option', () => {
    const types = (command.types() as CommandTypes);
    ['t', 'title', 'd', 'description', 'l', 'logoUrl'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
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

  it('supports specifying hub site ID', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hub site title', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hub site description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hub site logoUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--logoUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if id not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Sales' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc', title: 'Sales' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if no property to update specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if id and title specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', title: 'Sales' } });
    assert.equal(actual, true);
  });

  it('passes validation if id and description specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', description: 'All things sales' } });
    assert.equal(actual, true);
  });

  it('passes validation if id and logoUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', logoUrl: 'https://contoso.com/logo.png' } });
    assert.equal(actual, true);
  });

  it('passes validation if all options specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', title: 'Sales', description: 'All things sales', logoUrl: 'https://contoso.com/logo.png' } });
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
    assert(find.calledWith(commands.HUBSITE_SET));
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
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
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