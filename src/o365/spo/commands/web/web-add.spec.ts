import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./web-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.WEB_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
 
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc' }); });
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
      sinon.stub,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.WEB_ADD), true);
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
        assert.equal(telemetry.name, commands.WEB_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', title: 'Documents' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
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
    assert(find.calledWith(commands.WEB_ADD));
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

  it('fails validation if the title option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  webUrl: "subsite",
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('passes validation if all the options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", webUrl: "subsite",
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.equal(actual, true);
  });

  it('fails validation if the weburl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the parentWeburl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    webUrl:"subsite", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webTemplate option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    webUrl:"subsite", parentWebUrl:"https://contoso.sharepoint.com", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('creates web without inheriting the navigation', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }
      else if(opts.url.indexOf('_api/contextinfo') > -1){
        return Promise.resolve({ FormDigestValue: 'abc' }); 
      }

      return Promise.resolve('abc');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      local:1033,
      breakInheritance : true,
      inheritNavigation : false,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      done();
    });
  });

  it('creates web and does not set the inherit navigation (Noscript enabled)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }

      return Promise.resolve('abc');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        // PermissionKind.ManageLists, PermissionKind.AddListItems, PermissionKind.DeleteListItems
        return Promise.resolve(
          { High:2058,
            Low:0
          }
        );
      }

      return Promise.resolve('abc');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033
    } }, () => {

      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert(cmdInstanceLogSpy.calledWith("No script is enabled. Skipping the InheritParentNavigation settings."));
      done();
    });
  });

  
  it('creates web and inherits the navigation (debug)', (done) => {
    // Create web
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }
      else if ((opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) && (opts.body.indexOf("UseShared") > -1)) {
        return Promise.resolve([
          {
          "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7317.1203","ErrorInfo":null,"TraceCorrelationId":"4556449e-0067-4000-1529-39a0d88e307d"
          },1,{
          "IsNull":false
          },3,{
          "IsNull":false
          },5,{
          "IsNull":false
          },7,{
          "_ObjectType_":"SP.Navigation","UseShared":true
          }
          ]
        );
      }

      return Promise.resolve('abc');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          { High:2147483647,
            Low:4294967295
          }
        );
      }

      return Promise.resolve('abc');
    });
   
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert(cmdInstanceLogSpy.calledWith("Setting the navigation as the parent site."));
      assert(cmdInstanceLogSpy.calledWith("Response : SetInheritNavigation"));
      done();
    });
  });

  it('creates web and inherits the navigation', (done) => {
    // Create web
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }
      else if ((opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) && (opts.body.indexOf("UseShared") > -1)) {
        return Promise.resolve([
          {
          "SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7317.1203","ErrorInfo":null,"TraceCorrelationId":"4556449e-0067-4000-1529-39a0d88e307d"
          },1,{
          "IsNull":false
          },3,{
          "IsNull":false
          },5,{
          "IsNull":false
          },7,{
          "_ObjectType_":"SP.Navigation","UseShared":true
          }
          ]
        );
      }

      return Promise.resolve('abc');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          { High:2147483647,
            Low:4294967295
          }
        );
      }

      return Promise.resolve('abc');
    });
   
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: false
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert.equal(false, cmdInstanceLogSpy.calledWith("Setting the navigation as the parent site."));
      assert.equal(false, cmdInstanceLogSpy.calledWith("Response : SetInheritNavigation"));
      done();
    });
  });

  it('correctly handles the set inheritNavigation error', (done) => {
    // Create web
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }
      else if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        // SetInheritNavigation failed.
        return Promise.reject(false);
      }

      return Promise.resolve('abc');
    });
    // Full permission.
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          { High:2147483647,
            Low:4294967295
          }
        );
      }

      return Promise.resolve('abc');
    });
   
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert(cmdInstanceLogSpy.calledWith(new CommandError(`Failed to set inheritNavigation for the web - https://contoso.sharepoint.com/subsite`)));
      done();
    });
  });

  it('creates web and handles the subsite contextinfo call error while getting the effectivebasepermission', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }
      else if(opts.url.indexOf('/subsite/_api/contextinfo') > -1) {
        return Promise.reject(false);
      }

      return Promise.resolve('abc');
    });
    
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert(cmdInstanceLogSpy.calledWith(new CommandError('Failed to get the contextinfo for the web - https://contoso.sharepoint.com/subsite')));
      done();
    });
  });

  it('correctly handles the parentweb contextinfo call error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if(opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.reject(false);
      }
      return Promise.resolve('abc');
    });
    
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(new CommandError('Failed to get the contextinfo for the web - https://contoso.sharepoint.com')));
      done();
    });
  });


  it('correctly handles the createweb call error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.reject({"error":{"code":"-2147024713, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"The Web site address \"/subsite\" is already in use."}}});
      }
      return Promise.resolve('abc');
    });
   
    
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(new CommandError('Failed to create the web - subsite')));
      done();
    });
  });

  it('creates web and handles the effectivebasepermission call error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }

      return Promise.resolve('abc');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.reject("Failed to get the effectivebase permissions.");
      }

      return Promise.resolve('abc');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green(`Subsite subsite created.`)));
      assert(cmdInstanceLogSpy.calledWith(new CommandError("Failed to get the effectivebase permissions.")));
      done();
    });
  });

  it('correctly handles the getAccessToken error', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl:"https://contoso.sharepoint.com",
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

});