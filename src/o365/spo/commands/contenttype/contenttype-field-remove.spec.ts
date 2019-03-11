import commands from '../../commands';
import Command, { CommandError, CommandTypes, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./contenttype-field-remove');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

const FIELD_LINK_ID = "5ee2dd25-d941-455a-9bdb-7f2c54aed11b";
const CONTENT_TYPE_ID = "0x0100558D85B7216F6A489A499DB361E1AE2F";
const WEB_ID = "d1b7a30d-7c22-4c54-a686-f1c298ced3c7";
const SITE_ID = "50720268-eff5-48e0-835e-de588b007927";

describe(commands.CONTENTTYPE_FIELD_REMOVE, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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

    (command as any).fieldLinkId = null;
    (command as any).contentTypeId = null;
    (command as any).updateChildContentTypes = null;

    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.CONTENTTYPE_FIELD_REMOVE), true);
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
        assert.equal(telemetry.name, commands.CONTENTTYPE_FIELD_REMOVE);
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

  it('removes the field link from web content type with update child content types', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": WEB_ID
        });
      }
      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": SITE_ID
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });


    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: CONTENT_TYPE_ID, fieldId: FIELD_LINK_ID, updateChildContentTypes: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field link from web content type with update child content types (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": WEB_ID
        });
      }
      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": SITE_ID
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });


    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: CONTENT_TYPE_ID, fieldId: FIELD_LINK_ID, updateChildContentTypes: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field link from web content type without update child content types', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": WEB_ID
        });
      }
      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": SITE_ID
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });


    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: CONTENT_TYPE_ID, fieldId: FIELD_LINK_ID, updateChildContentTypes: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field link from web content type with update child content types (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": WEB_ID
        });
      }
      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": SITE_ID
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });


    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: CONTENT_TYPE_ID, fieldId: FIELD_LINK_ID, updateChildContentTypes: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

  it('configures contentTypeId as string option', () => {
    const types = (command.types() as CommandTypes);
    ['i', 'contentTypeId'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
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
    assert(find.calledWith(commands.CONTENTTYPE_FIELD_REMOVE));
  });

  it('fails validation if contentTypeId is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });
  it('fails validation if fieldLinkId is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });
  it('fails validation if webUrl is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { fieldLinkId: FIELD_LINK_ID } });
    assert.notEqual(actual, true);
  });
  it('fails validation if fieldLinkId is not valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { fieldLinkId: 'xxx' } });
    assert.notEqual(actual, true);
  });

});