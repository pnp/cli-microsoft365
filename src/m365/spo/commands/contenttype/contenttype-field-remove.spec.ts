import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError, CommandTypes } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

const command: Command = require('./contenttype-field-remove');
const WEB_URL = 'https://contoso.sharepoint.com';
const FIELD_LINK_ID = "5ee2dd25-d941-455a-9bdb-7f2c54aed11b";
const CONTENT_TYPE_ID = "0x0100558D85B7216F6A489A499DB361E1AE2F";
const LIST_CONTENT_TYPE_ID = "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A";
const LIST_ID = "8c7a0fcd-9d64-4634-85ea-ce2b37b2ec0c";
const WEB_ID = "d1b7a30d-7c22-4c54-a686-f1c298ced3c7";
const SITE_ID = "50720268-eff5-48e0-835e-de588b007927";
const LIST_TITLE = "TEST";

describe(commands.CONTENTTYPE_FIELD_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let promptOptions: any;

  const getStubCalls = (opts: any) => {
    if ((opts.url as string).indexOf(`_api/site?$select=Id`) > -1) {
      return Promise.resolve({
        "Id": SITE_ID
      });
    }
    if ((opts.url as string).indexOf(`_api/web?$select=Id`) > -1) {
      return Promise.resolve({
        "Id": WEB_ID
      });
    }
    if ((opts.url as string).indexOf(`/_api/lists/GetByTitle('${LIST_TITLE}')?$select=Id`) > -1) {
      return Promise.resolve({
        "Id": LIST_ID
      });
    }

    return Promise.reject('Invalid request');
  }
  const postStubSuccCalls = (opts: any) => {
    if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
      // Web CT
      if (opts.data.toLowerCase() === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`.toLowerCase()) {
        return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
      }
      // Web CT with update child CTs
      else if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
        return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
      }
      // List CT
      else if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:list:${LIST_ID}:contenttype:${LIST_CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
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
  }
  const postStubFailedCalls = (opts: any) => {
    if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
      // WEB CT
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
        return Promise.resolve(`[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": {
              "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
            },
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`);
      }
      // Web CT without update child CTs
      else if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:contenttype:${CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
        return Promise.resolve(`[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": {
              "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
            },
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`);
      }
      // LIST CT
      else if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${FIELD_LINK_ID}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${SITE_ID}:web:${WEB_ID}:list:${LIST_ID}:contenttype:${LIST_CONTENT_TYPE_ID}" /></ObjectPaths></Request>`) {
        return Promise.resolve(`[
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7911.1206",
            "ErrorInfo": {
              "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
            },
            "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
          }
        ]`);
      }

    }
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    loggerSpy = sinon.spy(logger, 'log');
    (command as any).requestDigest = '';
    (command as any).webId = '';
    (command as any).siteId = '';
    (command as any).listId = '';
    (command as any).fieldLinkId = '';
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_FIELD_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('configures contentTypeId as string option', () => {
    const types = (command.types() as CommandTypes);
    ['i', 'contentTypeId'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  // WEB CT
  it('removes the field link from web content type', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: false,
        debug: false
      }
    } as any, (err?: any) => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('doesnt remove the field link from web content type - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // WEB CT: with debug
  it('removes the field link from web content type with debug', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: true,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type with debug - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        let promptIssued = false;

        if (promptOptions && promptOptions.type === 'confirm') {
          promptIssued = true;
        }
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type with debug - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesnt remove the field link from web content type with debug - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // WEB CT: with update child content types
  it('removes the field link from web content type with update child content types', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(loggerSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type with update child content types - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        let promptIssued = false;

        if (promptOptions && promptOptions.type === 'confirm') {
          promptIssued = true;
        }
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type with update child content types - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('doesnt remove the field link from web content type with update child content types - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // WEB CT: with update child content types with debug
  it('removes the field link from web content type with update child content types with debug', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: true,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from web content type with update child content types with debug - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        let promptIssued = false;

        if (promptOptions && promptOptions.type === 'confirm') {
          promptIssued = true;
        }
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesnt remove the field link from web content type with update child content types with debug - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesnt remove the field link from web content type with update child content types with debug - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // LIST CT
  it('removes the field link from list content type', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        confirm: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from list content type - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        debug: false
      }
    } as any, (err?: any) => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field link from list content type - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: true,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field link from list content type - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: false,
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        assert(loggerSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // LIST CT with debug
  it('removes the field link from list content type with debug', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        confirm: true,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from list content type with debug - prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        debug: true
      }
    } as any, (err?: any) => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from list content type with debug - prompt: confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.called);
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('removes the field link from list content type with debug - prompt: declined', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    const postCallbackStub = sinon.stub(request, 'post').callsFake(postStubSuccCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, listTitle: LIST_TITLE, contentTypeId: LIST_CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: false,
        debug: true
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // Handles error
  it('handles error when remove the field link from web content type with update child content types', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubFailedCalls);

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: true,
        debug: false,
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('handles error when remove the field link from web content type with update child content types (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubFailedCalls);

    command.action(logger, { options: { debug: true, webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID, updateChildContentTypes: true, confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when remove the field link from web content type with update child content types with prompt', (done) => {
    sinon.stub(request, 'get').callsFake(getStubCalls);
    sinon.stub(request, 'post').callsFake(postStubFailedCalls);

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: false,
        debug: false,
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles a random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, {
      options: {
        webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID, fieldLinkId: FIELD_LINK_ID,
        updateChildContentTypes: true,
        confirm: true,
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // Fails validation
  it('fails validation if fieldLinkId is not passed', () => {
    const actual = command.validate({ options: { webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not passed', () => {
    const actual = command.validate({ options: { fieldLinkId: FIELD_LINK_ID, contentTypeId: CONTENT_TYPE_ID } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not correct', () => {
    const actual = command.validate({ options: { fieldLinkId: FIELD_LINK_ID, contentTypeId: CONTENT_TYPE_ID, webUrl: "test" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if fieldLinkId is not valid GUID', () => {
    const actual = command.validate({ options: { fieldLinkId: 'xxx', webUrl: WEB_URL, contentTypeId: CONTENT_TYPE_ID } });
    assert.notStrictEqual(actual, true);
  });

  // Passes validation
  it('passes validation', () => {
    const actual = command.validate({ options: { listId: LIST_ID, fieldLinkId: FIELD_LINK_ID, contentTypeId: CONTENT_TYPE_ID, webUrl: WEB_URL, debug: true } });
    assert.strictEqual(actual, true);
  });
});