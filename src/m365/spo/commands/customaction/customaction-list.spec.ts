import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./customaction-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('getCustomActions called once when scope is Web', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getRequestSpy.calledOnce);
        assert(getCustomActionsSpy.calledWith({
          debug: false,
          url: 'https://contoso.sharepoint.com',
          scope: 'Web'
        }));
        assert(getCustomActionsSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('getCustomActions called once when scope is Site', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(getRequestSpy.calledOnce);
        assert(getCustomActionsSpy.calledWith({
          debug: false,
          url: 'https://contoso.sharepoint.com',
          scope: 'Site'
        }));
        assert(getCustomActionsSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('returns all properties for output JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [{ "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }, { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }] });
      }

      return Promise.reject('Invalid request');
    });

    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{"ClientSideComponentId":"b41916e7-e69d-467f-b37f-ff8ecf8f99f2","ClientSideComponentProperties":"{\"testMessage\":\"Test message\"}","CommandUIExtension":null,"Description":null,"Group":null,"Id":"8b86123a-3194-49cf-b167-c044b613a48a","ImageUrl":null,"Location":"ClientSideExtension.ApplicationCustomizer","Name":"YourName","RegistrationId":null,"RegistrationType":0,"Rights":{"High":"0","Low":"0"},"Scope":3,"ScriptBlock":null,"ScriptSrc":null,"Sequence":0,"Title":"YourAppCustomizer","Url":null,"VersionOfUserCustomAction":"16.0.1.0"},{"ClientSideComponentId":"b41916e7-e69d-467f-b37f-ff8ecf8f99f2","ClientSideComponentProperties":"{\"testMessage\":\"Test message\"}","CommandUIExtension":null,"Description":null,"Group":null,"Id":"9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5","ImageUrl":null,"Location":"ClientSideExtension.ApplicationCustomizer","Name":"YourName","RegistrationId":null,"RegistrationType":0,"Rights":{"High":"0","Low":"0"},"Scope":3,"ScriptBlock":null,"ScriptSrc":null,"Sequence":0,"Title":"YourAppCustomizer","Url":null,"VersionOfUserCustomAction":"16.0.1.0"}]));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('getCustomActions called twice when scope is All', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');

    cmdInstance.action({
      options: {
        debug: true,
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(getRequestSpy.calledTwice);
        assert(getCustomActionsSpy.calledTwice);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('searchAllScopes called when scope is All', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: "All"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            url: 'https://contoso.sharepoint.com'
          })));
        assert(searchAllScopesSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['searchAllScopes']);
      }
    });
  });

  it('searchAllScopes correctly handles no custom actions when All scope specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
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

  it('correctly handles no custom actions when All scope specified (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`Custom actions not found`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles web custom action reject request', (done) => {
    const err = 'Invalid web custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site custom action reject request', (done) => {
    const err = 'Invalid site custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying scope', () => {
    const options = (command.options() as CommandOption[]);
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('retrieves all available user custom actions', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/abc' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Name: 'customaction2',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Site',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df56'
          },
          {
            Name: 'customaction1',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Web',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df55'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no scope entered (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ 
      options: { 
        url: 'https://contoso.sharepoint.com/sites/abc',
        debug: true 
      } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('Attempt to get custom actions list with scope: All') > -1) {
          correctLogStatement = true;
        }
      })
      try {
        assert(correctLogStatement);
        assert(cmdInstanceLogSpy.calledWith([
          {
            Name: 'customaction2',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Site',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df56'
          },
          {
            Name: 'customaction1',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Web',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df55'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: 'foo'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: "Site"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('humanize scope shows correct value when scope odata is 2', () => {
    const actual = (command as any)["humanizeScope"](2);
    assert(actual === "Site");
  });

  it('humanize scope shows correct value when scope odata is 3', () => {
    const actual = (command as any)["humanizeScope"](3);
    assert(actual === "Web");
  });

  it('humanize scope shows the scope odata value when is different than 2 and 3', () => {
    const actual = (command as any)["humanizeScope"](1);
    assert(actual === "1");
  });

  it('accepts scope to be All', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'All'
        }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Site', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Site'
        }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Web'
        }
    });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid string scope', () => {
    const scope = 'foo';
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('rejects invalid scope value specified as number', () => {
    const scope = 123;
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
          {
            url: "https://contoso.sharepoint.com"
          }
      });
    assert.strictEqual(actual, true);
  });
});