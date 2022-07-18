import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./customaction-remove');

describe(commands.CUSTOMACTION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };

    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    assert.deepStrictEqual(command.optionSets(), [
      ['id', 'title']
    ]);
  });

  it('handles error when multiple user custom actions with the specified title found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
          value: [
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request`);
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'YourAppCustomizer',
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple user custom actions with title 'YourAppCustomizer' found. Please disambiguate using IDs: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when no user custom actions with the specified title found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
          value: [
          ]
        });
      }

      return Promise.reject(`Invalid request`);
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'YourAppCustomizer',
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No user custom action with title 'YourAppCustomizer' found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should user custom action removed successfully without prompting with confirmation argument', (done) => {
    defaultPostCallsStub();

    command.action(logger, {
      options: {
        verbose: false,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should prompt before removing custom action when confirmation argument not passed', (done) => {
    command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' } }, () => {
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

  it('should abort custom action remove when prompt not confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' } } as any, () => {
      try {
        assert(postCallsSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should remove custom action by id when prompt confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' } } as any, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com'
          })));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should remove custom action by title when prompt confirmed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
          value: [
            {
              "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
              "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
              "CommandUIExtension": null,
              "Description": null,
              "Group": null,
              "Id": "b2307a39-e878-458b-bc90-03bc578531d6",
              "ImageUrl": null,
              "Location": "ClientSideExtension.ApplicationCustomizer",
              "Name": "{b2307a39-e878-458b-bc90-03bc578531d6}",
              "RegistrationId": null,
              "RegistrationType": 0,
              "Rights": { "High": 0, "Low": 0 },
              "Scope": "1",
              "ScriptBlock": null,
              "ScriptSrc": null,
              "Sequence": 65536,
              "Title": "Places",
              "Url": null,
              "VersionOfUserCustomAction": "1.0.1.0"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { title: 'Places', url: 'https://contoso.sharepoint.com' } } as any, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith(sinon.match(
          {
            title: 'Places',
            url: 'https://contoso.sharepoint.com'
          })));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is Web', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith({
          debug: false,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          url: 'https://contoso.sharepoint.com',
          scope: 'Web',
          confirm: true
        }), 'removeScopedCustomActionSpy data error');
        assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is Site', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site',
            confirm: true
          }), 'removeScopedCustomActionSpy data error');
        assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is All, but item found on web level', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    command.action(logger, {
      options: {
        confirm: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called twice when scope is All, but item not found on web level', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    command.action(logger, {
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {
      try {
        assert(removeScopedCustomActionSpy.calledTwice);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should searchAllScopes be called when scope is All', (done) => {
    defaultPostCallsStub();

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            confirm: true
          })), 'searchAllScopesSpy.calledWith');
        assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['searchAllScopes']);
      }
    });
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    command.action(logger, {
      options: {
        verbose: false,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(`Custom action with id ${actionId} not found`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle custom action reject request (web)', (done) => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    command.action(logger, {
      options: {
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }, (error: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle custom action reject request (site)', (done) => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // should return null to proceed with site when scope is All
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    command.action(logger, {
      options: {
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
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
    const options = command.options();
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying scope', () => {
    const options = command.options();
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
    const options = command.options();
    sinonUtil.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('should fail validation if the url option not specified', () => {
    const actual = command.validate({ options: { id: "BC448D63-484F-49C5-AB8C-96B14AA68D50" } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the id option is not a valid guid', () => {
    const actual = command.validate({
      options:
      {
        id: "foo",
        url: 'https://contoso.sharepoint.com'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the id and url options specified', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the id, url and scope options specified', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: "Site"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the id and url option specified', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: 'All'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Site', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Web', () => {
    const actual = command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: 'Web'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should reject invalid string scope', () => {
    const scope = 'foo';
    const actual = command.validate({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('should reject invalid scope value specified as number', () => {
    const scope = 123;
    const actual = command.validate({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = command.validate(
      {
        options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
      });
    assert.strictEqual(actual, true);
  });
});