import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./feature-list');

describe(commands.FEATURE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.FEATURE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves available features from site collection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available features from site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available features from site (default) when no scope is entered', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject('Invalid request');
      }

      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all properties for output JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              "DisplayName": "TenantSitesList"
            },
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              "DisplayName": "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      debug: true,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              "DisplayName": "TenantSitesList"
            },
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              "DisplayName": "Ratings"
            }
          ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no features in site collection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no features in site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no features in site collection (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      verbose: true,
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };
    command.action(logger, { options: options } as any, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No activated Features found') > -1) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no features in site (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      verbose: true,
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };
    command.action(logger, { options: options } as any, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No activated Features found') > -1) {
          correctLogStatement = true;
        }
      });
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles web feature reject request', (done) => {
    const err = 'Invalid web Features reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Web'
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

  it('correctly handles site Features reject request', (done) => {
    const err = 'Invalid site Features reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        verbose: true,
        url: 'https://contoso.sharepoint.com',
        scope: 'Site'
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

  it('retrieves all Web features', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "00bfea71-5932-4f9c-ad71-1557e5751100",
              DisplayName: "WebPageLibrary"
            }]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/abc', scope: 'Web' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            DefinitionId: '00bfea71-5932-4f9c-ad71-1557e5751100',
            DisplayName: 'WebPageLibrary'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all site features', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/abc', scope: 'Site' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            DefinitionId: '3019c9b4-e371-438d-98f6-0a08c34d06eb',
            DisplayName: 'TenantSitesList'
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
    const actual = command.validate({
      options:
      {
        url: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url options specified', () => {
    const actual = command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com",
        scope: "Site"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Site', () => {
    const actual = command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', () => {
    const actual = command.validate({
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
    const actual = command.validate({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('rejects invalid scope value specified as number', () => {
    const scope = 123;
    const actual = command.validate({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = command.validate(
      {
        options:
        {
          url: "https://contoso.sharepoint.com"
        }
      });
    assert.strictEqual(actual, true);
  });
});