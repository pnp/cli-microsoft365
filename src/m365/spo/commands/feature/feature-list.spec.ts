import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './feature-list.js';

describe(commands.FEATURE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FEATURE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves available features from site collection', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      })
    });
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
  });

  it('retrieves available features from site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      })
    });
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
  });

  it('retrieves available features from site (default) when no scope is entered', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        throw 'Invalid request';
      }

      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com'
      })
    });
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
  });

  it('returns all properties for output JSON', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site',
        output: 'json'
      })
    });
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
  });

  it('correctly handles no features in site collection', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      })
    });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles no features in site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      })
    });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles web feature reject request', async () => {
    const err = 'Invalid web Features reject request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      })
    }), new CommandError(err));
  });

  it('correctly handles site Features reject request', async () => {
    const err = 'Invalid site Features reject request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      })
    }), new CommandError(err));
  });

  it('retrieves all Web features', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "00bfea71-5932-4f9c-ad71-1557e5751100",
              DisplayName: "WebPageLibrary"
            }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/abc', scope: 'Web' }) });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: '00bfea71-5932-4f9c-ad71-1557e5751100',
        DisplayName: 'WebPageLibrary'
      }]
    ));
  });

  it('retrieves all site features', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/abc', scope: 'Site' }) });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: '3019c9b4-e371-438d-98f6-0a08c34d06eb',
        DisplayName: 'TenantSitesList'
      }]
    ));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation when the url options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: "Site"
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Site', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Site'
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Web', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Web'
    });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('rejects invalid scope value specified as number', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 123
    });
    assert.strictEqual(actual.success, false);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
