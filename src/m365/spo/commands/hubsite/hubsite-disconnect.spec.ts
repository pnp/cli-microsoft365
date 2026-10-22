import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command, { options } from './hubsite-disconnect.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.HUBSITE_DISCONNECT, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const id = '55b979e7-36b6-4968-b3af-6ae221a3483f';
  const title = 'Hub Site';
  const url = 'https://contoso.sharepoint.com/sites/HubSite';

  const etagValue = '1';
  const singleHubSiteResponse = {
    'odata.etag': etagValue,
    ID: id,
    Title: title,
    SiteUrl: url
  };

  const hubSitesResponse = {
    value: [
      singleHubSiteResponse,
      {
        'odata.etag': etagValue,
        ID: 'a9d15b9d-152c-4fa2-be3a-3fbf086f3d49',
        Title: 'Random site title',
        SiteUrl: 'https://contoso.sharepoint.com/sites/RandomSite'
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;
  let patchStub: sinon.SinonStub<[options: CliRequestOptions]>;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;

    sinon.stub(spo, 'getSpoAdminUrl').resolves(spoAdminUrl);
    patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/HubSites/GetById('${id}')`) {
        return;
      }

      throw 'Invalid requet URL: ' + opts.url;
    });
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_DISCONNECT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'Invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'Invalid' is not a valid GUID`);
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'Invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'Invalid' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if valid id is specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: id });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if valid title is specified', () => {
    const actual = commandOptionsSchema.safeParse({ title: title });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if valid url is specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: url });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if id and title are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, title: title });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if no hub site is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before disconnecting the hub site when confirmation argument not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: id }) });

    assert(promptIssued);
  });

  it('aborts disconnecting hub site when prompt not confirmed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ url: url }) });
    assert(patchStub.notCalled);
  });

  it('disconnects hub site when when id option is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites/GetById('${id}')`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return singleHubSiteResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: id,
        verbose: true,
        force: true
      })
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when when title option is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: title,
        verbose: true,
        force: true
      })
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when when url option is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: url,
        verbose: true,
        force: true
      })
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: title
      })
    });
  });

  it('throws an error when multiple hub sites with the same title were retrieved', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const response = {
      value: [
        singleHubSiteResponse,
        {
          'odata.etag': etagValue,
          ID: 'a9d15b9d-152c-4fa2-be3a-3fbf086f3d49',
          Title: title,
          SiteUrl: 'https://contoso.sharepoint.com/sites/RandomSite'
        }
      ]
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return response;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        title: title,
        force: true
      })
    }), new CommandError("Multiple hub sites with name 'Hub Site' found. Found: 55b979e7-36b6-4968-b3af-6ae221a3483f, a9d15b9d-152c-4fa2-be3a-3fbf086f3d49."));
  });

  it('handles selecting single result when multiple hubsites with the specified name found and cli is set to prompt', async () => {
    const response = {
      value: [
        singleHubSiteResponse,
        {
          'odata.etag': etagValue,
          ID: 'a9d15b9d-152c-4fa2-be3a-3fbf086f3d49',
          Title: title,
          SiteUrl: 'https://contoso.sharepoint.com/sites/RandomSite'
        }
      ]
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return response;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
      Title: title,
      ID: id,
      'odata.etag': etagValue
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: title,
        verbose: true,
        force: true
      })
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('throws an error when no hub sites with the same title were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return {
            value: []
          };
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        title: title,
        force: true
      })
    }), new CommandError(`The specified hub site '${title}' does not exist.`));
  });

  it('throws an error when no hub sites with the same url were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return {
            value: []
          };
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: url,
        force: true
      })
    }), new CommandError(`The specified hub site '${url}' does not exist.`));
  });

  it('handles random API error correctly', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/HubSites/GetById('${id}')?$select=ID`) {
        return singleHubSiteResponse;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    const errorMessage = 'Something went wrong';
    patchStub.restore();
    sinon.stub(request, 'patch').rejects({ error: { 'odata.error': { message: { value: errorMessage } } } });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: id,
        force: true
      })
    }), new CommandError(errorMessage));
  });
});
