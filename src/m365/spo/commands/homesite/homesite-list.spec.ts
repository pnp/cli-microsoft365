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
import command, { options } from './homesite-list.js';

describe(commands.HOMESITE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const homeSites = {
    "value": [
      {
        "Audiences": [
          {
            "Email": "ColumnSearchable@contoso.onmicrosoft.com",
            "Id": "978b5280-4f80-47ea-a1db-b0d1d2fb1ba4",
            "Title": "ColumnSearchable Members"
          },
          {
            "Email": "contosoteam@contoso.onmicrosoft.com",
            "Id": "21af775d-17b3-4637-94a4-2ba8625277cb",
            "Title": "Contoso TeamR Members"
          }
        ],
        "IsInDraftMode": false,
        "IsVivaBackendSite": false,
        "SiteId": "431d7819-4aaf-49a1-b664-b2fe9e609b63",
        "TargetedLicenseType": 2,
        "Title": "The Landing",
        "Url": "https://contoso.sharepoint.com/sites/TheLanding",
        "VivaConnectionsDefaultStart": true,
        "WebId": "626c1724-8ac8-45d5-af87-c07c752fab75"
      },
      {
        "Audiences": [],
        "IsInDraftMode": false,
        "IsVivaBackendSite": false,
        "SiteId": "45d4a135-40e4-4571-8340-61d17fdfd58a",
        "TargetedLicenseType": 0,
        "Title": "Contoso Electronics",
        "Url": "https://contoso.sharepoint.com/sites/contosoportal",
        "VivaConnectionsDefaultStart": true,
        "WebId": "9418e2a1-855c-4752-8dd4-48693f43b10a"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Url', 'Title']);
  });

  it('lists available home sites', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSites;
      }

      throw opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true }) });
    assert(loggerLogSpy.calledWith(homeSites.value));
  });

  it('correctly handles OData error when retrieving available home sites', async () => {
    sinon.stub(request, 'get').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('An error has occurred'));
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
