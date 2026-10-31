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
import command, { options } from './sitedesign-run-list.js';

describe(commands.SITEDESIGN_RUN_LIST, () => {
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_RUN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ID', 'SiteDesignID', 'SiteDesignTitle', 'StartTime']);
  });

  it('gets information about site designs applied to the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return {
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
              "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
              "SiteDesignTitle": "Contoso Communication Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548959800000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a' }) });
    assert(loggerLogSpy.calledWith([{
      "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
      "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
      "SiteDesignTitle": "Contoso Team Site",
      "SiteDesignVersion": 1,
      "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
      "StartTime": new Date(1548960114000).toLocaleString(),
      "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
    },
    {
      "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
      "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
      "SiteDesignTitle": "Contoso Communication Site",
      "SiteDesignVersion": 1,
      "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
      "StartTime": new Date(1548959800000).toLocaleString(),
      "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
    }]));
  });

  it('gets information about the specified site design applied to the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return {
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteDesignId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' }) });
    assert(loggerLogSpy.calledWith([{
      "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
      "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
      "SiteDesignTitle": "Contoso Team Site",
      "SiteDesignVersion": 1,
      "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
      "StartTime": new Date(1548960114000).toLocaleString(),
      "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
    }]));
  });

  it('outputs all information in JSON output mode', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return {
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
              "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
              "SiteDesignTitle": "Contoso Communication Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548959800000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a', output: 'json' }) });
    assert(loggerLogSpy.calledWith([
      {
        "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
        "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
        "SiteDesignTitle": "Contoso Team Site",
        "SiteDesignVersion": 1,
        "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
        "StartTime": "1548960114000",
        "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
      },
      {
        "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
        "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
        "SiteDesignTitle": "Contoso Communication Site",
        "SiteDesignVersion": 1,
        "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
        "StartTime": "1548959800000",
        "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
      }
    ]));
  });

  it('correctly handles OData error when retrieving information about site designs', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/team-a' }) }), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if siteDesignId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', siteDesignId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if webUrl is valid and siteDesignId is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if webUrl and siteDesignId are valid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', siteDesignId: '6ec3ca5b-d04b-4381-b169-61378556d76e' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
