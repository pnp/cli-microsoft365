import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-run-list');

describe(commands.SITEDESIGN_RUN_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_RUN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ID', 'SiteDesignID', 'SiteDesignTitle', 'StartTime']);
  });

  it('gets information about site designs applied to the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site design applied to the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteDesignId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
          "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
          "SiteDesignTitle": "Contoso Team Site",
          "SiteDesignVersion": 1,
          "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
          "StartTime": new Date(1548960114000).toLocaleString(),
          "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all information in JSON output mode', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', output: 'json' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving information about site designs', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteDesignId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', siteDesignId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl is valid and siteDesignId is not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if webUrl and siteDesignId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', siteDesignId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});