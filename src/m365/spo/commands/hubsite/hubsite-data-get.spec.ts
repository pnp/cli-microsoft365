import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './hubsite-data-get.js';

describe(commands.HUBSITE_DATA_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_DATA_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the specified hub site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData(false)`) > -1) {
        return {
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } });
    assert(loggerLogSpy.calledWith({
      "themeKey": null,
      "name": "CommunicationSite",
      "url": "https://contoso.sharepoint.com/sites/Sales",
      "logoUrl": "http://contoso.com/__siteIcon__.jpg",
      "usesMetadataNavigation": false,
      "navigation": []
    }));
  });

  it('gets information about the specified hub site with forced refresh', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData(true)`) > -1) {
        return {
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/Project-X', forceRefresh: true } });
    assert(loggerLogSpy.calledWith({
      "themeKey": null,
      "name": "CommunicationSite",
      "url": "https://contoso.sharepoint.com/sites/Sales",
      "logoUrl": "http://contoso.com/__siteIcon__.jpg",
      "usesMetadataNavigation": false,
      "navigation": []
    }));
  });

  it('gets information about the specified hub site (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return {
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } });
    assert(loggerLogSpy.calledWith({
      "themeKey": null,
      "name": "CommunicationSite",
      "url": "https://contoso.sharepoint.com/sites/Sales",
      "logoUrl": "http://contoso.com/__siteIcon__.jpg",
      "usesMetadataNavigation": false,
      "navigation": []
    }));
  });

  it('correctly handles empty response', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error when specified site is not connect to or is a hub site (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } });
    assert(loggerLogToStderrSpy.calledWith(`https://contoso.sharepoint.com/sites/Project-X is not connected to a hub site and is not a hub site itself`));
  });

  it('correctly handles error when hub site not found', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        "odata.error": {
          "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
          "message": {
            "lang": "en-US",
            "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
