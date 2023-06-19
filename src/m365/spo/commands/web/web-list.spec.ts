import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./web-list');

describe(commands.WEB_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEB_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Url', 'Id']);
  });

  it('retrieves all webs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
        return {
          "value": [{
            "AllowRssFeeds": false,
            "AlternateCssUrl": null,
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "Configuration": 0,
            "Created": null,
            "CurrentChangeToken": null,
            "CustomMasterUrl": null,
            "Description": null,
            "DesignPackageId": null,
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "HorizontalQuickLaunch": false,
            "Id": "d8d179c7-f459-4f90-b592-14b08e84accb",
            "IsMultilingual": false,
            "Language": 1033,
            "LastItemModifiedDate": null,
            "LastItemUserModifiedDate": null,
            "MasterUrl": null,
            "NoCrawl": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": null,
            "QuickLaunchEnabled": false,
            "RecycleBinEnabled": false,
            "ServerRelativeUrl": null,
            "SiteLogoUrl": null,
            "SyndicationEnabled": false,
            "Title": "Subsite",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://Contoso.sharepoint.com/Subsite",
            "WebTemplate": "STS"
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        url: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith([{
      "AllowRssFeeds": false,
      "AlternateCssUrl": null,
      "AppInstanceId": "00000000-0000-0000-0000-000000000000",
      "Configuration": 0,
      "Created": null,
      "CurrentChangeToken": null,
      "CustomMasterUrl": null,
      "Description": null,
      "DesignPackageId": null,
      "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
      "EnableMinimalDownload": false,
      "HorizontalQuickLaunch": false,
      "Id": "d8d179c7-f459-4f90-b592-14b08e84accb",
      "IsMultilingual": false,
      "Language": 1033,
      "LastItemModifiedDate": null,
      "LastItemUserModifiedDate": null,
      "MasterUrl": null,
      "NoCrawl": false,
      "OverwriteTranslationsOnChange": false,
      "ResourcePath": null,
      "QuickLaunchEnabled": false,
      "RecycleBinEnabled": false,
      "ServerRelativeUrl": null,
      "SiteLogoUrl": null,
      "SyndicationEnabled": false,
      "Title": "Subsite",
      "TreeViewEnabled": false,
      "UIVersion": 15,
      "UIVersionConfigurationEnabled": false,
      "Url": "https://Contoso.sharepoint.com/Subsite",
      "WebTemplate": "STS"
    }]));
  });

  it('command correctly handles web list reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        url: 'https://contoso.sharepoint.com'
      }
    } as any), new CommandError(err));
  });

  it('uses correct API url when output json option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      logger.log('Test Url:');
      logger.log(opts.url);
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/webs') {
        return 'Correct Url1';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        url: 'https://contoso.sharepoint.com'
      }
    });
    assert('Correct Url');
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<url>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
