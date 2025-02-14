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
import command from './list-list.js';

describe(commands.LIST_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const listResponse = {
    value: [{
      AllowContentTypes: true,
      BaseTemplate: 109,
      BaseType: 1,
      ContentTypesEnabled: false,
      CrawlNonDefaultViews: false,
      Created: null,
      CurrentChangeToken: null,
      CustomActionElements: null,
      DefaultContentApprovalWorkflowId: "00000000-0000-0000-0000-000000000000",
      DefaultItemOpenUseListSetting: false,
      Description: "",
      Direction: "none",
      DocumentTemplateUrl: null,
      DraftVersionVisibility: 0,
      EnableAttachments: false,
      EnableFolderCreation: true,
      EnableMinorVersions: false,
      EnableModeration: false,
      EnableVersioning: false,
      EntityTypeName: "Documents",
      ExemptFromBlockDownloadOfNonViewableFiles: false,
      FileSavePostProcessingEnabled: false,
      ForceCheckout: false,
      HasExternalDataSource: false,
      Hidden: false,
      Id: "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
      ImagePath: null,
      ImageUrl: null,
      IrmEnabled: false,
      IrmExpire: false,
      IrmReject: false,
      IsApplicationList: false,
      IsCatalog: false,
      IsPrivate: false,
      ItemCount: 69,
      LastItemDeletedDate: null,
      LastItemModifiedDate: null,
      LastItemUserModifiedDate: null,
      ListExperienceOptions: 0,
      ListItemEntityTypeFullName: null,
      MajorVersionLimit: 0,
      MajorWithMinorVersionsLimit: 0,
      MultipleDataList: false,
      NoCrawl: false,
      ParentWebPath: null,
      ParentWebUrl: null,
      ParserDisabled: false,
      ServerTemplateCanCreateFolders: true,
      TemplateFeatureId: null,
      Title: "Documents",
      RootFolder: { ServerRelativeUrl: "Documents" }
    }]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    assert.strictEqual(command.name, commands.LIST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Id']);
  });

  it('retrieves all lists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*') {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith(listResponse.value));
  });

  it('retrieves all lists when properties provided', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate,ParentWebUrl') {
        return {
          value: [{
            BaseTemplate: 100,
            ParentWebUrl: "/",
            RootFolder: {
              ServerRelativeUrl: "/Lists/test"
            }
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'BaseTemplate,ParentWebUrl,RootFolder/ServerRelativeUrl'
      }
    });
    assert(loggerLogSpy.calledWith([{
      BaseTemplate: 100,
      ParentWebUrl: "/",
      RootFolder: {
        ServerRelativeUrl: "/Lists/test"
      }
    }]));
  });

  it('retrieves the lists based on the provided filter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*&$filter=BaseTemplate eq 100') {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        filter: 'BaseTemplate eq 100'
      }
    });
    assert(loggerLogSpy.calledWith(listResponse.value));
  });

  it('command correctly handles list list reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*') {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert(actual);
  });
});