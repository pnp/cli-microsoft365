import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './model-apply.js';
import { spp } from '../../../../utils/spp.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MODEL_APPLY, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spp, 'assertSiteIsContentCenter').resolves();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MODEL_APPLY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when required parameters are valid with model id and list id', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with model title and list id', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with model id and list title', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listTitle: 'Documents' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with model id and list URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listUrl: '/Shared Documents' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with model title and list id and correct viewOption is provided', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'invalidUrl', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when contentCenterUrl is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'invalidUrl', title: 'ModelName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when model id is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: 'invalidId', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when list id is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listId: 'invalidId', viewOption: 'NewViewAsDefault' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when viewOption is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelName', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'InvalidViewOption' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('apply model to document library by id and list id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "NewViewAsDefault"
          }
        ]
      }
    });
  });

  it('apply model to document library by id and list title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('Documents')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listTitle: 'Documents' } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "NewViewAsDefault"
          }
        ]
      }
    });
  });

  it('apply model to document library by title and list id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbytitle('ModelTitle')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "NewViewAsDefault"
          }
        ]
      }
    });
  });

  it('apply model to document library by id and list url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('%2Fsites%2Fsales%2FShared%20Documents')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listUrl: '/Shared Documents' } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "NewViewAsDefault"
          }
        ]
      }
    });
  });

  it('apply model to document library by id and list id and DoNotChangeDefault viewOption', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'DoNotChangeDefault' } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "DoNotChangeDefault"
          }
        ]
      }
    });
  });

  it('apply model to document library by id and list id and TileViewAsDefault viewOption', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'DoNotChangeDefault' } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales",
            TargetLibraryServerRelativeUrl: "/sites/portal/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales",
            ViewOption: "DoNotChangeDefault"
          }
        ]
      }
    });
  });

  it('correctly handles error when list is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        throw {
          error: {
            "odata.error": {
              code: "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
              message: {
                lang: "en-US",
                value: "List does not exist. The page you selected contains a list that does not exist. It may have been deleted by another user."
              }
            }
          }
        };
      }

      throw `${opts.url} is invalid request`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return;
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } }), new CommandError('List does not exist. The page you selected contains a list that does not exist. It may have been deleted by another user.'));
  });

  it('corretly handles error when try to apply model to SP list', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return {
          UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/lists/SPList'
          },
          BaseType: 0
        };
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } }), new CommandError('The specified list is not a document library.'));
  });
});