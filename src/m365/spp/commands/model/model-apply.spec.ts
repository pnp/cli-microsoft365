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
import { z } from 'zod';

describe(commands.MODEL_APPLY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  const publicationsResult = {
    Details: [
      {
        ErrorMessage: null,
        Publication: {},
        StatusCode: 201
      }
    ],
    TotalFailures: 0,
    TotalSuccesses: 1
  };
  const listResponse = {
    RootFolder: {
      ServerRelativeUrl: '/sites/portal/Shared Documents'
    },
    BaseType: 1
  };
  const modelResponse = {
    UniqueId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spp, 'assertSiteIsContentCenter').resolves();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      request.get,
      request.post,
      spp.getModelById,
      spp.getModelByTitle
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
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with model title and list id', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with model title and list title', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listTitle: 'Documents' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with model title and list URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listUrl: '/Shared Documents' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with model title and list id and correct viewOption is provided', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when webUrl is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalidUrl', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when contentCenterUrl is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'invalidUrl', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when model id is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: 'invalidId', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'NewViewAsDefault' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when list id is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: 'invalidId', viewOption: 'NewViewAsDefault' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when viewOption is not valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'InvalidViewOption' });
    assert.strictEqual(actual.success, false);
  });

  it('does not log any output', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: false } });
    assert(loggerLogSpy.notCalled);
  });

  it('applies a model to document library by id and list id', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } });
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

  it('applies a model to document library by id and list id to subsite', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/subsite/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return {
          RootFolder: {
            ServerRelativeUrl: '/sites/portal/subsite/Shared Documents'
          },
          BaseType: 1
        };
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales/subsite', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } });
    assert.deepStrictEqual(stubPost.lastCall.args[0].data, {
      __metadata: {
        type: "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData"
      },
      Publications: {
        results: [
          {
            ModelUniqueId: "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            TargetSiteUrl: "https://contoso.sharepoint.com/sites/sales/subsite",
            TargetLibraryServerRelativeUrl: "/sites/portal/subsite/Shared Documents",
            TargetWebServerRelativeUrl: "/sites/sales/subsite",
            ViewOption: "NewViewAsDefault"
          }
        ]
      }
    });
  });

  it('applies a model to document library by id and list title', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('Documents')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listTitle: 'Documents' } });
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

  it('applies a model to document library by title and list id', async () => {
    sinon.stub(spp, 'getModelByTitle').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } });
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

  it('applies a model to document library by id and list url', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('%2Fsites%2Fsales%2FShared%20Documents')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listUrl: '/Shared Documents' } });
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

  it('applies a model to document library by id and list id and DoNotChangeDefault viewOption', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'DoNotChangeDefault' } });
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

  it('applies a model to document library by id and list id and TileViewAsDefault viewOption', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'TileViewAsDefault' } });
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
            ViewOption: "TileViewAsDefault"
          }
        ]
      }
    });
  });

  it('applies a model to document library by title with classifier suffix and by list id', async () => {
    sinon.stub(spp, 'getModelByTitle').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    const stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return publicationsResult;
      }

      throw `${opts.url} is invalid request`;
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle.classifier', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', viewOption: 'DoNotChangeDefault' } });
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
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
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

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d' } }),
      new CommandError('List does not exist. The page you selected contains a list that does not exist. It may have been deleted by another user.'));
  });

  it('correctly handles error when trying to apply a model to SP list', async () => {
    sinon.stub(spp, 'getModelById').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
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

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } }),
      new CommandError('The specified list is not a document library.'));
  });

  it('correctly handles error when model is not found by its id', async () => {
    sinon.stub(spp, 'getModelById').rejects({
      error: {
        "odata.error": {
          code: "-1, Microsoft.Office.Server.ContentCenter.ModelNotFoundException",
          message: {
            lang: "en-US",
            value: "File Not Found."
          }
        }
      }
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } }),
      new CommandError('File Not Found.'));
  });

  it('correctly handles error when applying a model failed', async () => {
    sinon.stub(spp, 'getModelByTitle').resolves(modelResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'421b1e42-794b-4c71-93ac-5ed92488b67d')?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder`) {
        return listResponse;
      }

      throw `${opts.url} is invalid request`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/contentCenter/_api/machinelearning/publications`) {
        return {
          Details: [
            {
              ErrorMessage: 'The content type is bound to another model. Please unpublish the existing model or remove the content type in order to publish the model.',
              Publication: {},
              StatusCode: 400
            }
          ],
          TotalFailures: 1,
          TotalSuccesses: 0
        };
      }

      throw `${opts.url} is invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentCenterUrl: 'https://contoso.sharepoint.com/sites/contentCenter', title: 'ModelTitle.classifier', listId: '421b1e42-794b-4c71-93ac-5ed92488b67d', verbose: true } }),
      new CommandError('The content type is bound to another model. Please unpublish the existing model or remove the content type in order to publish the model.'));
  });
});