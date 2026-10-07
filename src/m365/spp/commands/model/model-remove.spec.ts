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
import command, { options } from './model-remove.js';
import { spp } from '../../../../utils/spp.js';

describe(commands.MODEL_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spp, 'assertSiteIsContentCenter').resolves();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MODEL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when required parameters are valid with id', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with title', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', title: 'ModelName' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when required parameters are valid with id and force', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when siteUrl is not valid', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'invalidUrl', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when id is not valid', () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sales',
      id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly handles an error when the model id is not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        throw {
          error: {
            "odata.error": {
              code: "-1, Microsoft.Office.Server.ContentCenter.ModelNotFoundException",
              message: {
                lang: "en-US",
                value: "File Not Found."
              }
            }
          }
        };
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true }) }),
      new CommandError('File Not Found.'));
  });

  it('correctly handles an error when the model title is not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modeltitle.classifier')`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'modelTitle.classifier', force: true }) }),
      new CommandError('Model not found.'));
  });

  it('is the confirmation prompt called with id information', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' }) });
    assert(confirmationStub.args[0][0].message.startsWith(`Are you sure you want to remove model '9b1b1e42-794b-4c71-93ac-5ed92488b67f'?`));
  });

  it('is the confirmation prompt called with title information', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'modelTitle' }) });
    assert(confirmationStub.args[0][0].message.startsWith(`Are you sure you want to remove model 'modelTitle'?`));
  });

  it('deletes model by id', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true }) });
    assert(stubDelete.calledOnce);
  });

  it('does not delete model when confirmation is not accepted', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const stubDelete = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' }) });
    assert(stubDelete.notCalled);
  });

  it('deletes model when the the site URL has trailing slash', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal/', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true }) });
    assert(stubDelete.calledOnce);
  });

  it('deletes model by title', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName', force: true }) });
    assert(stubDelete.calledOnce);
  });

  it('deletes model by title with .classifier suffix', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName.classifier', force: true }) });
    assert(stubDelete.calledOnce);
  });
});