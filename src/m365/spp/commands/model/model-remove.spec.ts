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
import command from './model-remove.js';
import { spp } from '../../../../utils/spp.js';

describe(commands.MODEL_REMOVE, () => {
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

  it('passes validation when required parameters are valid with id', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with title', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', title: 'ModelName' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with id and force', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'invalidUrl', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles an error when the model title is not found', async () => {
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

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } }),
      new CommandError('File Not Found.'));
  });

  it('is the confirmation prompt called', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert(confirmationStub.calledOnce);
  });

  it('deletes model by id', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } });
    assert(stubDelete.calledOnce);
  });

  it('does not delete model when confirmation is not accepted', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert(stubDelete.notCalled);
    assert(confirmationStub.calledOnce);
  });

  it('deletes model by id without force when confirmation is accepted', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '164720c8-35ee-4157-ba26-db6726264f9d' } });
    assert(stubDelete.calledOnce);
    assert(confirmationStub.calledOnce);
  });

  it('deletes model when the the site URL has trailing slash', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal/', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } });
    assert(stubDelete.calledOnce);
  });

  it('deletes model by title', async () => {
    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('ModelName')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName', force: true } });
    assert(stubDelete.calledOnce);
  });
});