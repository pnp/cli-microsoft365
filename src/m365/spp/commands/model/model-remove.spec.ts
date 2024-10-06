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

describe(commands.MODEL_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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

  it('correctly handles site is not Content Site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'SITEPAGEPUBLISHING#0'
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } }),
      new CommandError('https://contoso.sharepoint.com/sites/portal is not a content site.'));
  });


  it('correctly handles an access denied error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        throw {
          error: {
            "odata.error": {
              message: {
                lang: "en-US",
                value: "Attempted to perform an unauthorized operation."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', force: true } }),
      new CommandError('Attempted to perform an unauthorized operation.'));
  });


  it('deletes model by id', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }
      throw 'Invalid request';
    });

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert.strictEqual(stubDelete.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`);
    assert(confirmationStub.calledOnce);
  });

  it('does not delete model when confirmation is not accepted', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }
      throw 'Invalid request';
    });

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

  it('deletes model by id with force', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }
      throw 'Invalid request';
    });

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '164720c8-35ee-4157-ba26-db6726264f9d', force: true } });
    assert.strictEqual(stubDelete.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`);
  });

  it('deletes model when the the site URL has trailing slash', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }
      throw 'Invalid request';
    });

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal/', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert.strictEqual(stubDelete.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`);
    assert(confirmationStub.calledOnce);
  });

  it('deletes model by title', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      throw 'Invalid request';
    });

    const stubDelete = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('ModelName')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName' } });
    assert.strictEqual(stubDelete.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('ModelName')`);
    assert(confirmationStub.calledOnce);
  });
});