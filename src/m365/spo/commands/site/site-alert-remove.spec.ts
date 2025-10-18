import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './site-alert-remove.js';

describe(commands.SITE_ALERT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;

  const webUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const alertId = '39d9e102-9e8f-4e74-8f17-84a92f972fcf';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.active = true;
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
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
    assert.strictEqual(command.name, commands.SITE_ALERT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if alertId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when valid webUrl and alertId are provided', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, id: alertId });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the alert', async () => {
    await command.action(logger, { options: { webUrl: webUrl, id: alertId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the alert when prompt is not confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { webUrl: webUrl, id: alertId } });
    assert(deleteStub.notCalled);
  });

  it('correctly removes the alert', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/Alerts/DeleteAlert('${formatting.encodeQueryParameter(alertId)}')`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: webUrl, id: alertId, force: true, verbose: true } });
    assert(deleteStub.calledOnce);
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2146232832, Microsoft.SharePoint.SPException',
          message: {
            value: 'The alert you are trying to access does not exist or has just been deleted.'
          }
        }
      }
    };
    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ force: true, webUrl: webUrl, id: alertId }) }),
      new CommandError(error.error['odata.error'].message.value));
  });
});