import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './recyclebinitem-restore.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';

describe(commands.OWNER_LIST, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';
  const flowName = 'd87a7535-dd31-4437-bfe1-95340acd55c6';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the flowName is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: environmentName, flowName: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the flowName is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: environmentName, flowName: flowName });
    assert.strictEqual(actual.success, true);
  });

  it('outputs no command output', async () => {
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        flowName: flowName,
        verbose: true
      }
    });

    assert(loggerLogSpy.notCalled);
  });

  it('correctly restores flow', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${flowName}/restore?api-version=2016-11-01`) {
        return;
      }

      throw 'Invalid request :' + opts.url;
    });

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        flowName: flowName
      }
    });

    assert(postStub.calledOnce);
  });

  it('correctly handles error when restoring flow', async () => {
    const message = 'Request to Azure Resource Manager failed.';

    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          code: 'AzureResourceManagerRequestFailed',
          message: message
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName } }),
      new CommandError(message));
  });
});