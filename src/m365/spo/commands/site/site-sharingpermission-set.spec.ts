import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-sharingpermission-set.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';

describe(commands.SITE_SHARINGPERMISSION_SET, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/marketing';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let patchStub: sinon.SinonStub;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    loggerLogSpy = sinon.spy(logger, 'log');

    patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `${siteUrl}/_api/Web`) {
        return;
      }
      if (opts.url === `${siteUrl}/_api/Web/AssociatedMemberGroup`) {
        return;
      }

      throw 'Invalid request :' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_SHARINGPERMISSION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if siteUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'invalid', capability: 'full' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when capability is not a valid value', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, capability: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the input is correct', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: siteUrl, capability: 'limited' });
    assert.strictEqual(actual.success, true);
  });

  it('outputs no command output', async () => {
    patchStub.restore();
    sinon.stub(request, 'patch').resolves();

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        capability: 'full',
        verbose: true
      }
    });

    assert(loggerLogSpy.notCalled);
  });

  it('correctly sets sharing permissions to full', async () => {
    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        capability: 'full'
      }
    });

    assert.deepStrictEqual(patchStub.firstCall.args[0].data, { MembersCanShare: true });
    assert.deepStrictEqual(patchStub.secondCall.args[0].data, { AllowMembersEditMembership: true });
  });

  it('correctly sets sharing permissions to limited', async () => {
    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        capability: 'limited'
      }
    });

    assert.deepStrictEqual(patchStub.firstCall.args[0].data, { MembersCanShare: true });
    assert.deepStrictEqual(patchStub.secondCall.args[0].data, { AllowMembersEditMembership: false });
  });

  it('correctly sets sharing permissions to ownersOnly', async () => {
    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        capability: 'ownersOnly'
      }
    });

    assert.deepStrictEqual(patchStub.firstCall.args[0].data, { MembersCanShare: false });
    assert.deepStrictEqual(patchStub.secondCall.args[0].data, { AllowMembersEditMembership: false });
  });

  it('correctly handles error when updating sharing permissions', async () => {
    patchStub.restore();
    const errorMessage = 'Access is denied.';

    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          message: {
            lang: 'en-US',
            value: errorMessage
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, capability: 'limited' } }),
      new CommandError(errorMessage));
  });
});