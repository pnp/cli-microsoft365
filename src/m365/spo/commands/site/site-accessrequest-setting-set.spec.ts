import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-accessrequest-setting-set.js';
import { z } from 'zod';

describe(commands.SITE_ACCESSREQUEST_SETTING_SET, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/Management';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let baseSchema: z.ZodTypeAny;
  let refinedSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;

    commandInfo = cli.getCommandInfo(command);
    baseSchema = commandInfo.command.getSchemaToParse()!;
    refinedSchema = commandInfo.command.getRefinedSchema!(baseSchema as any)!;
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
      request.patch,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_ACCESSREQUEST_SETTING_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  // Schema validation
  it('fails validation if siteUrl is not a valid URL', () => {
    const actual = baseSchema.safeParse({ siteUrl: 'invalid', disabled: true });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when none of disabled, ownerGroup or email are provided', () => {
    const actual = refinedSchema.safeParse({ siteUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when two options among disabled, ownerGroup, email are provided', () => {
    const actual = refinedSchema.safeParse({ siteUrl, disabled: true, ownerGroup: true });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when message is used with disabled', () => {
    const actual = refinedSchema.safeParse({ siteUrl, disabled: true, message: 'not allowed' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when email is invalid', () => {
    const actual = refinedSchema.safeParse({ siteUrl, email: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when disabled specified', () => {
    const actual = refinedSchema.safeParse({ siteUrl, disabled: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when ownerGroup specified', () => {
    const actual = refinedSchema.safeParse({ siteUrl, ownerGroup: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when email specified with a valid email', () => {
    const actual = refinedSchema.safeParse({ siteUrl, email: 'john.doe@contoso.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when email specified with a message', () => {
    const actual = refinedSchema.safeParse({ siteUrl, email: 'john.doe@contoso.com', message: 'Hello' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when ownerGroup specified with a message', () => {
    const actual = refinedSchema.safeParse({ siteUrl, ownerGroup: true, message: 'Hello' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when email specified with empty message (clear message)', () => {
    const actual = refinedSchema.safeParse({ siteUrl, email: 'john.doe@contoso.com', message: '' });
    assert.strictEqual(actual.success, true);
  });

  // Execution tests: check direct URLs in stubs and assert call counts
  it('correctly disables access requests', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: '' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: false })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, disabled: true, verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledOnce, true);
  });

  it('sends access requests to owner group', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: '' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: true })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, ownerGroup: true, verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledOnce, true);
  });

  it('sends access requests to a specific email', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: 'john.doe@contoso.com' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: false })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, email: 'john.doe@contoso.com', verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledOnce, true);
  });

  it('sets custom message when provided together with email', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: 'john.doe@contoso.com' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: false })) {
        return;
      }
      if (opts.url === `${siteUrl}/_api/Web/SetAccessRequestSiteDescriptionAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ description: 'Motivate why you need access.' })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, email: 'john.doe@contoso.com', message: 'Motivate why you need access.', verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledTwice, true);
  });

  it('clears custom message when empty message provided together with email', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: 'john.doe@contoso.com' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: false })) {
        return;
      }
      if (opts.url === `${siteUrl}/_api/Web/SetAccessRequestSiteDescriptionAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ description: '' })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, email: 'john.doe@contoso.com', message: '', verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledTwice, true);
  });

  it('sets custom message when provided together with ownerGroup', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: '' })) {
        return;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: true })) {
        return;
      }
      if (opts.url === `${siteUrl}/_api/Web/SetAccessRequestSiteDescriptionAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ description: 'Hello' })) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl, ownerGroup: true, message: 'Hello', verbose: true } });

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.calledTwice, true);
  });

  it('handles error when updating RequestAccessEmail (PATCH Web)', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: 'john.doe@contoso.com' })) {
        throw {
          error: {
            'odata.error': {
              message: {
                lang: 'en-US',
                value: 'Access is denied.'
              }
            }
          }
        };
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async () => {
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl, email: 'john.doe@contoso.com', verbose: true } }), new CommandError('Access is denied.'));

    assert.strictEqual(patchStub.calledOnce, true);
    assert.strictEqual(postStub.called, false);
  });

  it('handles error when updating default target (POST UseAccessRequestDefault)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web` && JSON.stringify(opts.data) === JSON.stringify({ RequestAccessEmail: '' })) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate` && JSON.stringify(opts.data) === JSON.stringify({ useAccessRequestDefault: true })) {
        throw {
          error: {
            'odata.error': {
              message: {
                lang: 'en-US',
                value: 'Failed to set default.'
              }
            }
          }
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl, ownerGroup: true, verbose: true } }), new CommandError('Failed to set default.'));
  });
});


