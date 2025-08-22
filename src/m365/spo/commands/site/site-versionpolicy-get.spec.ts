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
import { z } from 'zod';
import commands from '../../commands.js';
import command from './site-versionpolicy-get.js';

describe(commands.SITE_VERSIONPOLICY_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  const validSiteUrl = "https://contoso.sharepoint.com";

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
    assert.strictEqual(command.name, commands.SITE_VERSIONPOLICY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if site URL is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if valid site URL is specified', async () => {
    const actual = await command.validate({ options: { siteUrl: validSiteUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves "age" version policy settings for the specified site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validSiteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`) {
        return {
          VersionPolicies: {
            DefaultTrimMode: 1,
            DefaultExpireAfterDays: 200
          },
          MajorVersionLimit: 100
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: validSiteUrl } });
    assert(loggerLogSpy.calledWith({
      defaultTrimMode: 'age',
      defaultExpireAfterDays: 200,
      majorVersionLimit: 100
    }));
  });

  it('retrieves "automatic" version policy settings for the specified site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validSiteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`) {
        return {
          VersionPolicies: {
            DefaultTrimMode: 2,
            DefaultExpireAfterDays: 30
          },
          MajorVersionLimit: 500
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: validSiteUrl } });
    assert(loggerLogSpy.calledWith({
      defaultTrimMode: 'automatic',
      defaultExpireAfterDays: 30,
      majorVersionLimit: 500
    }));
  });

  it('retrieves "number" version policy settings for the specified site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validSiteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`) {
        return {
          VersionPolicies: {
            DefaultTrimMode: 0,
            DefaultExpireAfterDays: 0
          },
          MajorVersionLimit: 300
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: validSiteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith({
      defaultTrimMode: 'number',
      defaultExpireAfterDays: 0,
      majorVersionLimit: 300
    }));
  });

  it('retrieves "inheritTenant" version policy settings for the specified site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${validSiteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`) {
        return {
          MajorVersionLimit: -1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: validSiteUrl, verbose: true } });
    assert(loggerLogSpy.calledWith({
      defaultTrimMode: 'inheritTenant',
      defaultExpireAfterDays: null,
      majorVersionLimit: -1
    }));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { siteUrl: validSiteUrl } }),
      new CommandError('An error has occurred'));
  });
});