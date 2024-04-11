
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-community-add.js';

describe(commands.ENGAGE_COMMUNITY_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;
  const operationLocation = `https://graph.microsoft.com/beta/employeeExperience/engagementAsyncOperations('eyJfdHlwZSI6IkxvbmdSdW5uaW5nT3BlcmF0aW9uIiwiaWQiOiI4ZmM2NzEyZS0wMWY4LTQxN2YtYWNmMS1iZTJiYmMxY2FjNGQiLCJvcGVyYXRpb24iOiJDcmVhdGVDb21tdW5pdHkifQ')`;

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      global.setTimeout
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ENGAGE_COMMUNITY_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if \'displayName\' is more than 255 characters', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries.",
        description: "A community for all software engineer",
        privacy: "public"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if \'description\' is more than 1024 characters', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum. There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.",
        privacy: "public"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid privacy option is provided', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "invalid"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid adminEntraId is provided', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "private",
        adminEntraIds: "invalid"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid adminEntraUserName is provided', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "private",
        adminEntraUserNames: "invalid"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both adminEntraIds and adminEntraUserNames are specified', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer.",
        privacy: "public",
        adminEntraIds: "50674d84-6bf1-470b-89b5-d55ce0a5a720",
        adminEntraUserNames: "john.doe@contoso.onmicrosoft.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid options are provided with adminEntraIds', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer.",
        privacy: "public",
        adminEntraIds: "50674d84-6bf1-470b-89b5-d55ce0a5a720"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid options are provided with adminEntraUserNames', async () => {
    const actual = await command.validate({
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer.",
        privacy: "public",
        adminEntraUserNames: "john.doe@contoso.onmicrosoft.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('creates a community without waiting for provisioning to complete', async () => {
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/communities`) {
        return {
          headers: {
            location: operationLocation
          }
        };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { displayName: "Software engineers", description: "A community for all software engineer", privacy: "public", verbose: true } } as any);
    assert(loggerLogSpy.calledWith(operationLocation));
  });

  it('creates a community with adminEntraIds and waits for provisioning to complete', async () => {
    let waitsForCompletion = false;
    let i = 0;

    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/communities`) {
        return {
          headers: {
            location: operationLocation
          }
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === operationLocation) {
        if (i++ < 2) {
          return {
            status: 'inprogress'
          };
        }

        waitsForCompletion = true;
        return {
          status: 'succeeded'
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await command.action(logger, {
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "public",
        adminEntraIds: "50674d84-6bf1-470b-89b5-d55ce0a5a720",
        wait: true,
        verbose: true
      }
    } as any);
    assert.strictEqual(waitsForCompletion, true);
  });

  it('creates a community with adminEntraUserNames and waits for provisioning to complete', async () => {
    let waitsForCompletion = false;
    let i = 0;

    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/communities`) {
        return {
          headers: {
            location: operationLocation
          }
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === operationLocation) {
        if (i++ < 2) {
          return {
            status: 'inprogress'
          };
        }

        waitsForCompletion = true;
        return {
          status: 'succeeded'
        };
      }

      if (opts.url.startsWith(`https://graph.microsoft.com/v1.0/users/`) && opts.url.endsWith(`?$select=id`)) {
        return { id: '50674d84-6bf1-470b-89b5-d55ce0a5a720' };
      }

      throw 'Invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await command.action(logger, {
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "public",
        adminEntraUserNames: "john.doe@consoto.onmicrosoft.com",
        wait: true,
        debug: true
      }
    } as any);
    assert.strictEqual(waitsForCompletion, true);
  });

  it('handles error when waiting for provisioning to complete fails', async () => {
    let i = 0;

    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/communities`) {
        return {
          headers: {
            location: operationLocation
          }
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === operationLocation) {
        if (i++ < 2) {
          return {
            status: 'inprogress'
          };
        }

        return {
          status: 'failed',
          error: {
            message: 'An error has occurred'
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await assert.rejects(command.action(logger, {
      options: {
        displayName: "Software engineers",
        description: "A community for all software engineer",
        privacy: "public",
        wait: true
      }
    } as any), new CommandError('Community creation failed: An error has occurred'));
  });
});