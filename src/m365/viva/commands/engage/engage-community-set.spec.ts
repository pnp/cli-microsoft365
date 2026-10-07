import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './engage-community-set.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.ENGAGE_COMMUNITY_SET, () => {
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9';
  const displayName = 'Software Engineers';
  const entraGroupId = '0bed8b86-5026-4a93-ac7d-56750cc099f1';
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when id is specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: communityId, description: 'Community for all devs' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'Software Engineers', description: 'Community for all devs' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when entraGroupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({ entraGroupId: '0bed8b86-5026-4a93-ac7d-56750cc099f1', description: 'Community for all devs' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when newDisplayName, description or privacy is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'Software Engineers' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if newDisplayName is more than 255 characters', () => {
    const actual = commandOptionsSchema.safeParse({
      id: communityId,
      newDisplayName: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries."
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if description is more than 1024 characters', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'Software engineers',
      description: `Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.`
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when invalid privacy option is provided', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'Software engineers',
      privacy: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when entraGroupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ entraGroupId: 'foo', description: 'Community for all devs' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: communityId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('updates info about a community specified by id', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: communityId, newDisplayName: 'Software Engineers', verbose: true }) });
    assert(patchRequestStub.called);
  });

  it('updates info about a community specified by displayName', async () => {
    sinon.stub(vivaEngage, 'getCommunityByDisplayName').resolves({ id: communityId });
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: displayName, description: 'Community for all devs', privacy: 'Public', verbose: true }) });
    assert(patchRequestStub.called);
  });

  it('updates info about a community specified by entraGroupId', async () => {
    sinon.stub(vivaEngage, 'getCommunityByEntraGroupId').resolves({ id: communityId });
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ entraGroupId: entraGroupId, description: 'Community for all devs', privacy: 'Public', verbose: true }) });
    assert(patchRequestStub.called);
  });

  it('handles error when updating Viva Engage community failed', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: communityId, description: 'test' }) }),
      new CommandError('An error has occurred')
    );
  });
});