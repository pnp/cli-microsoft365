import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
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
import command from './m365group-add.js';

describe(commands.M365GROUP_ADD, () => {

  const groupResponse: any = {
    id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76',
    deletedDateTime: null,
    classification: null,
    createdDateTime: '2018-02-24T18:38:53Z',
    description: 'My awesome group',
    displayName: 'My group',
    groupTypes: ['Unified'],
    mail: 'my_group@contoso.onmicrosoft.com',
    mailEnabled: true,
    mailNickname: 'my_group',
    onPremisesLastSyncDateTime: null,
    onPremisesProvisioningErrors: [],
    onPremisesSecurityIdentifier: null,
    onPremisesSyncEnabled: null,
    preferredDataLocation: null,
    proxyAddresses: ['SMTP:my_group@contoso.onmicrosoft.com'],
    renewedDateTime: '2018-02-24T18:38:53Z',
    resourceBehaviorOptions: [],
    securityEnabled: false,
    visibility: 'Public'
  };

  const fsStats: fs.Stats = {
    isDirectory: () => false,
    isFile: () => false,
    isBlockDevice: () => false,
    isCharacterDevice: () => false,
    isSymbolicLink: () => false,
    isFIFO: () => false,
    isSocket: () => false,
    dev: 0,
    ino: 0,
    mode: 0,
    nlink: 0,
    uid: 0,
    gid: 0,
    rdev: 0,
    size: 0,
    blksize: 0,
    blocks: 0,
    atimeMs: 0,
    mtimeMs: 0,
    ctimeMs: 0,
    birthtimeMs: 0,
    atime: new Date(),
    mtime: new Date(),
    ctime: new Date(),
    birthtime: new Date()
  };

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
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.put,
      request.get,
      fs.readFileSync,
      fs.existsSync,
      fs.lstatSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates Microsoft 365 Group using basic info', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', verbose: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('creates private Microsoft 365 Group using basic info', async () => {
    const privateGroupResponse = { ...groupResponse };
    privateGroupResponse.visibility = 'Private';

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return privateGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', visibility: 'Private' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Private'
    });
    assert(loggerLogSpy.calledOnceWith(privateGroupResponse));
  });

  it('creates Microsoft 365 Group with resourceBehaviorOptions', async () => {
    const response = { ...groupResponse };
    response.resourceBehaviorOptions = ['AllowOnlyMembersToPost', 'HideGroupInOutlook', 'SubscribeNewGroupMembers', 'WelcomeEmailDisabled'];

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', allowMembersToPost: true, hideGroupInOutlook: true, subscribeNewGroupMembers: true, welcomeEmailDisabled: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: ['AllowOnlyMembersToPost', 'HideGroupInOutlook', 'SubscribeNewGroupMembers', 'WelcomeEmailDisabled'],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('creates Microsoft 365 Group with a png logo', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });
    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert.strictEqual(putStub.lastCall.args[0].headers!['content-type'], 'image/png');
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('creates Microsoft 365 Group with a jpg logo', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.jpg' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert.strictEqual(putStub.lastCall.args[0].headers!['content-type'], 'image/jpeg');
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('creates Microsoft 365 Group with a gif logo', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/gif') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.gif' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert.strictEqual(putStub.lastCall.args[0].headers!['content-type'], 'image/gif');
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('handles failure when creating Microsoft 365 Group with a logo and succeeds on tenth call', async () => {
    let amountOfCalls = 1;
    sinon.stub(fs, 'readFileSync').returns('abc');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });
    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        if (amountOfCalls++ < 10) {
          throw 'Invalid request';
        }
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
    assert.strictEqual(putStub.callCount, 10);
  });

  it('handles failure when creating Microsoft 365 Group with a logo (debug)', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'put').rejects(new Error('Invalid request'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' }) }),
      new CommandError('Invalid request'));

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      description: 'My awesome group',
      displayName: 'My group',
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: 'my_group',
      resourceBehaviorOptions: [],
      securityEnabled: false,
      visibility: 'Public'
    });
  });

  it('creates Microsoft 365 Group with specific owner', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data['@odata.id'], 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a');
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('creates Microsoft 365 Group with specific owners', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b',
              userPrincipalName: 'user2@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) });
    assert.deepStrictEqual(postStub.getCall(-2).args[0].data, {
      '@odata.id': 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a'
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      '@odata.id': 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b'
    });
  });

  it('creates Microsoft 365 Group with specific member', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data['@odata.id'], 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a');
    assert(loggerLogSpy.calledOnceWith(groupResponse));
  });

  it('creates Microsoft 365 Group with specific members (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupResponse;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b',
              userPrincipalName: 'user2@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) });
    assert.deepStrictEqual(postStub.getCall(-2).args[0].data, {
      '@odata.id': 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a'
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      '@odata.id': 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b'
    });
  });

  it('fails when an invalid user is specified as owner', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) }),
      new CommandError('Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('fails when an invalid user is specified as owner (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) }),
      new CommandError('Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('fails when an invalid user is specified as member', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) }),
      new CommandError('Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('fails when an invalid user is specified as member (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' }) }),
      new CommandError('Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' }) }),
      new CommandError('Invalid request'));
  });

  it('passes validation when the displayName, description and mailNickname are specified', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when mailNickname contains spaces', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my group' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if one of the owners is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the owner is valid', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with multiple owners, comma-separated', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with multiple owners, comma-separated with an additional space', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if one of the members is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the member is valid', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with multiple members, comma-separated', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with multiple members, comma-separated with an additional space', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if logoPath points to a non-existent file', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'some-image.png' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if logoPath points to a folder', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => fsStats);
    sinon.stub(fsStats, 'isDirectory').callsFake(() => true);

    const actual = commandOptionsSchema.safeParse({ displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'some-folder' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if incorrect visibility is specified.', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'My group',
      description: 'My awesome group',
      mailNickname: 'my_group',
      visibility: "invalid"
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if logoPath points to an existing file', () => {
    // Use the package.json which definitely exists
    const actual = commandOptionsSchema.safeParse({ 
      displayName: 'My group', 
      description: 'My awesome group', 
      mailNickname: 'my_group', 
      logoPath: 'package.json'
    });
    assert.strictEqual(actual.success, true);
  });

  it('supports specifying displayName', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'displayName') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'description') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying mailNickname', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'mailNickname') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying owners', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'owners') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying members', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'members') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying group type', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'visibility') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying logo file path', () => {
    const options = commandInfo.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.long === 'logoPath') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
