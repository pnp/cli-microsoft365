import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as fs from 'fs';
import { odata } from '../../../../utils/odata';
import request from '../../../../request';
import { PassThrough } from 'stream';
const command: Command = require('./message-export');

describe(commands.MESSAGE_EXPORT, () => {
  const userId = '11f43044-095e-456a-b339-7e1901b0c3ae';
  const userPrincipalName = 'john@contoso.com';
  const teamId = '75619fc7-5dce-412b-82ee-f76988d3efaa';
  const fromDateTime = '2023-04-01T00:00:00Z';
  const toDateTime = '2023-04-30T23:59:59Z';
  const folderPath = 'C:\\Temp';

  const teamMessageResponse = [{ 'id': '1683611790633', 'replyToId': null, 'etag': '1683612581046', 'messageType': 'message', 'createdDateTime': '2023-05-09T05:56:30.633Z', 'lastModifiedDateTime': '2023-05-09T06:09:41.046Z', 'lastEditedDateTime': '2023-05-09T06:09:40.914Z', 'deletedDateTime': null, 'subject': null, 'summary': null, 'chatId': null, 'importance': 'normal', 'locale': 'en-us', 'webUrl': 'https://teams.microsoft.com/l/message/19%3A0ade024bcdec4f4fbb03dfa0e5afc3ee%40thread.tacv2/1683611790633?groupId=80fc64e4-f5e1-4dc1-b34c-3198375bd9b2&tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&createdTime=1683611790633&parentMessageId=1683611790633', 'policyViolation': null, 'eventDetail': null, 'from': { 'application': null, 'device': null, 'user': { 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'displayName': 'John Doe', 'userIdentityType': 'aadUser', 'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d' } }, 'body': { 'contentType': 'text', 'content': 'Hello! <attachment id=\'0e7fbdea-21ec-4525-aa30-c94e4e24c90b\'></attachment><attachment id=\'161D73CB-20D7-47FC-95CA-2E60A5A44D8D\'></attachment>' }, 'channelIdentity': { 'teamId': '80fc64e4-f5e1-4dc1-b34c-3198375bd9b2', 'channelId': '19:0ade024bcdec4f4fbb03dfa0e5afc3ee@thread.tacv2' }, 'attachments': [{ 'id': '161D73CB-20D7-47FC-95CA-2E60A5A44D8D', 'contentType': 'reference', 'contentUrl': 'https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/Documents/File1.pdf', 'content': null, 'name': 'File1.pdf', 'thumbnailUrl': null, 'teamsAppId': null }], 'mentions': [], 'reactions': [] }];
  const userMessageResponse = [{ 'id': '1668781541156', 'replyToId': null, 'etag': '1668781541156', 'messageType': 'message', 'createdDateTime': '2022-11-18T14:25:41.156Z', 'lastModifiedDateTime': '2022-11-18T14:25:41.156Z', 'lastEditedDateTime': null, 'deletedDateTime': null, 'subject': null, 'summary': null, 'chatId': '19:meeting_YmQwYbNzZgUtNmYxMC00YzFjLWE1MDctY2QwNmVkMGU4N2Ex@thread.v2', 'importance': 'normal', 'locale': 'en-us', 'webUrl': null, 'channelIdentity': null, 'policyViolation': null, 'eventDetail': null, 'from': { 'application': null, 'device': null, 'user': { 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'displayName': 'John Doe', 'userIdentityType': 'aadUser', 'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d' } }, 'body': { 'contentType': 'text', 'content': 'CLI For Microsoft 365 Rocks!' }, 'attachments': [], 'mentions': [], 'reactions': [] }];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.get,
      fs.createWriteStream,
      fs.existsSync,
      fs.mkdirSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_EXPORT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves messages for a specific team without downloading attachments', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/getAllMessages`) {
        return teamMessageResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: teamId } });
    assert(loggerLogSpy.calledWith(teamMessageResponse));
  });

  it('retrieves messages for a specific team with attachments', async () => {
    const mockResponse = `CLI For Microsoft 365 Rocks!`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/getAllMessages`) {
        return teamMessageResponse;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/_api/web/getFileByServerRelativePath(decodedUrl=@decodedUrl)/$value?@decodedUrl='/personal/john_contoso_onmicrosoft_com/Documents/File1.pdf'`) {
        return { data: responseStream };
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'mkdirSync').returns(`${folderPath}\\${teamMessageResponse[0].id}`);

    const writeStream = new PassThrough();
    sinon.stub(fs, 'createWriteStream').returns(writeStream as any);
    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    await command.action(logger, { options: { teamId: teamId, withAttachments: true, folderPath: folderPath, verbose: true } });
    assert(loggerLogSpy.calledWith(teamMessageResponse));
  });

  it('retrieves messages for a specific user by id without downloading attachments', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userId}/chats/getAllMessages`) {
        return userMessageResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });
    assert(loggerLogSpy.calledWith(userMessageResponse));
  });

  it('retrieves messages for a specific user by user name without downloading attachments and filtering on dates', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/chats/getAllMessages?$filter=createdDateTime ge ${fromDateTime} and createdDateTime lt ${toDateTime}`) {
        return userMessageResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userPrincipalName, fromDateTime: fromDateTime, toDateTime: toDateTime } });
    assert(loggerLogSpy.calledWith(userMessageResponse));
  });

  it('retrieves messages for a specific user by user name and skips downloading attachments when there are no attachments available', async () => {
    const userMessageResponseWithoutAttachments = [{ 'id': '1668781541156', 'replyToId': null, 'etag': '1668781541156', 'messageType': 'message', 'createdDateTime': '2022-11-18T14:25:41.156Z', 'lastModifiedDateTime': '2022-11-18T14:25:41.156Z', 'lastEditedDateTime': null, 'deletedDateTime': null, 'subject': null, 'summary': null, 'chatId': '19:meeting_YmQwYbNzZgUtNmYxMC00YzFjLWE1MDctY2QwNmVkMGU4N2Ex@thread.v2', 'importance': 'normal', 'locale': 'en-us', 'webUrl': null, 'channelIdentity': null, 'policyViolation': null, 'eventDetail': null, 'from': { 'application': null, 'device': null, 'user': { 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'displayName': 'John Doe', 'userIdentityType': 'aadUser', 'tenantId': 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d' } }, 'body': { 'contentType': 'text', 'content': 'CLI For Microsoft 365 Rocks!' }, 'mentions': [], 'reactions': [] }];

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/chats/getAllMessages?$filter=createdDateTime ge ${fromDateTime} and createdDateTime lt ${toDateTime}`) {
        return userMessageResponseWithoutAttachments;
      }
      throw 'Invalid request';
    });

    const getStub = sinon.stub(request, 'get').resolves();

    await command.action(logger, { options: { userName: userPrincipalName, fromDateTime: fromDateTime, toDateTime: toDateTime, withAttachments: true, folderPath: folderPath } });
    assert(getStub.notCalled);
  });

  it('handles error when request to retrieve data fails', async () => {
    const errorMessage = {
      'error': {
        'code': 'PaymentRequired',
        'message': 'Evaluation mode capacity has been exceeded. Use a valid billing model. Visit https://docs.microsoft.com/en-us/graph/teams-licenses for more details.',
        'innerError': {
          'date': '2023-05-21T20:06:07',
          'request-id': '2e9937c1-307c-46f3-a656-76deb2d8d77f',
          'client-request-id': '2e9937c1-307c-46f3-a656-76deb2d8d77f'
        }
      }
    };
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/chats/getAllMessages`) {
        throw errorMessage;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { userName: userPrincipalName } }), new CommandError('Evaluation mode capacity has been exceeded. Use a valid billing model. Visit https://docs.microsoft.com/en-us/graph/teams-licenses for more details.'));
  });

  it('handles error when attachment cannot be saved properly', async () => {
    const mockResponse = `CLI For Microsoft 365 Rocks!`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/getAllMessages`) {
        return teamMessageResponse;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/_api/web/getFileByServerRelativePath(decodedUrl=@decodedUrl)/$value?@decodedUrl='/personal/john_contoso_onmicrosoft_com/Documents/File1.pdf'`) {
        return { data: responseStream };
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'mkdirSync').returns(`${folderPath}\\${teamMessageResponse[0].id}`);

    const writeStream = new PassThrough();
    sinon.stub(fs, 'createWriteStream').returns(writeStream as any);
    setTimeout(() => {
      writeStream.emit('error', 'ENOENT: no such file or directory');
    }, 5);

    await assert.rejects(command.action(logger, { options: { teamId: teamId, withAttachments: true, folderPath: folderPath, verbose: true } }), new CommandError('ENOENT: no such file or directory'));
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid userPrincipalName', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userName: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, teamId: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if fromDateTime is not a valid ISO DateTime', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, fromDateTime: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if toDateTime is not a valid ISO DateTime', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, toDateTime: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if folderPath does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

  it('passes validation if folderPath exists and userId is a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, withAttachments: false } }, commandInfo);
    assert.strictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

  it('passes validation if folderPath exists, teamId is a valid GUID and both dates are valid ISO dates', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { folderPath: folderPath, teamId: teamId, fromDateTime: fromDateTime, toDateTime: toDateTime, withAttachments: false } }, commandInfo);
    assert.strictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

  it('passes validation if userName is a valid userPrincipalName', async () => {
    const actual = await command.validate({ options: { userName: userPrincipalName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
