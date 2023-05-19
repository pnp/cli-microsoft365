import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
import * as fs from 'fs';
const command: Command = require('./threatassessment-add');

describe(commands.THREATASSESSMENT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    (command as any).items = [];
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.post,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THREATASSESSMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a mail assessment request', async () => {
    const mailThreatAssessmentRequest = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#informationProtection/threatAssessmentRequests/$entity',
      '@odata.type': '#microsoft.graph.mailAssessmentRequest',
      'id': '49c5ef5b-1f65-444a-e6b9-08d772ea2059',
      'createdDateTime': '2019-11-27T03:30:18.6890937Z',
      'contentType': 'mail',
      'expectedAssessment': 'block',
      'category': 'spam',
      'status': 'pending',
      'requestSource': 'administrator',
      'recipientEmail': 'john@doe.com',
      'destinationRoutingReason': 'notJunk',
      'messageUri': 'https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=',
      'createdBy': {
        'user': {
          'id': 'c52ce8db-3e4b-4181-93c4-7d6b6bffaf60',
          'displayName': 'John Doe'
        }
      }
    };

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`) {
        return mailThreatAssessmentRequest;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'mail', expectedAssessment: 'block', category: 'spam', recipientEmail: 'john@doe.com', messageUri: 'https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=', verbose: true } });
    assert.strictEqual(postStub.lastCall.args[0].data['@odata.type'], '#microsoft.graph.mailAssessmentRequest');
    assert(loggerLogSpy.calledWith(mailThreatAssessmentRequest));
  });

  it('creates an email file assessment request', async () => {
    const emailFileThreatAssessmentRequest = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#informationProtection/threatAssessmentRequests/$entity',
      '@odata.type': '#microsoft.graph.emailFileAssessmentRequest',
      'id': '49c5ef5b-1f65-444a-e6b9-08d772ea2059',
      'createdDateTime': '2019-11-27T03:30:18.6890937Z',
      'contentType': 'mail',
      'expectedAssessment': 'block',
      'category': 'malware',
      'status': 'completed',
      'requestSource': 'administrator',
      'recipientEmail': 'john@doe.com',
      'destinationRoutingReason': 'notJunk',
      'contentData': 'UmVjZWl2ZWQ6IGZyb20gTVcyUFIwME1CMDMxNC5uYW1wcmQwMC',
      'createdBy': {
        'user': {
          'id': 'c52ce8db-3e4b-4181-93c4-7d6b6bffaf60',
          'displayName': 'John Doe'
        }
      }
    };

    sinon.stub(fs, 'readFileSync').callsFake(() => 'SGVsbG8gdGhlcmUgY2xpIGZvciBNaWNyb3NvZnQgMzY1IHVzZXJzIQ==');

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`) {
        return emailFileThreatAssessmentRequest;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'emailFile', expectedAssessment: 'block', category: 'malware', recipientEmail: 'john@doe.com', path: 'C:\\Temp\\DummyFile.txt', verbose: true } });
    assert.strictEqual(postStub.lastCall.args[0].data['@odata.type'], '#microsoft.graph.emailFileAssessmentRequest');
    assert(loggerLogSpy.calledWith(emailFileThreatAssessmentRequest));
  });

  it('creates a file assessment request', async () => {
    const fileThreatAssessmentRequest = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#informationProtection/threatAssessmentRequests/$entity',
      '@odata.type': '#microsoft.graph.fileAssessmentRequest',
      'id': '18406a56-7209-4720-a250-08d772fccdaa',
      'createdDateTime': '2019-11-27T05:44:00.4051536Z',
      'contentType': 'file',
      'expectedAssessment': 'block',
      'category': 'malware',
      'status': 'completed',
      'requestSource': 'administrator',
      'fileName': 'illegalfile.txt',
      'contentData': '',
      'createdBy': {
        'user': {
          'id': 'c52ce8db-3e4b-4181-93c4-7d6b6bffaf60',
          'displayName': 'John Doe'
        }
      }
    };

    sinon.stub(fs, 'readFileSync').callsFake(() => 'SGVsbG8gdGhlcmUgY2xpIGZvciBNaWNyb3NvZnQgMzY1IHVzZXJzIQ==');

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`) {
        return fileThreatAssessmentRequest;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'file', expectedAssessment: 'block', category: 'malware', fileName: 'illegalfile.txt', path: 'C:\\Temp\\DummyFile.txt', verbose: true } });
    assert.strictEqual(postStub.lastCall.args[0].data['@odata.type'], '#microsoft.graph.fileAssessmentRequest');
    assert(loggerLogSpy.calledWith(fileThreatAssessmentRequest));
  });

  it('creates an url assessment request', async () => {
    const urlThreatAssessmentRequest = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#informationProtection/threatAssessmentRequests/$entity',
      '@odata.type': '#microsoft.graph.urlAssessmentRequest',
      'id': '8d87d2b2-ca4d-422c-f8df-08d774a5c9ac',
      'createdDateTime': '2019-11-29T08:26:09.8196703Z',
      'contentType': 'url',
      'expectedAssessment': 'block',
      'category': 'phishing',
      'status': 'pending',
      'requestSource': 'administrator',
      'url': 'http://test.com',
      'createdBy': {
        'user': {
          'id': 'c52ce8db-3e4b-4181-93c4-7d6b6bffaf60',
          'displayName': 'John Doe'
        }
      }
    };

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`) {
        return urlThreatAssessmentRequest;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'url', expectedAssessment: 'block', category: 'phishing', url: 'https://phisingurl.be', verbose: true } });
    assert.strictEqual(postStub.lastCall.args[0].data['@odata.type'], '#microsoft.graph.urlAssessmentRequest');
    assert(loggerLogSpy.calledWith(urlThreatAssessmentRequest));
  });

  it('passes validation if all options are passed propertly', async () => {
    const actual = await command.validate({ options: { type: 'url', expectedAssessment: 'block', category: 'spam', url: 'https://pnp.github.io/cli-microsoft365/' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if type is not a valid type', async () => {
    const actual = await command.validate({ options: { type: 'invalid', expectedAssessment: 'block', category: 'spam', url: 'https://pnp.github.io/cli-microsoft365/' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if expectedAssessment is not a valid expectedAssessment', async () => {
    const actual = await command.validate({ options: { type: 'url', expectedAssessment: 'invalid', category: 'spam', url: 'https://pnp.github.io/cli-microsoft365/' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if category is not a valid category', async () => {
    const actual = await command.validate({ options: { type: 'url', expectedAssessment: 'block', category: 'invalid', url: 'https://pnp.github.io/cli-microsoft365/' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if path is specified and file does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { type: 'file', expectedAssessment: 'block', category: 'malware', path: 'C:\\Path\\That\\Does\\Not\\Exist' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws an error when we execute the command using application permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: { type: 'url', expectedAssessment: 'block', category: 'spam', url: 'https://pnp.github.io/cli-microsoft365/' } }),
      new CommandError('This command currently does not support app only permissions.'));
  });
});