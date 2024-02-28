import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { session } from '../../../../utils/session.js';
import { accessToken } from '../../../../utils/accessToken.js';
import fs from 'fs';
import commands from '../../commands.js';
import command from './threatassessment-add.js';

describe(commands.THREATASSESSMENT_ADD, () => {

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    (command as any).items = [];
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('VGhpcyBpcyBhIHRlc3QgZmlsZQ==');
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      fs.existsSync,
      fs.readFileSync,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.THREATASSESSMENT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests`) {
        return fileThreatAssessmentRequest;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'file', expectedAssessment: 'block', category: 'malware', path: 'C:\\Temp\\DummyFile.txt', verbose: true } });
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
    sinonUtil.restore(fs.existsSync);
    sinon.stub(fs, 'existsSync').returns(false);
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