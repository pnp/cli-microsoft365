import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./threatassessment-get');

describe(commands.THREATASSESSMENT_GET, () => {
  const threatAssessmentId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';
  const threatAssessmentGetResponse = {
    'id': '8aaba0ac-ec4d-4e62-5774-08db16c68731',
    'createdDateTime': '2023-02-25T00:23:33.0550644Z',
    'contentType': 'mail',
    'expectedAssessment': 'block',
    'category': 'spam',
    'status': 'pending',
    'requestSource': 'administrator',
    'recipientEmail': 'john@contoso.com',
    'destinationRoutingReason': 'notJunk',
    'messageUri': 'https://graph.microsoft.com/v1.0/users/john@contoso.com/messages/AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAEMAABiOC8xvYmdT6G2E_hLMK5kAALHNaMuAAA=',
    'createdBy': {
      'user': {
        'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
        'displayName': 'John Doe'
      }
    }
  };

  const threatAssessmentGetResponseIncludingResults = {
    ...threatAssessmentGetResponse,
    'results': [
      {
        'id': 'a5455871-18d1-44d8-0866-08db16c68b85',
        'createdDateTime': '2023-02-25T00:23:40.28Z',
        'resultType': 'checkPolicy',
        'message': 'No policy was hit.'
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THREATASSESSMENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered and includeResults is specified', async () => {
    const actual = await command.validate({ options: { id: threatAssessmentId, includeResults: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves threat assessment by specified id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests/${threatAssessmentId}`) {
        return threatAssessmentGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: threatAssessmentId, verbose: true } });
    assert(loggerLogSpy.calledWith(threatAssessmentGetResponse));
  });

  it('retrieves threat assessment by specified id including results', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests/${threatAssessmentId}?$expand=results`) {
        return threatAssessmentGetResponseIncludingResults;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: threatAssessmentId, includeResults: true, verbose: true } });
    assert(loggerLogSpy.calledWith(threatAssessmentGetResponseIncludingResults));
  });

  it('handles error when threat assessment by specified id is not found', async () => {
    const error = {
      'error': {
        'code': 'ResourceNotFound',
        'message': 'The requested resource does not exist.',
        'innerError': {
          'date': '2023-02-25T16:13:25',
          'request-id': 'a9e23bc8-0845-4eef-8ba1-e031b098c955',
          'client-request-id': 'a9e23bc8-0845-4eef-8ba1-e031b098c955'
        }
      }
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests/${threatAssessmentId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: threatAssessmentId } }), new CommandError(error.error.message));
  });
});