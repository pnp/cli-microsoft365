import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./threatassessment-list');

describe(commands.THREATASSESMENT_LIST, () => {
  //#region Mocked Responses
  const validType: string = '';
  const threatAssesmentResponse: any = {};
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THREATASSESMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'contentType', 'category']);
  });

  it('retrieves threat assesments', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(threatAssesmentResponse));
  });

  it('retrieves threat assesments with type', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/v1.0/informationProtection/threatAssessmentRequests`)) {
        return threatAssesmentResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, type: validType } });
    assert(loggerLogSpy.calledWith(threatAssesmentResponse));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The threat assesments could not be retrieved'
      }
    };
    sinon.stub(request, 'get').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError('The threat assesments could not be retrieved'));
  });
});
