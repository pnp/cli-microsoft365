import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './people-profilecardproperty-get.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_GET, () => {
  const profileCardPropertyName = 'customAttribute1';

  //#region Mocked responses
  const response = {
    directoryPropertyName: profileCardPropertyName,
    annotations: [
      {
        displayName: 'Cost center',
        localizations: [
          {
            languageTag: 'nl-NL',
            displayName: 'Kostencentrum'
          }
        ]
      }
    ]
  };
  //#endregion

  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PEOPLE_PROFILECARDPROPERTY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when name is invalid', async () => {
    const actual = await command.validate({ options: { name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name is valid', async () => {
    const actual = await command.validate({ options: { name: profileCardPropertyName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name is valid with different capitalization', async () => {
    const actual = await command.validate({ options: { name: 'cUstoMATTriBUtE1' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('gets profile card property information', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('gets profile card property information for text output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    const textOutput = {
      directoryPropertyName: profileCardPropertyName,
      displayName: response.annotations[0].displayName,
      ['displayName ' + response.annotations[0].localizations[0].languageTag]: response.annotations[0].localizations[0].displayName
    };

    await command.action(logger, { options: { name: profileCardPropertyName, output: 'text' } });
    assert(loggerLogSpy.calledOnceWith(textOutput));
  });

  it('uses correct casing for name when incorrect casing is used', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName.toUpperCase() } });
    assert(getStub.called);
  });

  it('handles error when profile card property does not exist', async () => {
    sinon.stub(request, 'get').rejects({
      response: {
        status: 404
      }
    });

    await assert.rejects(command.action(logger, { options: { name: profileCardPropertyName } } as any),
      new CommandError(`Profile card property '${profileCardPropertyName}' does not exist.`));
  });

  it('handles unexpected API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { name: profileCardPropertyName } } as any),
      new CommandError(errorMessage));
  });
});