import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./solution-publisher-add');

describe(commands.SOLUTION_PUBLISHER_ADD, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const validName = "PublisherName";
  const validDisplayName = "Publisher Name";
  const validPrefix = "c6rx";
  const validChoiceValuePrefix = 10000;
  //#endregion

  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_PUBLISHER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if choiceValuePrefix is not a number', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        name: validName,
        displayName: validDisplayName,
        prefix: validPrefix,
        choiceValuePrefix: 'Not A Number'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if choiceValuePrefix is more then the upper bound', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        name: validName,
        displayName: validDisplayName,
        prefix: validPrefix,
        choiceValuePrefix: 100000
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if choiceValuePrefix is less then the lower bound', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        name: validName,
        displayName: validDisplayName,
        prefix: validPrefix,
        choiceValuePrefix: 9999
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name is not a valid value', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        name: '9_PublisherName',
        displayName: validDisplayName,
        prefix: validPrefix,
        choiceValuePrefix: validChoiceValuePrefix
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if prefix is not a valid value', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        name: validName,
        displayName: validDisplayName,
        prefix: 'mscrmtes',
        choiceValuePrefix: validChoiceValuePrefix
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified', async () => {
    const actual = await command.validate({ options: { environment: validEnvironment, name: validName, displayName: validDisplayName, prefix: validPrefix, choiceValuePrefix: validChoiceValuePrefix } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds a specific publisher with the required parameters', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, name: validName, displayName: validDisplayName, prefix: validPrefix, choiceValuePrefix: validChoiceValuePrefix } });
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, InvalidOperationException',
                message: {
                  value: `Resource '' does not exist or one of its queried reference-property objects are not present`
                }
              }
            }
          };
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { debug: false, environment: validEnvironment, name: validName, displayName: validDisplayName, prefix: validPrefix, choiceValuePrefix: validChoiceValuePrefix } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});