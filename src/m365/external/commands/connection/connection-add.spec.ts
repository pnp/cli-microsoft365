import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
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
import command from './connection-add.js';

describe(commands.CONNECTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const externalConnectionAddResponse: ExternalConnectors.ExternalConnection = {
    configuration: {
      authorizedAppIds: []
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLI',
    name: 'Test Connection for CLI'
  };

  const externalConnectionAddResponseWithAppIDs: ExternalConnectors.ExternalConnection = {
    configuration: {
      'authorizedAppIds': [
        '00000000-0000-0000-0000-000000000000',
        '00000000-0000-0000-0000-000000000001',
        '00000000-0000-0000-0000-000000000002'
      ]
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLI',
    name: 'Test Connection for CLI'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONNECTION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('adds an external connection', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return externalConnectionAddResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything'
    };
    await command.action(logger, { options: options } as any);
    assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponse);
  });

  it('adds an external connection with authorized app id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return externalConnectionAddResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything',
      authorizedAppIds: '00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'
    };
    await command.action(logger, { options: options } as any);
    assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponseWithAppIDs);
  });

  it('adds an external connection with authorised app IDs', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return externalConnectionAddResponseWithAppIDs;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything',
      authorizedAppIds: '00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'
    };
    await command.action(logger, { options: options } as any);
    assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponseWithAppIDs);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw {
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } } as any),
      new CommandError(`An error has occurred`));
  });

  it('fails validation if id is less than 3 characters', async () => {
    const actual = await command.validate({
      options: {
        id: 'T',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is more than 32 characters', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestConnectionForCLIXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is not alphanumeric', async () => {
    const actual = await command.validate({
      options: {
        id: 'Test_Connection!',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id starts with Microsoft', async () => {
    const actual = await command.validate({
      options: {
        id: 'MicrosoftTestConnectionForCLI',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is SharePoint', async () => {
    const actual = await command.validate({
      options: {
        id: 'SharePoint',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for a valid id', async () => {
    const actual = await command.validate({
      options: {
        id: 'myapp',
        name: 'Test Connection for CLI',
        description: 'Test connection'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
