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
import { formatting } from '../../../../utils/formatting';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
const command: Command = require('./owner-list');

describe(commands.OWNER_LIST, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';
  const flowName = '1c6ee23a-a835-44bc-a4f5-462b658efc12';
  const requestUrl = `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/permissions?api-version=2016-11-01`;
  const requestUrlAdmin = `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/permissions?api-version=2016-11-01`;
  const ownerResponseJson = [{ 'name': '8323f7fe-e8a4-46c4-b5ea-f4864887d160', 'id': '/providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/d2642355-a6b8-4662-a418-ce3741584031/permissions/8323f7fe-e8a4-46c4-b5ea-f4864887d160', 'type': '/providers/Microsoft.ProcessSimple/environments/flows/permissions', 'properties': { 'roleName': 'CanEdit', 'permissionType': 'Principal', 'principal': { 'id': '8323f7fe-e8a4-46c4-b5ea-f4864887d160', 'type': 'User' } } }, { 'name': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'id': '/providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/d2642355-a6b8-4662-a418-ce3741584031/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'type': '/providers/Microsoft.ProcessSimple/environments/flows/permissions', 'properties': { 'roleName': 'Owner', 'permissionType': 'Principal', 'principal': { 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'type': 'User' } } }];
  const ownerResponse = { value: ownerResponseJson };
  const ownerResponseText = [{ 'roleName': 'CanEdit', 'id': '8323f7fe-e8a4-46c4-b5ea-f4864887d160', 'type': 'User' }, { 'roleName': 'Owner', 'id': 'fe36f75e-c103-410b-a18a-2bf6df06ac3a', 'type': 'User' }];

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
    commandInfo = Cli.getCommandInfo(command);
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
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OWNER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['roleName', 'id', 'type']);
  });

  it('retrieves owners from a specific flow with output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return ownerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, output: 'json' } });
    assert(loggerLogSpy.calledWith(ownerResponseJson));
  });

  it('retrieves owners from a specific flow with output text as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrlAdmin) {
        return ownerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, name: flowName, asAdmin: true, output: 'text' } });
    assert(loggerLogSpy.calledWith(ownerResponseText));
  });

  it('throws error when no environment found', async () => {
    const error = {
      'error': {
        'code': 'EnvironmentAccessDenied',
        'message': `Access to the environment '${environmentName}' is denied.`
      }
    };
    sinon.stub(request, 'get').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, name: flowName } } as any),
      new CommandError(error.error.message));
  });

  it('throws error when Flow not found', async () => {
    const error = {
      'error': {
        'code': 'FlowNotFound',
        'message': `Could not find flow '${flowName}'.`
      }
    };
    sinon.stub(request, 'get').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, name: flowName } } as any),
      new CommandError(error.error.message));
  });

  it('fails validation if flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if flowName a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, name: flowName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});