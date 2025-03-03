import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './roledefinition-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.ROLEDEFINITION_ADD, () => {
  const roleDefinitionResponse = {
    "id": "e1ede50a-487c-49b3-a43e-cda270d3341f",
    "description": null,
    "displayName": "Custom Role",
    "isBuiltIn": false,
    "isEnabled": true,
    "resourceScopes": [
      "/"
    ],
    "templateId": "e1ede50a-487c-49b3-a43e-cda270d3341f",
    "version": null,
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "microsoft.directory/groups.unified/create",
          "microsoft.directory/groups.unified/delete"
        ],
        "condition": null
      }
    ]
  };

  const roleDefinitionWithDetailsResponse = {
    "id": "abcde50a-487c-49b3-a43e-cda270d3341f",
    "description": "Allows creating and deleting unified groups",
    "displayName": "Custom Role",
    "isBuiltIn": false,
    "isEnabled": false,
    "resourceScopes": [
      "/"
    ],
    "templateId": "abcnpm instade50a-487c-49b3-a43e-cda270d3341f",
    "version": "1",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "microsoft.directory/groups.unified/create",
          "microsoft.directory/groups.unified/delete"
        ],
        "condition": null
      }
    ]
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEDEFINITION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if displayName is not provided', () => {
    const actual = commandOptionsSchema.safeParse({ allowedResourceActions: "microsoft.directory/groups.unified/create" });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if allowedResourceActions is not provided', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: "Custom Role" });
    assert.notStrictEqual(actual.success, true);
  });

  it('creates a custom role definition with a specific display name and resource actions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions') {
        return roleDefinitionResponse;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse(
      {
        displayName: 'Custom Role',
        allowedResourceActions: "microsoft.directory/groups.unified/create,microsoft.directory/groups.unified/delete"
      });
    await command.action(logger, { options: parsedSchema.data });
    assert(loggerLogSpy.calledOnceWithExactly(roleDefinitionResponse));
  });

  it('creates a custom role definition with a specific display name, description, version and resource actions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions') {
        return roleDefinitionWithDetailsResponse;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      displayName: 'Custom Role',
      description: 'Allows creating and deleting unified groups',
      allowedResourceActions: "microsoft.directory/groups.unified/create,microsoft.directory/groups.unified/delete",
      enabled: false,
      version: "1",
      verbose: true
    });
    await command.action(logger, {
      options: parsedSchema.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(roleDefinitionWithDetailsResponse));
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

    const parsedSchema = commandOptionsSchema.safeParse({
      displayName: 'Custom Role',
      allowedResourceActions: "microsoft.directory/groups.unified/create"
    });
    await assert.rejects(command.action(logger, {
      options: parsedSchema.data
    }), new CommandError('Invalid request'));
  });
});