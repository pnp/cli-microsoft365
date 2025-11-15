import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './rolepermission-list.js';

describe(commands.ROLEDEFINITION_LIST, () => {
  const resourceNamespace = 'microsoft.directory';

  const resourceActions = [
    {
      "actionVerb": null,
      "description": "Create and delete access reviews, and read and update all properties of access reviews in Microsoft Entra ID",
      "id": "microsoft.directory-accessReviews-allProperties-allTasks",
      "isPrivileged": false,
      "name": "microsoft.directory/accessReviews/allProperties/allTasks",
      "resourceScopeId": null
    },
    {
      "actionVerb": "GET",
      "description": "Read all properties of access reviews",
      "id": "microsoft.directory-accessReviews-allProperties-read-get",
      "isPrivileged": false,
      "name": "microsoft.directory/accessReviews/allProperties/read",
      "resourceScopeId": null
    },
    {
      "actionVerb": null,
      "description": "Create and delete groups, and read and update all properties",
      "id": "microsoft.directory-groups-allProperties-allTasks",
      "isPrivileged": true,
      "name": "microsoft.directory/groups/allProperties/allTasks",
      "resourceScopeId": null
    },
    {
      "actionVerb": null,
      "description": "Create and delete OAuth 2.0 permission grants, and read and update all properties",
      "id": "microsoft.directory-oAuth2PermissionGrants-allProperties-allTasks",
      "isPrivileged": true,
      "name": "microsoft.directory/oAuth2PermissionGrants/allProperties/allTasks",
      "resourceScopeId": null
    }
  ];

  const filteredResourceActions = [
    {
      "actionVerb": null,
      "description": "Create and delete groups, and read and update all properties",
      "id": "microsoft.directory-groups-allProperties-allTasks",
      "isPrivileged": true,
      "name": "microsoft.directory/groups/allProperties/allTasks",
      "resourceScopeId": null
    },
    {
      "actionVerb": null,
      "description": "Create and delete OAuth 2.0 permission grants, and read and update all properties",
      "id": "microsoft.directory-oAuth2PermissionGrants-allProperties-allTasks",
      "isPrivileged": true,
      "name": "microsoft.directory/oAuth2PermissionGrants/allProperties/allTasks",
      "resourceScopeId": null
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEPERMISSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'actionVerb', 'isPrivileged']);
  });

  it(`should get a list of Entra ID role permissions from a resource namespace`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/directory/resourceNamespaces/${resourceNamespace}/resourceActions`) {
        return {
          value: resourceActions
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceNamespace: resourceNamespace,
      verbose: true
    });
    await command.action(logger, {
      options: parsedSchema.data!
    });

    assert(loggerLogSpy.calledWith(resourceActions));
  });

  it(`should get a list of privileged Entra ID role permissions`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/directory/resourceNamespaces/${resourceNamespace}/resourceActions?$filter=isPrivileged eq true`) {
        return {
          value: filteredResourceActions
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceNamespace: resourceNamespace,
      privileged: true,
      verbose: true
    });
    await command.action(logger, {
      options: parsedSchema.data!
    });

    assert(loggerLogSpy.calledWith(filteredResourceActions));
  });

  it('handles error when retrieving role permissions failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/directory/resourceNamespaces/${resourceNamespace}/resourceActions`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      resourceNamespace: resourceNamespace,
      verbose: true
    });
    await assert.rejects(
      command.action(logger, { options: parsedSchema.data! }),
      new CommandError('An error has occurred')
    );
  });
});