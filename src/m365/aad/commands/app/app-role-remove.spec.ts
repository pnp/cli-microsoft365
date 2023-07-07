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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-role-remove');

describe(commands.APP_ROLE_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ROLE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --confirm option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role name and --confirm option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'ProductRead',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role id and --confirm option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role claim and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        claim: 'Product.Read',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --confirm option specified (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --confirm option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --confirm option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --confirm option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --confirm option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --confirm option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    });
  });

  it('handles error when multiple apps with the specified appName found and --confirm option is specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            },
            {
              id: 'a39c738c-939e-433b-930d-b02f2931a08b'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }), new CommandError(`Multiple Azure AD application registration with name App-Name found. Please disambiguate using app object IDs: 5b31c38c-2584-42f0-aa47-657fb3a84230, a39c738c-939e-433b-930d-b02f2931a08b`));
  });

  it('handles when multiple roles with the same name are found and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product get",
              "displayName": "ProductRead",
              "id": "9267ab18-8d09-408d-8c94-834662ed16d1",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Get"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }), new CommandError(`Multiple roles with the provided 'name' were found. Please disambiguate using the claims : Product.Read, Product.Get`));
  });

  it('handles when no roles with the specified name are found and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }), new CommandError(`No app role with name 'ProductRead' found.`));
  });

  it('handles when no roles with the specified claim are found and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }), new CommandError(`No app role with claim 'Product.Read' found.`));
  });

  it('handles when no roles with the specified id are found and --confirm option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }), new CommandError(`No app role with id 'c4352a0a-494f-46f9-b843-479855c173a7' found.`));
  });

  it('prompts before removing the specified app role when confirm option not passed', async () => {
    await command.action(logger, { options: { appName: 'App-Name', claim: 'Product.Read' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified app role when confirm option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and the prompt is confirmed (debug)', async () => {

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: false
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and prompt is confirmed', async () => {

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: false
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and prompt is confirmed (debug)', async () => {

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: false
      }
    });
  });

  it('aborts deleting app role when prompt is not confirmed', async () => {
    // represents the aad app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { appName: 'App-Name', claim: 'Product.Read' } });
    assert(patchStub.notCalled);
  });

  it('aborts deleting app role when prompt is not confirmed (debug)', async () => {
    // represents the aad app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } });
    assert(patchStub.notCalled);
  });

  it('handles error when the app specified with appObjectId not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        throw {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-04-20T17:22:30",
              "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
              "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
            }
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'App-Role',
        confirm: true
      }
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'App-Role',
        confirm: true
      }
    }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
  });

  it('handles error when the app specified with appName not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app',
        name: 'App-Role',
        confirm: true
      }
    }), new CommandError(`No Azure AD application registration with name My app found`));
  });

  it('fails validation if appId and appObjectId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role name and id is specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation role name and claim is specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", claim: "Product.Read" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role id and claim is specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: "Product.Read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither role name, id or claim specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified role id is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '77355bee' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,name', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,claim', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,id', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,name', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,claim', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,id', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,name', async () => {
    const actual = await command.validate({ options: { appName: 'My App', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,claim', async () => {
    const actual = await command.validate({ options: { appName: 'My App', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,id', async () => {
    const actual = await command.validate({ options: { appName: 'My App', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
