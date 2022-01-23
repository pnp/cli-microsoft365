import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-role-delete');

describe(commands.APP_ROLE_DELETE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch,
      Cli.prompt
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_ROLE_DELETE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });


  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --confirm option specified', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role name and --confirm option specified', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role id and --confirm option specified', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role claim and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --confirm option specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --confirm option specified (debug)', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --confirm option specified (debug)', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --confirm option specified (debug)', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --confirm option specified (debug)', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --confirm option specified (debug)', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when multiple apps with the specified appName found and --confirm option is specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            },
            {
              id: 'a39c738c-939e-433b-930d-b02f2931a08b'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    //sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple Azure AD application registration with name App-Name found. Please disambiguate using app object IDs: 5b31c38c-2584-42f0-aa47-657fb3a84230, a39c738c-939e-433b-930d-b02f2931a08b`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles when multiple roles with the same name are found and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple roles with the provided 'name' were found. Please disambiguate using the claims : Product.Read, Product.Get`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles when no roles with the specified name are found and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        name: 'ProductRead',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No app role with name 'ProductRead' found.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles when no roles with the specified claim are found and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        claim: 'Product.Read',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No app role with claim 'Product.Read' found.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles when no roles with the specified id are found and --confirm option specified', (done) => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No app role with id 'c4352a0a-494f-46f9-b843-479855c173a7' found.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified app role when confirm option not passed', (done) => {
    command.action(logger, { options: { debug: false, appName: 'App-Name', claim: 'Product.Read' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified app role when confirm option not passed (debug)', (done) => {
    command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and the prompt is confirmed (debug)', (done) => {

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        confirm: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and prompt is confirmed', (done) => {

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        confirm: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and prompt is confirmed (debug)', (done) => {

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7'&$select=id`) {
        return Promise.resolve({
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    getRequestStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
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
        });
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    patchStub.onSecondCall().callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return Promise.resolve();
        }
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });


    command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        confirm: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts deleting app role when prompt is not confirmed', (done) => {
    // represents the aad app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, appName: 'App-Name', claim: 'Product.Read' } }, () => {
      try {
        assert(patchStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts deleting app role when prompt is not confirmed (debug)', (done) => {
    // represents the aad app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } }, () => {
      try {
        assert(patchStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with appObjectId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.reject({
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-04-20T17:22:30",
              "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
              "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
            }
          }
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'App-Role',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with the appId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'App-Role',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with appName not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app',
        name: 'App-Role',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with name My app found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if appId and appObjectId specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role name and id is specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation role name and claim is specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", claim: "Product.Read" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role id and claim is specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: "Product.Read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither role name, id or claim specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified role id is not a valid guid', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '77355bee' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,name', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,claim', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,id', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,name', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,claim', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,id', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,name', () => {
    const actual = command.validate({ options: { appName: 'My App', name: 'ProductRead' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,claim', () => {
    const actual = command.validate({ options: { appName: 'My App', claim: 'Product.Read' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,id', () => {
    const actual = command.validate({ options: { appName: 'My App', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });










});
