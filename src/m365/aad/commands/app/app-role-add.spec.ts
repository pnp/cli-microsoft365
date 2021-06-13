import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-role-add');

describe(commands.APP_ROLE_ADD, () => {
  let log: string[];
  let logger: Logger;

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
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.APP_ROLE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app role for the specified appId, app has no roles', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['User'])) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
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

  it('creates app role for the specified appObjectId, app has one role', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [{
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Managers",
            "displayName": "Managers",
            "id": "c4352a0a-494f-46f9-b843-479855c173a7",
            "isEnabled": true,
            "lang": null,
            "origin": "Application",
            "value": "managers"
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[1];
        if (JSON.stringify({
          "allowedMemberTypes": [
            "User"
          ],
          "description": "Managers",
          "displayName": "Managers",
          "id": "c4352a0a-494f-46f9-b843-479855c173a7",
          "isEnabled": true,
          "lang": null,
          "origin": "Application",
          "value": "managers"
        }) === JSON.stringify(opts.data.appRoles[0]) &&
          appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['Application'])) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'applications',
        claim: 'Custom.Role'
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

  it('creates app role for the specified appName, app has multiple roles', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Managers",
              "displayName": "Managers",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "managers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Team leads",
              "displayName": "Team leads",
              "id": "c4352a0a-494f-46f9-b843-479855c173a8",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "teamLeads"
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 3) {
        const appRole = opts.data.appRoles[2];
        if (appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['User', 'Application'])) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'both',
        claim: 'Custom.Role'
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
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
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
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
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
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
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

  it('handles error when multiple apps with the specified appName found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple Azure AD application registration with name My app found. Please disambiguate (app object IDs): 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appId failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appName failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving app roles failed', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when updating app roles failed', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return Promise.resolve({
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [{
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Managers",
            "displayName": "Managers",
            "id": "c4352a0a-494f-46f9-b843-479855c173a7",
            "isEnabled": true,
            "lang": null,
            "origin": "Application",
            "value": "managers"
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('An error has occurred'));

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
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
    const actual = command.validate({ options: { } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid allowedMembers specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim length exceeds 120 chars', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Cras ullamcorper, arcu vel finibus facilisis, orci velit lectus.' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim starts with a .', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: '.claim' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim contains invalid characters', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: 'cláim' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (appObjectId)', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (appName)', () => {
    const actual = command.validate({ options: { appName: 'My app', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } });
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

  it('returns an empty array for an invalid member type', () => {
    const actual = (command as any).getAllowedMemberTypes({ options: { allowedMembers: 'foo' } });
    assert.deepStrictEqual(actual, []);
  });
});
