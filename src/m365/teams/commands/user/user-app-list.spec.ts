import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./user-app-list');

describe(commands.USER_APP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.USER_APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', () => {
    const actual = command.validate({
      options: {
        userId: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and userName are not provided.', () => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN.', () => {
    const actual = command.validate({
      options: {
        userName: "no-an-email"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the both userId and userName are provided.', () => {
    const actual = command.validate({
      options: {
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userName: "admin@contoso.com"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct (userId)', () => {
    const actual = command.validate({
      options: {
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c6'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct (userName)', () => {
    const actual = command.validate({
      options: {
        userName: "admin@contoso.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('list apps from the catalog for the specified user (userId)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/5c705288-ed7f-44fc-af0a-ac164419901c/teamwork/installedApps?$expand=teamsAppDefinition`) {
        return Promise.resolve({
          "value": [
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
              "teamsAppDefinition": {
                "id": "MzT1NWIxZjktODUwNy00ZjU3LWLmNDktZGI5YXRiNGMyZWRkIyMxLjAuMS4wIyNQpWJsaXNoZDQ=",
                "teamsAppId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
                "displayName": "Whiteboard",
                "version": "1.0.5",
                "publishingState": "published",
                "shortDescription": "Use Microsoft Whiteboard to collaborate, visualize ideas, and work creatively",
                "description": "Create a new whiteboard and collaborate with others in real time.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            },
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
              "teamsAppDefinition": {
                "id": "MoT1NVIxZjktODUwNy033ZjU3LWLmNDktZGI5YXTiNGMyZWRkIyMxLjAuMS4wIyNQpWJsqXNoZLQ=",
                "teamsAppId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
                "displayName": "Evernote",
                "version": "1.0.1",
                "publishingState": "published",
                "shortDescription": "Capture, organize, and share notes",
                "description": "Unlock the power of teamwork—collect, organize and share the information, documents and ideas you encounter every day.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        userId: '5c705288-ed7f-44fc-af0a-ac164419901c'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
            "appId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
            "displayName": "Whiteboard",
            "version": "1.0.5"
          },
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
            "appId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
            "displayName": "Evernote",
            "version": "1.0.1"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('list apps from the catalog for the specified user (userName)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/5c705288-ed7f-44fc-af0a-ac164419901c/teamwork/installedApps?$expand=teamsAppDefinition`) {
        return Promise.resolve({
          "value": [
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
              "teamsAppDefinition": {
                "id": "MzT1NWIxZjktODUwNy00ZjU3LWLmNDktZGI5YXRiNGMyZWRkIyMxLjAuMS4wIyNQpWJsaXNoZDQ=",
                "teamsAppId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
                "displayName": "Whiteboard",
                "version": "1.0.5",
                "publishingState": "published",
                "shortDescription": "Use Microsoft Whiteboard to collaborate, visualize ideas, and work creatively",
                "description": "Create a new whiteboard and collaborate with others in real time.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            },
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
              "teamsAppDefinition": {
                "id": "MoT1NVIxZjktODUwNy033ZjU3LWLmNDktZGI5YXTiNGMyZWRkIyMxLjAuMS4wIyNQpWJsqXNoZLQ=",
                "teamsAppId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
                "displayName": "Evernote",
                "version": "1.0.1",
                "publishingState": "published",
                "shortDescription": "Capture, organize, and share notes",
                "description": "Unlock the power of teamwork—collect, organize and share the information, documents and ideas you encounter every day.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users/${encodeURIComponent("admin@contoso.com")}/id`) {
        return Promise.resolve({ "value": "5c705288-ed7f-44fc-af0a-ac164419901c" });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        userName: 'admin@contoso.com'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
            "appId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
            "displayName": "Whiteboard",
            "version": "1.0.5"
          },
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
            "appId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
            "displayName": "Evernote",
            "version": "1.0.1"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('list apps from the catalog for the specified user (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/5c705288-ed7f-44fc-af0a-ac164419901c/teamwork/installedApps?$expand=teamsAppDefinition`) {
        return Promise.resolve({
          "value": [
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
              "teamsAppDefinition": {
                "id": "MzT1NWIxZjktODUwNy00ZjU3LWLmNDktZGI5YXRiNGMyZWRkIyMxLjAuMS4wIyNQpWJsaXNoZDQ=",
                "teamsAppId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
                "displayName": "Whiteboard",
                "version": "1.0.5",
                "publishingState": "published",
                "shortDescription": "Use Microsoft Whiteboard to collaborate, visualize ideas, and work creatively",
                "description": "Create a new whiteboard and collaborate with others in real time.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            },
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
              "teamsAppDefinition": {
                "id": "MoT1NVIxZjktODUwNy033ZjU3LWLmNDktZGI5YXTiNGMyZWRkIyMxLjAuMS4wIyNQpWJsqXNoZLQ=",
                "teamsAppId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
                "displayName": "Evernote",
                "version": "1.0.1",
                "publishingState": "published",
                "shortDescription": "Capture, organize, and share notes",
                "description": "Unlock the power of teamwork—collect, organize and share the information, documents and ideas you encounter every day.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        userId: '5c705288-ed7f-44fc-af0a-ac164419901c',
        debug: true
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
            "appId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
            "displayName": "Whiteboard",
            "version": "1.0.5"
          },
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
            "appId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
            "displayName": "Evernote",
            "version": "1.0.1"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('list apps from the catalog for the specified user (json)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/5c705288-ed7f-44fc-af0a-ac164419901c/teamwork/installedApps?$expand=teamsAppDefinition`) {
        return Promise.resolve({
          "value": [
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
              "teamsAppDefinition": {
                "id": "MzT1NWIxZjktODUwNy00ZjU3LWLmNDktZGI5YXRiNGMyZWRkIyMxLjAuMS4wIyNQpWJsaXNoZDQ=",
                "teamsAppId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
                "displayName": "Whiteboard",
                "version": "1.0.5",
                "publishingState": "published",
                "shortDescription": "Use Microsoft Whiteboard to collaborate, visualize ideas, and work creatively",
                "description": "Create a new whiteboard and collaborate with others in real time.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            },
            {
              "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
              "teamsAppDefinition": {
                "id": "MoT1NVIxZjktODUwNy033ZjU3LWLmNDktZGI5YXTiNGMyZWRkIyMxLjAuMS4wIyNQpWJsqXNoZLQ=",
                "teamsAppId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
                "displayName": "Evernote",
                "version": "1.0.1",
                "publishingState": "published",
                "shortDescription": "Capture, organize, and share notes",
                "description": "Unlock the power of teamwork—collect, organize and share the information, documents and ideas you encounter every day.",
                "lastModifiedDateTime": null,
                "createdBy": null
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        userId: '5c705288-ed7f-44fc-af0a-ac164419901c',
        output: 'json'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyMwOTg5ZjNhNC0yNWY3LTQ2YWItYTNjMC1iY2MwZWNmY2E2ZWY=",
            "teamsAppDefinition": {
              "id": "MzT1NWIxZjktODUwNy00ZjU3LWLmNDktZGI5YXRiNGMyZWRkIyMxLjAuMS4wIyNQpWJsaXNoZDQ=",
              "teamsAppId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef",
              "displayName": "Whiteboard",
              "version": "1.0.5",
              "publishingState": "published",
              "shortDescription": "Use Microsoft Whiteboard to collaborate, visualize ideas, and work creatively",
              "description": "Create a new whiteboard and collaborate with others in real time.",
              "lastModifiedDateTime": null,
              "createdBy": null
            },
            "appId": "0989f3a4-25f7-46ab-a3c0-bcc0ecfca6ef"
          },
          {
            "id": "NWM3MDUyODgtZWQ3Zi00NGZjLWFmMGEtYWMxNjQ0MTk5MDFjIyM5OTlhNTViOS00OTFlLTQ1NGEtODA4Yy1jNzVjNWM3NWZjMGE=",
            "teamsAppDefinition": {
              "id": "MoT1NVIxZjktODUwNy033ZjU3LWLmNDktZGI5YXTiNGMyZWRkIyMxLjAuMS4wIyNQpWJsqXNoZLQ=",
              "teamsAppId": "999a55b9-491e-454a-808c-c75c5c75fc0a",
              "displayName": "Evernote",
              "version": "1.0.1",
              "publishingState": "published",
              "shortDescription": "Capture, organize, and share notes",
              "description": "Unlock the power of teamwork—collect, organize and share the information, documents and ideas you encounter every day.",
              "lastModifiedDateTime": null,
              "createdBy": null
            },
            "appId": "999a55b9-491e-454a-808c-c75c5c75fc0a"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while listing teams apps', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');

    });

    command.action(logger, {
      options: {
        userId: '5c705288-ed7f-44fc-af0a-ac164419901c'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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