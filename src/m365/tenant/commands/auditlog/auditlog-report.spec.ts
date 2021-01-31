import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./auditlog-report');

describe(commands.TENANT_AUDITLOG_REPORT, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  let JSONActiveSubscription =
    [
      {
        "contentType": "Audit.Exchange",
        "status": "enabled",
        "webhook": null
      }
    ];

  const JSONListAuditContent =
    [
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      }
    ];

  const EmptyJSONListAuditContent: any = [];

  //Assumed Batch size is 10 or less than 10
  const JSONListAuditContentMultipleBatches =
    [
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212000045238027763$20201212002126942044256$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:21:26.942Z",
        "contentExpiration": "2020-12-26T00:00:45.238Z"
      },
      {
        "contentUri": "https://manage.office.com/api/v1.0/48526e9f-60c5-3000-31d7-aa1dc75ecf3c/activity/feed/audit/20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentId": "20201212002151960017850$20201212004708382033048$audit_exchange$Audit_Exchange$na0017",
        "contentType": "Audit.Exchange",
        "contentCreated": "2020-12-12T00:47:08.382Z",
        "contentExpiration": "2020-12-26T00:21:51.960Z"
      }
    ];

  const JSONAuditLogReport =
    [
      {
        "CreationTime": "2020-12-11T23:07:44",
        "Id": "fbcdc0c0-035c-43b7-9026-08d89e299188",
        "Operation": "ModifyFolderPermissions",
        "OrganizationId": "48526e9f-60c5-3000-31d7-aa1dc75ecf3c",
        "RecordType": 2,
        "ResultStatus": "Succeeded",
        "UserKey": "S-1-5-18",
        "UserType": 2,
        "Version": 1,
        "Workload": "Exchange",
        "ClientIP": "::1",
        "UserId": "S-1-5-18",
        "ClientIPAddress": "::1",
        "ClientInfoString": "Client=WebServices;Action=ConfigureGroupMailbox",
        "ExternalAccess": true,
        "InternalLogonType": 1,
        "LogonType": 1,
        "LogonUserSid": "S-1-5-18",
        "MailboxGuid": "ce1acd09-63e8-4885-8759-8a0cbb5dbd7c",
        "MailboxOwnerMasterAccountSid": "S-1-5-10",
        "MailboxOwnerSid": "S-1-5-21-612462314-3678279482-3515889748-25365147",
        "MailboxOwnerUPN": "myteam@contoso.onmicrosoft.com",
        "OrganizationName": "contoso.onmicrosoft.com",
        "OriginatingServer": "DM6PR04MB6251 (15.20.3654.014)\r\n",
        "Item": {
          "Id": "LgAAAADZgwex6+gNQpjB0CF3ufwTAQAGSRWQ/dSVRpRPkKGQJSE5AAAAAAENAAAC",
          "ParentFolder": {
            "Id": "LgAAAADZgwex6+gNQpjB0CF3ufwTAQAGSRWQ/dSVRpRPkKGQJSE5AAAAAAENAAAC",
            "MemberRights": "ReadAny, Create, EditOwned, DeleteOwned, EditAny, DeleteAny, Visible, FreeBusySimple, FreeBusyDetailed",
            "MemberSid": "S-1-8-3457862921-1216701416-210393479-2092785083-1",
            "MemberUpn": "Member@local",
            "Name": "Calendar",
            "Path": "\\Calendar"
          }
        }
      },
      {
        "CreationTime": "2020-12-11T23:07:44",
        "Id": "fbcdc0c0-035c-43b7-9026-08d89e299188",
        "Operation": "ModifyFolderPermissions",
        "OrganizationId": "48526e9f-60c5-3000-31d7-aa1dc75ecf3c",
        "RecordType": 2,
        "ResultStatus": "Succeeded",
        "UserKey": "S-1-5-18",
        "UserType": 2,
        "Version": 1,
        "Workload": "Exchange",
        "ClientIP": "::1",
        "UserId": "S-1-5-18",
        "ClientIPAddress": "::1",
        "ClientInfoString": "Client=WebServices;Action=ConfigureGroupMailbox",
        "ExternalAccess": true,
        "InternalLogonType": 1,
        "LogonType": 1,
        "LogonUserSid": "S-1-5-18",
        "MailboxGuid": "ce1acd09-63e8-4885-8759-8a0cbb5dbd7c",
        "MailboxOwnerMasterAccountSid": "S-1-5-10",
        "MailboxOwnerSid": "S-1-5-21-612462314-3678279482-3515889748-25365147",
        "MailboxOwnerUPN": "myteam@contoso.onmicrosoft.com",
        "OrganizationName": "contoso.onmicrosoft.com",
        "OriginatingServer": "DM6PR04MB6251 (15.20.3654.014)\r\n",
        "Item": {
          "Id": "LgAAAADZgwex6+gNQpjB0CF3ufwTAQAGSRWQ/dSVRpRPkKGQJSE5AAAAAAENAAAC",
          "ParentFolder": {
            "Id": "LgAAAADZgwex6+gNQpjB0CF3ufwTAQAGSRWQ/dSVRpRPkKGQJSE5AAAAAAENAAAC",
            "MemberRights": "ReadAny, Create, EditOwned, DeleteOwned, EditAny, DeleteAny, Visible, FreeBusySimple, FreeBusyDetailed",
            "MemberSid": "S-1-8-3457862921-1216701416-210393479-2092785083-1",
            "MemberUpn": "Member@local",
            "Name": "Calendar",
            "Path": "\\Calendar"
          }
        }
      }
    ];

  const AuditLogOutput =
    [
      {
        "CreationTime": "2020-12-11T23:07:44",
        "Id": "fbcdc0c0-035c-43b7-9026-08d89e299188",
        "Operation": "ModifyFolderPermissions",
        "Workload": "Exchange",
        "UserId": "S-1-5-18"
      },
      {
        "CreationTime": "2020-12-11T23:07:44",
        "Id": "fbcdc0c0-035c-43b7-9026-08d89e299188",
        "Operation": "ModifyFolderPermissions",
        "Workload": "Exchange",
        "UserId": "S-1-5-18"
      }
    ];

  const NOAuditLogOutput: any = [];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.tenantId = '48526e9f-60c5-3000-31d7-aa1dc75ecf3c|908bel80-a04a-4422-b4a0-883d9847d110:c8e761e2-d528-34d1-8776-dc51157d619a&#xA;Tenant';
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        value: 'abc'
      };
    }
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_AUDITLOG_REPORT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['CreationTime', 'Operation', 'ClientIP', 'UserId', 'Workload']);
  });

  it('Gets Audit Log Report - Subscription is Active', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(JSONAuditLogReport);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        contentType: 'Exchange'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report -  Subscription is NOT Active', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve([]);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(AuditLogOutput);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/start') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        contentType: 'Exchange'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report -  Subscription is NOT Active (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve([]);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(AuditLogOutput);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/start') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: 'Exchange'
      }
    } as any, (err?: any) => {
      try {
        console.log(err);
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report - Multiple Batches - Start and End time Specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContentMultipleBatches);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(JSONAuditLogReport);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        contentType: 'Exchange',
        startTime: '2020-12-13',
        endTime: '2020-12-14'
      }
    } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.args[0][0][0].Id, "fbcdc0c0-035c-43b7-9026-08d89e299188");
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report - Multiple Batches - Start and End time Specified (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContentMultipleBatches);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(JSONAuditLogReport);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: 'Exchange',
        startTime: '2020-12-13',
        endTime: '2020-12-14'
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report - Start and End time Specified - Subscription is Active (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(JSONAuditLogReport);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: 'Exchange',
        startTime: '2020-12-13',
        endTime: '2020-12-14'
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report - Only startTime Specified - Subscription is Active (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.resolve(JSONAuditLogReport);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: 'Exchange',
        startTime: '2020-12-13'
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('Gets Audit Log Report - When no Audit Logs Available', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(EmptyJSONListAuditContent);
      }

      return Promise.resolve(NOAuditLogOutput);
    });

    command.action(logger, {
      options: {
        debug: false,
        contentType: 'Exchange'
      }
    } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.args[0][0][0].Id, "fbcdc0c0-035c-43b7-9026-08d89e299188");
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('command correctly handles error while getting Complete Audit Log Report', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: "Exchange"
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles error for a particular Content Audit Content URL of a Batched Promise', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('activity/feed/subscriptions/list') > -1) {
        return Promise.resolve(JSONActiveSubscription);
      }

      if ((opts.url as string).indexOf('activity/feed/subscriptions/content') > -1) {
        return Promise.resolve(JSONListAuditContent);
      }

      if ((opts.url as string).indexOf('/activity/feed/audit/') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        contentType: 'Exchange'
      }
    } as any, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('passes validation if contentType is passed', () => {
    const actual = command.validate({ options: { contentType: "Exchange" } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if Content Type is not proper', () => {
    const actual = command.validate({ options: { contentType: "InvalidExchange" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if only endTime is entered', () => {
    const actual = command.validate({ options: { contentType: "Exchange", endTime: "2020-12-16" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime is Invalid', () => {
    const actual = command.validate({ options: { contentType: "Exchange", startTime: "InValidStartdate", endTime: "2020-12-16" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is Invalid', () => {
    const validstartTime: any = new Date();
    validstartTime.setDate(validstartTime.getDate() - 4);
    const actual = command.validate({ options: { contentType: "Exchange", startTime: validstartTime.toISOString(), endTime: "InValidEnddate" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime and endTime is more than 24 hours apart', () => {
    const validstartTime: any = new Date();
    validstartTime.setDate(validstartTime.getDate() - 4);
    const longerEndDateDuration: any = new Date();
    longerEndDateDuration.setDate(longerEndDateDuration.getDate() - 2);//End date more than 2 days from startdate
    const actual = command.validate({ options: { contentType: "Exchange", startTime: validstartTime.toISOString(), endTime: longerEndDateDuration.toISOString() } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime is more than 7 days before today', () => {
    const olderStarttime: any = new Date();
    olderStarttime.setDate(olderStarttime.getDate() - 9);//Start Date more than 9 days from today
    const actual = command.validate({ options: { contentType: "Exchange", startTime: olderStarttime.toISOString() } });
    assert.notStrictEqual(actual, true);
  });

});