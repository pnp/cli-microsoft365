import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./bucket-remove');

describe(commands.BUCKET_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  const bucketResponseValue = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/buckets/$entity",
    "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
    "name": "Planner Bucket A",
    "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
    "orderHint": "8585768731950308408",
    "id": "FtzysDykv0-9s9toWiZhdskAD67z"
  };

  const bucketListResponseValue = [
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket A",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585768731950308408",
      "id": "FtzysDykv0-9s9toWiZhdskAD67z"
    },
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket 2",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585784565[8",
      "id": "ZpcnnvS9ZES2pb91RPxQx8kAMLo5"
    }
  ];

  const bucketListResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerBucket)",
    "@odata.count": 2,
    "value": bucketListResponseValue
  };

  const groupByDisplayNameResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": "f3f985d0-a4e0-4891-83f6-08d88bf44e5e",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2021-06-08T11:04:45Z",
        "creationOptions": [],
        "description": "My Planner Group",
        "displayName": "My Planner Group",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "isAssignableToRole": null,
        "mail": "MyPlannerGroup@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "My Planner Group",
        "membershipRule": null,
        "membershipRuleProcessingState": null,
        "onPremisesDomainName": null,
        "onPremisesLastSyncDateTime": null,
        "onPremisesNetBiosName": null,
        "onPremisesSamAccountName": null,
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "preferredLanguage": null,
        "proxyAddresses": [
          "SPO:SPO_e13f6193-fb01-43e8-8e8d-557796b82ebf@SPO_cc6fafe9-dd93-497c-b521-1d971b1471c7",
          "SMTP:MyPlannerGroup@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2021-06-08T11:04:45Z",
        "resourceBehaviorOptions": [],
        "resourceProvisioningOptions": [],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-218366702-1230083855-573552016-1076796785",
        "theme": null,
        "visibility": "Private",
        "onPremisesProvisioningErrors": []
      }
    ]
  };

  const plansInOwnerGroup: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans",
    "@odata.count": 2,
    "value": [
      {
        "@odata.etag": "W/\"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBASCc=\"",
        "createdDateTime": "2021-06-08T12:24:57.3312829Z",
        "owner": "f3f985d0-a4e0-4891-83f6-08d88bf44e5e",
        "title": "My Planner Plan",
        "id": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "73829066-5f0a-4745-8f72-12a17bacadea"
          },
          "application": {
            "displayName": null,
            "id": "09abbdfd-ed25-47ee-a2d9-a627aa1c90f3"
          }
        }
      },
      {
        "@odata.etag": "W/\"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBASCc=\"",
        "createdDateTime": "2021-06-08T12:25:09.3751058Z",
        "owner": "f3f985d0-a4e0-4891-83f6-08d88bf44e5e",
        "title": "Sample Plan",
        "id": "uO1bj3fdekKuMitpeJqaj8kADBxO",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "73829066-5f0a-4745-8f72-12a17bacadea"
          },
          "application": {
            "displayName": null,
            "id": "09abbdfd-ed25-47ee-a2d9-a627aa1c90f3"
          }
        }
      }
    ]
  };

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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.delete,
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
    assert.strictEqual(command.name.startsWith(commands.BUCKET_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['bucketId', 'bucketName', 'planId', 'confirm']);
  });

  it('passes validation when valid bucketId is specified', () => {
    const actual = command.validate({
      options: {
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName and planId is specified', () => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName, planName & ownerGroupId is specified', () => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        ownerGroupId: 'f3f985d0-a4e0-4891-83f6-08d88bf44e5e'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName, planName & ownerGroupName is specified', () => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        ownerGroupName: 'My Planner Plan'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('removes the specified bucket by id when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) > -1) {
        return Promise.resolve(bucketResponseValue);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) > -1) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        bucketId: 'FtzysDykv0-9s9toWiZhdskAD67z'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified bucket by name when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/C0tdybe9NE2VvNrgR5v7xskAAJOi/buckets/bucketid`) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((_options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'object');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified planner bucket with bucketName, planName, and ownerGroupName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if ((opts.url as string).indexOf(`planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('f3f985d0-a4e0-4891-83f6-08d88bf44e5e')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/C0tdybe9NE2VvNrgR5v7xskAAJOi/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((_options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: false,
        bucketName: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'object');
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('removes the specified planner bucket with bucketName, planName, and ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('f3f985d0-a4e0-4891-83f6-08d88bf44e5e')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      if ((opts.url as string).indexOf(`planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/C0tdybe9NE2VvNrgR5v7xskAAJOi/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((_options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: false,
        bucketName: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupId: 'f3f985d0-a4e0-4891-83f6-08d88bf44e5e'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'object');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('fails to remove the specified planner bucket when planName does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('f3f985d0-a4e0-4891-83f6-08d88bf44e5e')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      if ((opts.url as string).indexOf(`planner/plans/undefined/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/C0tdybe9NE2VvNrgR5v7xskAAJOi/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((_options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: false,
        bucketName: 'Planner Bucket A',
        planName: 'Invalid plan',
        ownerGroupId: 'f3f985d0-a4e0-4891-83f6-08d88bf44e5e'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified plan does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });


  it('fails to removes the specified planner bucket when ownerGroupName does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq`) > -1) {
        return Promise.resolve({ value: [] });
      }
      if ((opts.url as string).indexOf(`planner/plans/undefined/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/undeifined/buckets/undefined`) {
        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        bucketName: 'Planner Bucket A',
        planName: 'Invalid plan',
        ownerGroupName: 'Invalid Group',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified owner group does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to remove the bucket by name and prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) > -1) {
        return Promise.resolve(bucketListResponse);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets/undefined`) {

        return Promise.resolve('');
      }
      return Promise.reject('Invalid request');
    });
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((_options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        bucketName: 'Invalid bucket name',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified bucket does not exist in the Microsoft Planner`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the bucketId is not provided', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the  bucketName & planId are not provided', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the bucketName, planName & ownerGroupId are not provided', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the bucketName, planName & ownerGroupName are not provided', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupName or ownerGroupId are not provided', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'My Bucket',
        planName: 'My Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'My Bucket',
        planName: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both bucketName and bucketId are provided', () => {
    const actual = command.validate({
      options: {
        bucketName: 'My Bucket',
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both planName and planId are provided', () => {
    const actual = command.validate({
      options: {
        planName: 'My Plan',
        planId: 'C0tdybe9NE2VvNrgR5v7xskAAJOi',
        bucketName: 'My Bucket'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both ownerGroupName and ownerGroupId are provided', () => {
    const actual = command.validate({
      options: {
        planName: 'My Plan',
        bucketName: 'My Bucket',
        ownerGroupName: 'My Group',
        ownerGroupId: '0b974c15-ff0e-410b-b26b-856a1f0ac593'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/IObYKVZEVEK9qDa5RmeszskAJwCp`) {
        return Promise.reject("The specified bucket does not exist in the Microsoft Planner");
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "The specified bucket does not exist in the Microsoft Planner");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle Microsoft graph error response with confirm', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/IObYKVZEVEK9qDa5RmeszskAJwCp`) {
        return Promise.reject("The specified bucket does not exist in the Microsoft Planner");
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp',
        confirm: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "The specified bucket does not exist in the Microsoft Planner");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified bucket when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    }, () => {
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

  it('prompts before removing the specified bucket when confirm option not passed (debug)', (done) => {
    command.action(logger, {
      options: {
        debug: true,
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    }, () => {
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

  it('aborts removing the specified bucket when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    command.action(logger, {
      options: {
        debug: true,
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified bucket when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    command.action(logger, {
      options: {
        debug: true,
        bucketId: 'IObYKVZEVEK9qDa5RmeszskAJwCp'
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
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
