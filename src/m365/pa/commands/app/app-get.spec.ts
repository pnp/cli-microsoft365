import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import * as paAppListCommand from '../app/app-list';
const command: Command = require('./app-get');

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      Cli.executeCommand,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.APP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName', 'description', 'appVersion', 'owner']);
  });

  it('fails validation if name or displayName not specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name and displayName are both specified', () => {
    const actual = command.validate({ options: { name: "5369f386-e380-46cb-82a4-4e18f9e4f3a7", displayName: "Playwright" } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name is not GUID', () => {
    const actual = command.validate({ options: { name: "TestApp" } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if name is specified', () => {
    const actual = command.validate({ options: { name: "5369f386-e380-46cb-82a4-4e18f9e4f3a7" } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if displayName is specified', () => {
    const actual = command.validate({ options: { displayName: "Playwright" } });
    assert.strictEqual(actual, true);
  });

  it('retrieves information about the specified app using name (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "test", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp" });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, name: '5369f386-e380-46cb-82a4-4e18f9e4f3a7' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "test", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp", displayName: 'App', description: 'test', appVersion: '2021-01-18T12:54:47Z', owner: 'admin@contoso.com' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified app using name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp" });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: '5369f386-e380-46cb-82a4-4e18f9e4f3a7' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp", displayName: 'App', description: '', appVersion: '2021-01-18T12:54:47Z', owner: 'admin@contoso.com' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`retrieves information about the specified app using displayName`, (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === paAppListCommand) {
        return Promise.resolve({ "stdout": JSON.stringify(apps) });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'Playwright' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          {
            "name": "f581c872-9852-4100-8e25-3d6891595204",
            "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
            "type": "Microsoft.PowerApps/apps",
            "tags": {
              "primaryDeviceWidth": "640",
              "primaryDeviceHeight": "1136",
              "sienaVersion": "20200812T204016Z-3.20074.20.0",
              "deviceCapabilities": "",
              "supportsPortrait": "true",
              "supportsLandscape": "false",
              "primaryFormFactor": "Phone",
              "publisherVersion": "3.20074.20",
              "minimumRequiredApiVersion": "2.2.0",
              "hasComponent": "false",
              "hasUnlockedComponent": "false",
              "isUnifiedRootApp": "false"
            },
            "properties": {
              "appVersion": "2020-08-12T20:40:16Z",
              "lastDraftVersion": "2020-08-12T20:40:16Z",
              "lifeCycleId": "Published",
              "status": "Ready",
              "createdByClientVersion": {
                "major": 3,
                "minor": 20074,
                "build": 20,
                "revision": 0,
                "majorRevision": 0,
                "minorRevision": 0
              },
              "minClientVersion": {
                "major": 3,
                "minor": 20074,
                "build": 20,
                "revision": 0,
                "majorRevision": 0,
                "minorRevision": 0
              },
              "owner": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "garry@trinder365dev.onmicrosoft.com",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "createdBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "garry@trinder365dev.onmicrosoft.com",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "lastModifiedBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "garry@trinder365dev.onmicrosoft.com",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "lastPublishedBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "garry@trinder365dev.onmicrosoft.com",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "backgroundColor": "rgba(0, 176, 240, 1)",
              "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "displayName": "Playwright",
              "description": "",
              "commitMessage": "",
              "appUris": {
                "documentUri": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
                },
                "imageUris": []
              },
              "createdTime": "2020-08-10T23:28:41.8191546Z",
              "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
              "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
              "sharedGroupsCount": 0,
              "sharedUsersCount": 0,
              "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
              "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
              "connectionReferences": {
                "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
                  "id": "/providers/microsoft.powerapps/apis/shared_office365users",
                  "displayName": "Office 365 Users",
                  "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
                  "dataSources": [
                    "Office365Users"
                  ],
                  "dependencies": [],
                  "dependents": [],
                  "isOnPremiseConnection": false,
                  "bypassConsent": false,
                  "dataSets": {},
                  "apiTier": "Standard",
                  "isCustomApiConnection": false
                }
              },
              "userAppMetadata": {
                "favorite": "NotSpecified",
                "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
                "includeInAppsList": true
              },
              "isFeaturedApp": false,
              "bypassConsent": false,
              "isHeroApp": false,
              "environment": {
                "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
                "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
              },
              "appPackageDetails": {
                "playerPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "sizeInBytes": 0
                },
                "webPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
                },
                "unauthenticatedWebPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
                },
                "documentServerVersion": {
                  "major": 3,
                  "minor": 20074,
                  "build": 20,
                  "revision": 0,
                  "majorRevision": 0,
                  "minorRevision": 0
                },
                "appPackageResourcesKind": "Split",
                "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
              },
              "almMode": "Environment",
              "performanceOptimizationEnabled": true,
              "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
              "canConsumeAppPass": true,
              "appPlanClassification": "Standard",
              "usesPremiumApi": false,
              "usesOnlyGrandfatheredPremiumApis": true,
              "usesCustomApi": false,
              "usesOnPremiseGateway": false
            },
            "isAppComponentLibrary": false,
            "appType": "ClassicCanvasApp",
            "displayName": "Playwright",
            "description": "",
            "appVersion": "2020-08-12T20:40:16Z",
            "owner": "garry@trinder365dev.onmicrosoft.com"
          }
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renders empty string for missing properties using name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp" });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, name: '5369f386-e380-46cb-82a4-4e18f9e4f3a7' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({ "name": "5369f386-e380-46cb-82a4-4e18f9e4f3a7", "id": "/providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "type": "Microsoft.PowerApps/apps", "tags": { "primaryDeviceWidth": "640", "primaryDeviceHeight": "1136", "sienaVersion": "20210118T125447Z-3.20123.24.0", "deviceCapabilities": "", "supportsPortrait": "true", "supportsLandscape": "false", "primaryFormFactor": "Phone", "publisherVersion": "3.20123.24", "minimumRequiredApiVersion": "2.2.0", "hasComponent": "false", "hasUnlockedComponent": "false", "isUnifiedRootApp": "false" }, "properties": { "appVersion": "2021-01-18T12:54:47Z", "createdByClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "minClientVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "owner": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "createdBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Contoso Admin", "email": "admin@contoso.com", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-1172edf9f339", "userPrincipalName": "admin@contoso.com" }, "backgroundColor": "rgba(0, 176, 240, 1)", "backgroundImageUri": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/logoSmallFile?sv=2018-03-28&sr=c&sig=Y5OvBpqU9EXwpXthPre62%2B24zwCNS9ihPfCNmBmuXro%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsColorIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "teamsOutlineIconUrl": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z7f1ae936e0c84483a435f047250f9a90/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=ECrnlPPWWirepGA0BwPAfi3RDcNxugeuBajcE9DvJcI%3D&se=2021-04-03T02%3A12%3A49Z&sp=rl", "displayName": "App", "description": "", "appUris": { "documentUri": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z6c2a80eb28694f6d8ded45d116516cfe/document.msapp?sv=2018-03-28&sr=c&sig=bx4dLMgjOurFPWW%2FuanqlJlb1clSqH05cZDPgbEkEno%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "imageUris": [] }, "createdTime": "2021-01-18T12:54:47.9591484Z", "lastModifiedTime": "2021-01-18T12:54:48.0998249Z", "sharedGroupsCount": 0, "sharedUsersCount": 0, "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/5369f386-e380-46cb-82a4-4e18f9e4f3a7", "appOpenUri": "https://apps.powerapps.com/play/5369f386-e380-46cb-82a4-4e18f9e4f3a7?tenantId=2ca3eaa5-140f-4175-9563-1172edf9f339&hint=c67d9086-a429-45b8-8b19-91fab5174177", "connectionReferences": { "c189738b-6a2d-4713-939d-e2ed35101124": { "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline", "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1431/1.0.1431.2301/sharepointonline/icon.png", "dataSources": ["ICT Aanvragen"], "dependencies": [], "dependents": [], "parameterHints": {}, "isOnPremiseConnection": false, "bypassConsent": false, "apiTier": "Standard", "isCustomApiConnection": false } }, "userAppMetadata": { "favorite": "NotSpecified", "includeInAppsList": true }, "isFeaturedApp": false, "bypassConsent": false, "isHeroApp": false, "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-1172edf9f339", "name": "Default-2ca3eaa5-140f-4175-9563-1172edf9f339" }, "appPackageDetails": { "playerPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/player.msappk?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "webPackage": { "value": "https://pafeblobprodam.blob.core.windows.net:443/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl", "readonlyValue": "https://pafeblobprodam-secondary.blob.core.windows.net/20210118t000000z1e14514d87e148ef915413781752b174/bda64b54-eb93-4ca8-b1a7-d6f85536c985/web/index.web.html?sv=2018-03-28&sr=c&sig=TwHX9CQJk8KEC6sjIWvYkMJ60v7jY%2FH534DYfL7tyAw%3D&se=2021-02-17T00%3A00%3A00Z&sp=rl" }, "unauthenticatedWebPackage": { "value": "https://pafeblobprodam.blob.core.windows.net/alt20210118t000000z84a0aa047e784c87b15a6529a8f7ab50/20210118T125458Z/index.web.html" }, "documentServerVersion": { "major": 3, "minor": 20123, "build": 24, "revision": 0, "majorRevision": 0, "minorRevision": 0 }, "appPackageResourcesKind": "Split", "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/a5ccjkurht6g5/preloadindex.web.html\",\"id\":\"637465712959244709\",\"v\":2.1}" }, "almMode": "Environment", "performanceOptimizationEnabled": true, "unauthenticatedWebPackageHint": "c67d9086-a429-45b8-8b19-91fab5174177", "canConsumeAppPass": true, "executionRestrictions": { "isTeamsOnly": false, "dataLossPreventionEvaluationResult": { "status": "Compliant", "lastEvaluationDate": "2021-01-18T12:54:58.0824718Z", "violationDetails": [] } }, "appPlanClassification": "Standard", "usesPremiumApi": false, "usesOnlyGrandfatheredPremiumApis": true, "usesCustomApi": false, "usesOnPremiseGateway": false, "isCustomizable": true }, "isAppComponentLibrary": false, "appType": "ClassicCanvasApp", displayName: 'App', description: '', appVersion: '2021-01-18T12:54:47Z', owner: '' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renders empty string for missing properties using displayName', (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === paAppListCommand) {
        return Promise.resolve({ "stdout": JSON.stringify(apps) });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'Playwright' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          {
            "name": "f581c872-9852-4100-8e25-3d6891595204",
            "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
            "type": "Microsoft.PowerApps/apps",
            "tags": {
              "primaryDeviceWidth": "640",
              "primaryDeviceHeight": "1136",
              "sienaVersion": "20200812T204016Z-3.20074.20.0",
              "deviceCapabilities": "",
              "supportsPortrait": "true",
              "supportsLandscape": "false",
              "primaryFormFactor": "Phone",
              "publisherVersion": "3.20074.20",
              "minimumRequiredApiVersion": "2.2.0",
              "hasComponent": "false",
              "hasUnlockedComponent": "false",
              "isUnifiedRootApp": "false"
            },
            "properties": {
              "appVersion": "2020-08-12T20:40:16Z",
              "lastDraftVersion": "2020-08-12T20:40:16Z",
              "lifeCycleId": "Published",
              "status": "Ready",
              "createdByClientVersion": {
                "major": 3,
                "minor": 20074,
                "build": 20,
                "revision": 0,
                "majorRevision": 0,
                "minorRevision": 0
              },
              "minClientVersion": {
                "major": 3,
                "minor": 20074,
                "build": 20,
                "revision": 0,
                "majorRevision": 0,
                "minorRevision": 0
              },
              "owner": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "createdBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "lastModifiedBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "lastPublishedBy": {
                "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
                "displayName": "Garry Trinder",
                "email": "",
                "type": "User",
                "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
                "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
              },
              "backgroundColor": "rgba(0, 176, 240, 1)",
              "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
              "displayName": "Playwright",
              "description": "",
              "commitMessage": "",
              "appUris": {
                "documentUri": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
                },
                "imageUris": []
              },
              "createdTime": "2020-08-10T23:28:41.8191546Z",
              "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
              "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
              "sharedGroupsCount": 0,
              "sharedUsersCount": 0,
              "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
              "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
              "connectionReferences": {
                "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
                  "id": "/providers/microsoft.powerapps/apis/shared_office365users",
                  "displayName": "Office 365 Users",
                  "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
                  "dataSources": [
                    "Office365Users"
                  ],
                  "dependencies": [],
                  "dependents": [],
                  "isOnPremiseConnection": false,
                  "bypassConsent": false,
                  "dataSets": {},
                  "apiTier": "Standard",
                  "isCustomApiConnection": false
                }
              },
              "userAppMetadata": {
                "favorite": "NotSpecified",
                "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
                "includeInAppsList": true
              },
              "isFeaturedApp": false,
              "bypassConsent": false,
              "isHeroApp": false,
              "environment": {
                "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
                "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
              },
              "appPackageDetails": {
                "playerPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "sizeInBytes": 0
                },
                "webPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
                  "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
                },
                "unauthenticatedWebPackage": {
                  "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
                },
                "documentServerVersion": {
                  "major": 3,
                  "minor": 20074,
                  "build": 20,
                  "revision": 0,
                  "majorRevision": 0,
                  "minorRevision": 0
                },
                "appPackageResourcesKind": "Split",
                "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
              },
              "almMode": "Environment",
              "performanceOptimizationEnabled": true,
              "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
              "canConsumeAppPass": true,
              "appPlanClassification": "Standard",
              "usesPremiumApi": false,
              "usesOnlyGrandfatheredPremiumApis": true,
              "usesCustomApi": false,
              "usesOnPremiseGateway": false
            },
            "isAppComponentLibrary": false,
            "appType": "ClassicCanvasApp",
            "displayName": "Playwright",
            "description": "",
            "appVersion": "2020-08-12T20:40:16Z",
            "owner": ""
          }
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles App not found using name', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "AppNotFound",
          "message": "Could not find App '1c6ee23a-a835-44bc-a4f5-462b658efc12'."
        }
      });
    });

    command.action(logger, { options: { debug: false, name: '1c6ee23a-a835-44bc-a4f5-462b658efc12' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Could not find App '1c6ee23a-a835-44bc-a4f5-462b658efc12'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly handles App not found using displayName (debug)', (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === paAppListCommand) {
        return Promise.resolve({ "stdout": JSON.stringify(apps) });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'NoAppFound' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(`No app found with displayName 'NoAppFound'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps found using displayName (debug)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === paAppListCommand) {
        return Promise.resolve({ "stdout": JSON.stringify([]) });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'Playwright' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(`No apps found`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } } as any, (err?: any) => {
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

  it('supports specifying name', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
