import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError, CommandTypes } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import * as aadO365GroupSetCommand from '../../../aad/commands/o365group/o365group-set';
import commands from '../../commands';
import * as spoSiteDesignApplyCommand from '../sitedesign/sitedesign-apply';
const command: Command = require('./site-set');

describe(commands.SITE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let executeCommandSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.tenantId = 'a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant';
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
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      Cli.executeCommand,
      global.setTimeout,
      spo.getSpoAdminUrl
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });


  it('fails validation if the url is not a valid url', () => {
    const actual = command.validate({
      options: {
        url: 'ABC', title: 'Team'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url is not a valid SharePoint url', () => {
    const actual = command.validate({
      options: {
        url: 'http://contoso', title: 'Team'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuota is not a number', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', resourceQuota: 'ABC'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuotaWarningLevel is not a number', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', resourceQuotaWarningLevel: 'ABC'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuotaWarningLevel is greater than resourceQuota', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', resourceQuotaWarningLevel: 10, resourceQuota: 5
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuota is not a number', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        storageQuota: 'ABC'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuotaWarningLevel is not a number', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        storageQuotaWarningLevel: 'ABC'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuotaWarningLevel is greater than storageQuota', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        storageQuotaWarningLevel: 10, storageQuota: 5
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sharing is invalid', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', sharing: 'Invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('fails validation if allowSelfServiceUpgrade is invalid', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', allowSelfServiceUpgrade: 'Invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if lockState is invalid', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', lockState: 'Invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if noScriptSite is invalid', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', noScriptSite: 'Invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified id is not a valid GUID', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', id: 'ABC' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no property to update specified (id specified)', () => {
    const actual = command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid value specified for disableFlows', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteLogoUrl is not a string', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/logo', siteLogoUrl: true } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if non-GUID value specified for siteDesignId', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', siteDesignId: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid value specified for shareByEmailEnabled', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', shareByEmailEnabled: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid value specified for isPublic', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', isPublic: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if allowSelfServiceUpgrade is set to true', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', allowSelfServiceUpgrade: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowSelfServiceUpgrade is set to false', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', allowSelfServiceUpgrade: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if lockState is set to Unlock', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', lockState: 'Unlock'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if lockState is set to NoAdditions', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAdditions'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if lockState is set to ReadOnly', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', lockState: 'ReadOnly'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if lockState is set to NoAccess', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if noScriptSite is set to true', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', noScriptSite: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if noScriptSite is set to false', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', noScriptSite: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options are correct', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        resourceQuota: 100, resourceQuotaWarningLevel: 90,
        storageQuota: 100, storageQuotaWarningLevel: 90,
        sharing: 'Disabled', allowSelfServiceUpgrade: 'true',
        owners: 'admin@contoso.com', lockState: 'Unlock', noScriptSite: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if url and classification specified', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', classification: 'HBI' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if url and empty classification specified', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', classification: '' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if url and disableFlows true specified', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if url and disableFlows false specified', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if url, id, classification and disableFlows specified', () => {
    const actual = command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com', classification: 'HBI', disableFlows: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if true specified for isPublic', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', isPublic: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if false specified for isPublic', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', isPublic: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if true specified for shareByEmailEnabled', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', shareByEmailEnabled: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if false specified for shareByEmailEnabled', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', shareByEmailEnabled: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if a valid GUID specified for siteDesignId', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', siteDesignId: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if non existing sharingCapability specified', () => {
    const sharingCapabilityvalue = 'nonExistentSharingCapabilityValue';
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', sharingCapability: sharingCapabilityvalue } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if correct sharingCapability specified', () => {
    const sharingCapabilityvalue = 'ExternalUserSharingOnly';
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com', sharingCapability: sharingCapabilityvalue } });
    assert.strictEqual(actual, true);
  });

  it('updates title of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'New title', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site title. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title' } }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site title. wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636540580851601240&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site title. wait for completion (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636540580851601240&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site title. wait for completion (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636540580851601240&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, verbose: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates title of the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_api/SPOGroup/UpdateGroupPropertiesBySiteId' &&
        JSON.stringify(opts.data) === JSON.stringify({
          groupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e',
          siteId: '255a50b2-527f-4413-8485-57f4c17a24d1',
          displayName: 'New title'
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'New title', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site description.', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team/_api/web`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.headers['X-HTTP-Method'] === 'MERGE' &&
          JSON.stringify(opts.data) === JSON.stringify({ Description: 'New description' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', description: 'New description' } }, (err) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates isPublic property of the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    executeCommandSpy = sinon.stub(Cli, 'executeCommand').callsFake(() => Promise.resolve());
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/e10a459e-60c8-4000-8240-a68d6a12d39e') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, isPublic: true, description: 'Some description', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        const options = {
          id: 'e10a459e-60c8-4000-8240-a68d6a12d39e',
          isPrivate: 'false',
          debug: false,
          verbose: false,
          _: []
        };
        assert(executeCommandSpy.calledWith(aadO365GroupSetCommand, { options: options }));
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess' } }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636543176568112997&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. wait for completion. error while polling', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636543176568112997&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess', wait: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. wait for completion two rounds', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // not done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636543176568112997&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636543176568112997&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site lockState. wait for completion, immediate complete', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates all properties. wait for completion, immediately complete', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><SetProperty Id="28" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">100</Parameter></SetProperty><SetProperty Id="29" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">100</Parameter></SetProperty><SetProperty Id="30" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">100</Parameter></SetProperty><SetProperty Id="31" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">100</Parameter></SetProperty><SetProperty Id="32" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="33" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">0</Parameter></SetProperty><SetProperty Id="34" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">2</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">admin@contoso.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "b3d8499e-1079-5000-cb83-9da72405dfa6"
            }, 48, {
              "IsNull": false
            }
          ]));
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title', sharingCapability: 'Disabled', resourceQuota: 100, resourceQuotaWarningLevel: 100, storageQuota: 100, storageQuotaWarningLevel: 100, allowSelfServiceUpgrade: 'true', noScriptSite: 'true', owners: 'admin@contoso.com', lockState: 'NoAccess', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site sharing mode. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', sharing: 'Disabled' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site resourceQuota. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">100</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', resourceQuota: 100 } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site resourceQuotaWarningLevel. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">100</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', resourceQuotaWarningLevel: 100 } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site storageQuota. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">100</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', storageQuota: 100 } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site storageQuotaWarningLevel. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">100</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', storageQuotaWarningLevel: 100 } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site allowSelfServiceUpgrade. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', allowSelfServiceUpgrade: 'true' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site noScriptSite to true. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">2</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', noScriptSite: 'true' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site noScriptSite to false. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">1</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', noScriptSite: 'false' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site title. waits for completion, immediately complete', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, classification: 'HBI', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, classification: 'HBI', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Site is not group connected'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, classification: 'HBI', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified groupified site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, classification: 'HBI', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(`Site attached to group e10a459e-60c8-4000-8240-a68d6a12d39e`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification, disableFlows and shareByEmailEnabled of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty><SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="29" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, classification: 'HBI', disableFlows: 'true', shareByEmailEnabled: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="String">admin@contoso.onmicrosoft.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: 'admin@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified site (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="String">admin@contoso.onmicrosoft.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, verbose: true, owners: 'admin@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified groupified site with one owner', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d729' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d729', principalName='')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: 'admin@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified groupified site with one owner (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d729' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d729', principalName='')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, owners: 'admin@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Retrieving user information to set group owners...'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified groupified site with multiple owners', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin1@contoso.onmicrosoft.com' or userPrincipalName eq 'admin2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d728' },
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d729' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d728', principalName='')`) {
        return Promise.resolve();
      }
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d729', principalName='')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: 'admin1@contoso.onmicrosoft.com,admin2@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates owners of the specified groupified site with multiple owners with extra spaces', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin1@contoso.onmicrosoft.com' or userPrincipalName eq 'admin2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d728' },
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d729' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d728', principalName='')`) {
        return Promise.resolve();
      }
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d729', principalName='')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: ' admin1@contoso.onmicrosoft.com , admin2@contoso.onmicrosoft.com ', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('resets the classification of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String"></Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, classification: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('resets the classification of the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String"></Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, classification: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets disableFlows to true for the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, disableFlows: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets disableFlows to true for the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, disableFlows: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets disableFlows to false for the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">false</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, disableFlows: 'false', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets shareByEmailEnabled to true for the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, shareByEmailEnabled: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets shareByEmailEnabled to true for the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, shareByEmailEnabled: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets shareByEmailEnabled to false for the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">false</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, shareByEmailEnabled: 'false', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets shareByEmailEnabled to false for the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery` &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">false</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, shareByEmailEnabled: 'false', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets sharingCapabilities for Site - Disabled', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            },
            2,
            {
              "IsNull": false
            },
            4,
            {
              "IsNull": false
            },
            7,
            {
              "IsNull": false
            },
            8,
            {
              "_ObjectIdentity_": "b61d6a9f-d0ca-0000-4814-29cf3242c81a|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%contoso.sharepoint.com%2fsites%2fSales"
            }
          ]
        ));
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: false, sharingCapability: 'Disabled', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets sharingCapabilities for Site - (Debug) -  Disabled', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">0</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            },
            2,
            {
              "IsNull": false
            },
            4,
            {
              "IsNull": false
            },
            7,
            {
              "IsNull": false
            },
            8,
            {
              "_ObjectIdentity_": "b61d6a9f-d0ca-0000-4814-29cf3242c81a|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%contoso.sharepoint.com%2fsites%2fSales"
            }
          ]
        ));
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, sharingCapability: 'Disabled', url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('throws error when trying to update isPublic property on a non-groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, isPublic: true, url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The isPublic option can't be set on a site that is not groupified`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when updating title of the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FSales" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'New title', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while updating isPublic property of the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    executeCommandSpy = sinon.stub(Cli, 'executeCommand').callsFake(() => Promise.reject(new CommandError('An error has occurred')));

    command.action(logger, { options: { debug: false, isPublic: true, url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips users that could not be resolves when setting groupified site owners', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin1@contoso.onmicrosoft.com' or userPrincipalName eq 'admin2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d728' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d728', principalName='')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: 'admin1@contoso.onmicrosoft.com,admin2@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails silently if could not resolve users when setting groupified site owners', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'admin1@contoso.onmicrosoft.com' or userPrincipalName eq 'admin2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, owners: 'admin1@contoso.onmicrosoft.com,admin2@contoso.onmicrosoft.com', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies site design to the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    executeCommandSpy = sinon.stub(Cli, 'executeCommand').callsFake(() => Promise.resolve());

    command.action(logger, { options: { debug: false, siteDesignId: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        const options = {
          webUrl: 'https://contoso.sharepoint.com/sites/Sales',
          id: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e',
          asTask: false,
          debug: false,
          verbose: false,
          _: []
        };
        assert(executeCommandSpy.calledWith(spoSiteDesignApplyCommand, { options: options }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies site design to the specified groupified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });
    executeCommandSpy = sinon.stub(Cli, 'executeCommand').callsFake(() => Promise.resolve());

    command.action(logger, { options: { debug: false, siteDesignId: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        const options = {
          webUrl: 'https://contoso.sharepoint.com/sites/Sales',
          id: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e',
          asTask: false,
          debug: false,
          verbose: false,
          _: []
        };
        assert(executeCommandSpy.calledWith(spoSiteDesignApplyCommand, { options: options }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies site relative logo url to the specified site', (done) => {
    let data: any = {};

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/siteiconmanager/setsitelogo') {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/logo', siteLogoUrl: "/sites/logo/SiteAssets/parker-ms-1200.png" } }, () => {
      try {
        assert.strictEqual(data.relativeLogoUrl, "/sites/logo/SiteAssets/parker-ms-1200.png");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies site absolute logo url to the specified site', (done) => {
    let data: any = {};

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/siteiconmanager/setsitelogo') {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/logo', siteLogoUrl: "https://contoso.sharepoint.com/sites/logo/SiteAssets/parker-ms-1200.png" } }, () => {
      try {
        assert.strictEqual(data.relativeLogoUrl, "/sites/logo/SiteAssets/parker-ms-1200.png");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles unsetting the logo from the specified site', (done) => {
    let data: any = {};

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/logo/_api/siteiconmanager/setsitelogo') {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/logo', siteLogoUrl: "" } }, () => {
      try {
        assert.strictEqual(data.relativeLogoUrl, "");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when applying site design to the specified site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    executeCommandSpy = sinon.stub(Cli, 'executeCommand').callsFake(() => Promise.reject(new CommandError('An error has occurred')));

    command.action(logger, { options: { debug: false, siteDesignId: 'eb2f31da-9461-4fbf-9ea1-9959b134b89e', url: 'https://contoso.sharepoint.com/sites/Sales' } } as any, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site not found error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1', classification: 'HBI' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("404 - \"404 FILE NOT FOUND\"")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API error while updating shared properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1', classification: 'HBI' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API error while updating sharingCapability properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', sharingCapability: 'Invalid' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Generic API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/Sales/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', sharingCapability: 'Disabled' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures classification as string option', () => {
    const types = (command.types() as CommandTypes);
    ['classification'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
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

  it('supports specifying site url', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site classification', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--classification') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying disableFlows', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--disableFlows') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site logo', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteLogoUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('handles error while adding site admin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">admin@contoso.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', owners: 'admin@contoso.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles generic error while adding site admin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">admin@contoso.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.reject('Unknown Error');
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', owners: 'admin@contoso.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while updating site lockState', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'ABC' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', lockState: 'NoAccess' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in the request', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team/_api/site?$select=GroupId,Id') {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">New title&gt;</Parameter></SetProperty><SetProperty Id="28" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">100</Parameter></SetProperty><SetProperty Id="29" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">100</Parameter></SetProperty><SetProperty Id="30" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">100</Parameter></SetProperty><SetProperty Id="31" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">100</Parameter></SetProperty><SetProperty Id="32" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="33" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">2</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2Fteam" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">admin@contoso.com&gt;</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "b3d8499e-1079-5000-cb83-9da72405dfa6"
            }, 48, {
              "IsNull": false
            }
          ]));
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">NoAccess</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1203", "ErrorInfo": null, "TraceCorrelationId": "e0cf4a9e-301d-5000-8242-77fc92412701"
            }, 9, {
              "IsNull": false
            }, 10, {
              "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 11, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e0cf4a9e-301d-5000-8242-77fc92412701|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636543176568112997\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "HasTimedout": false, "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'New title>', sharing: 'Disabled', resourceQuota: 100, resourceQuotaWarningLevel: 100, storageQuota: 100, storageQuotaWarningLevel: 100, allowSelfServiceUpgrade: 'true', noScriptSite: 'true', owners: 'admin@contoso.com>', lockState: 'NoAccess', wait: true } }, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});