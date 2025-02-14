import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-recyclebinitem-remove.js';

describe(commands.TENANT_RECYCLEBINITEM_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(spo, 'ensureFormDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(spo, 'waitUntilFinished').resolves();
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
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation,
      spo.waitUntilFinished
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name, commands.TENANT_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert(actual);
  });

  it('aborts removing deleting site when prompt not confirmed', async () => {
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', debug: true, verbose: true } });
  });

  it('removes the deleted site collection from the tenant recycle bin when prompt confirmed, doesn\'t wait for completion', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "PollingInterval": 0, "IsComplete": true
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "PollingInterval": 0, "IsComplete": true
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, debug: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "PollingInterval": 0, "IsComplete": true
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, verbose: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion. Operation immediately completed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "PollingInterval": 0, "IsComplete": true
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, wait: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, wait: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, wait: true, debug: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('removes the deleted site collection from the tenant recycle bin, wait for completion (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": null, "TraceCorrelationId": "5e0d879f-207a-2000-5eb4-5be71488a82a"
            }, 16, {
              "IsNull": false
            }, 17, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637392077403920220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n67edbe85-2b95-4c7b-a34a-1abbcd68dbe4", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="5e0d879f-207a-2000-5eb4-5be71488a82a|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;637392077403920220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;67edbe85-2b95-4c7b-a34a-1abbcd68dbe4" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRemoveDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, wait: true, verbose: true } });
    sinonUtil.restore([
      request.post
    ]);
  });

  it('did not remove the deleted site collection from the tenant recycle bin', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20614.12002", "ErrorInfo": {
                "ErrorMessage": "Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.", "ErrorValue": null, "TraceCorrelationId": "b319879f-4090-2000-6ca2-90bc9381b277", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "b319879f-4090-2000-6ca2-90bc9381b277"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr', force: true, wait: true } } as any), new CommandError('Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.'));
  });
});
