import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./externaluser-list');

describe(commands.EXTERNALUSER_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.EXTERNALUSER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists first page of 10 tenant external users (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 tenant external users', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 50 tenant external users', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">50</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, pageSize: '50' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists second page of 50 tenant external users', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">1</Parameter><Parameter Type="Int32">50</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, position: '1', pageSize: '50' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 tenant external users whose name match Vesa', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String">Vesa</Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, filter: 'Vesa' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 tenant external users sorted descending by email', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, sortOrder: 'desc' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 external users for the specified site (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 external users for the specified site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 50 external users for the specified site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">50</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, pageSize: '50', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists second page of 50 external users for the specified site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">1</Parameter><Parameter Type="Int32">50</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, position: '1', pageSize: '50', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 external users for the specified site whose name match Vesa', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String">Vesa</Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, filter: 'Vesa', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('lists first page of 10 external users for the specified site sorted descending by email', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String"></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "Dear Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, sortOrder: 'desc', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: 'Dear Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('escapes XML in user input', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Int32">0</Parameter><Parameter Type="Int32">10</Parameter><Parameter Type="String">&lt;Vesa</Parameter><Parameter Type="Enum">0</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "52e83b9e-4011-4000-c878-ab3e0977e9c2"
          }, 159, {
            "IsNull": false
          }, 160, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 1, "UserCollectionPosition": -1, "ExternalUserCollection": {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [
                {
                  "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUser", "DisplayName": "<Vesa", "InvitedAs": "me@dearvesa.fi", "UniqueId": "100300009BF10C95", "AcceptedAs": "me@dearvesa.fi", "WhenCreated": "\/Date(2016,10,2,21,50,52,0)\/", "InvitedBy": null
                }
              ]
            }
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, filter: '<Vesa', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([{
      DisplayName: '<Vesa',
      InvitedAs: 'me@dearvesa.fi',
      UniqueId: '100300009BF10C95',
      AcceptedAs: 'me@dearvesa.fi',
      WhenCreated: new Date(2016, 10, 2, 21, 50, 52, 0),
      InvitedBy: null
    }]));
  });

  it('correctly handles no results', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": null, "TraceCorrelationId": "5ae83b9e-30dd-4000-c878-aba6be0addde"
        }, 168, {
          "IsNull": false
        }, 169, {
          "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.GetExternalUsersResults", "TotalUserCount": 0, "UserCollectionPosition": -1, "ExternalUserCollection": {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantManagement.ExternalUserCollection", "_Child_Items_": [

            ]
          }
        }
      ]));
    });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles a generic error when retrieving external users', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "0ee83b9e-0000-4000-f938-f16a74e2f588", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "0ee83b9e-0000-4000-f938-f16a74e2f588"
        }
      ]));
    });
    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('File Not Found.'));
  });

  it('correctly handles a random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });
    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying page size', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--pageSize') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page number', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--position') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying filter', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--filter') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying sort order', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--sortOrder') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site URL', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('passes validation when no options have been specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when page size is not a number', async () => {
    const actual = await command.validate({ options: { pageSize: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when page size is a negative number', async () => {
    const actual = await command.validate({ options: { pageSize: '-10' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when page size is > 50', async () => {
    const actual = await command.validate({ options: { pageSize: '51' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when page size is 0 < x <= 50 (min)', async () => {
    const actual = await command.validate({ options: { pageSize: '1' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when page size is 0 < x <= 50 (max)', async () => {
    const actual = await command.validate({ options: { pageSize: '50' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when page number is not a number', async () => {
    const actual = await command.validate({ options: { position: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when page number is a negative number', async () => {
    const actual = await command.validate({ options: { position: '-1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when page number is a positive number', async () => {
    const actual = await command.validate({ options: { position: '1' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when sort order contains invalid value', async () => {
    const actual = await command.validate({ options: { sortOrder: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when sort order is set to asc', async () => {
    const actual = await command.validate({ options: { sortOrder: 'asc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when sort order is set to desc', async () => {
    const actual = await command.validate({ options: { sortOrder: 'desc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when site URL is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
