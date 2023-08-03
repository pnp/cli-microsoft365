import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import command from './term-set-get.js';

describe(commands.TERM_SET_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const id = '7a167c47-2b37-41d0-94d0-e962c1a4f2ed';
  const name = 'PnP-CollabFooter-SharedLinks';
  const termGroupId = '0e8f395e-ff58-4d45-9ff7-e331ab728beb';
  const termGroupName = 'PnPTermSets';
  const getTermSetResponse = [
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.8112.1218",
      "ErrorInfo": null,
      "TraceCorrelationId": "2994929e-20f1-0000-2cdb-e577d70db169"
    },
    55,
    {
      "IsNull": false
    },
    56,
    {
      "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
    },
    58,
    {
      "IsNull": false
    },
    59,
    {
      "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
    },
    61,
    {
      "IsNull": false
    },
    63,
    {
      "IsNull": false
    },
    64,
    {
      "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
    },
    66,
    {
      "IsNull": false
    },
    68,
    {
      "IsNull": false
    },
    69,
    {
      "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt"
    },
    70,
    {
      "_ObjectType_": "SP.Taxonomy.TermSet",
      "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt",
      "CreatedDate": "\/Date(1536839573337)\/",
      "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/",
      "LastModifiedDate": "\/Date(1536840826883)\/",
      "Name": "PnP-CollabFooter-SharedLinks",
      "CustomProperties": {
        "_Sys_Nav_IsNavigationTermSet": "True"
      },
      "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
      "IsAvailableForTagging": true,
      "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
      "Contact": "",
      "Description": "",
      "IsOpenForTermCreation": false,
      "Names": {
        "1033": "PnP-CollabFooter-SharedLinks"
      },
      "Stakeholders": []
    }
  ];

  const termSetGetResponse = {
    "CreatedDate": "2018-09-13T11:52:53.337Z",
    "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed",
    "LastModifiedDate": "2018-09-13T12:13:46.883Z",
    "Name": "PnP-CollabFooter-SharedLinks",
    "CustomProperties": {
      "_Sys_Nav_IsNavigationTermSet": "True"
    },
    "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
    "Contact": "",
    "Description": "",
    "IsOpenForTermCreation": false,
    "Names": {
      "1033": "PnP-CollabFooter-SharedLinks"
    },
    "Stakeholders": []
  };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TERM_SET_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets taxonomy term set by id, term group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(getTermSetResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, termGroupId: termGroupId } });
    assert(loggerLogSpy.calledWith(termSetGetResponse));
  });

  it('gets taxonomy term set by name, term group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">PnP-CollabFooter-SharedLinks</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(getTermSetResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, termGroupId: termGroupId, debug: true } });
    assert(loggerLogSpy.calledWith(termSetGetResponse));
  });

  it('gets taxonomy term set by id, term group by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(getTermSetResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, termGroupName: termGroupName } });
    assert(loggerLogSpy.calledWith(termSetGetResponse));
  });

  it('gets taxonomy term set by name, term group by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">PnP-CollabFooter-SharedLinks</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(getTermSetResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, termGroupName: termGroupName } });
    assert(loggerLogSpy.calledWith(termSetGetResponse));
  });

  it('gets taxonomy term set by name, term group by name from the specified sitecollection', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">PnP-CollabFooter-SharedLinks</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(getTermSetResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, termGroupName: termGroupName, webUrl: webUrl } });
    assert(loggerLogSpy.calledWith(termSetGetResponse));
  });

  it('escapes XML in term group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">PnP-CollabFooter-SharedLinks</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8112.1218",
            "ErrorInfo": null,
            "TraceCorrelationId": "2994929e-20f1-0000-2cdb-e577d70db169"
          },
          55,
          {
            "IsNull": false
          },
          56,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          58,
          {
            "IsNull": false
          },
          59,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          61,
          {
            "IsNull": false
          },
          63,
          {
            "IsNull": false
          },
          64,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          },
          66,
          {
            "IsNull": false
          },
          68,
          {
            "IsNull": false
          },
          69,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt"
          },
          70,
          {
            "_ObjectType_": "SP.Taxonomy.TermSet",
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt",
            "CreatedDate": "\/Date(1536839573337)\/",
            "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/",
            "LastModifiedDate": "\/Date(1536840826883)\/",
            "Name": "PnP-CollabFooter-SharedLinks",
            "CustomProperties": {
              "_Sys_Nav_IsNavigationTermSet": "True"
            },
            "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
            "IsAvailableForTagging": true,
            "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Contact": "",
            "Description": "",
            "IsOpenForTermCreation": false,
            "Names": {
              "1033": "PnP-CollabFooter-SharedLinks"
            },
            "Stakeholders": []
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: name, termGroupName: 'PnPTermSets>' } });
    assert(loggerLogSpy.calledWith({
      "CreatedDate": "2018-09-13T11:52:53.337Z",
      "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed",
      "LastModifiedDate": "2018-09-13T12:13:46.883Z",
      "Name": "PnP-CollabFooter-SharedLinks",
      "CustomProperties": {
        "_Sys_Nav_IsNavigationTermSet": "True"
      },
      "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
      "IsAvailableForTagging": true,
      "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
      "Contact": "",
      "Description": "",
      "IsOpenForTermCreation": false,
      "Names": {
        "1033": "PnP-CollabFooter-SharedLinks"
      },
      "Stakeholders": []
    }));
  });

  it('escapes XML in term set name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /><Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">PnP-CollabFooter-SharedLinks&gt;</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8112.1218",
            "ErrorInfo": null,
            "TraceCorrelationId": "2994929e-20f1-0000-2cdb-e577d70db169"
          },
          55,
          {
            "IsNull": false
          },
          56,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          58,
          {
            "IsNull": false
          },
          59,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          61,
          {
            "IsNull": false
          },
          63,
          {
            "IsNull": false
          },
          64,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          },
          66,
          {
            "IsNull": false
          },
          68,
          {
            "IsNull": false
          },
          69,
          {
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt"
          },
          70,
          {
            "_ObjectType_": "SP.Taxonomy.TermSet",
            "_ObjectIdentity_": "2994929e-20f1-0000-2cdb-e577d70db169|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt",
            "CreatedDate": "\/Date(1536839573337)\/",
            "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/",
            "LastModifiedDate": "\/Date(1536840826883)\/",
            "Name": "PnP-CollabFooter-SharedLinks",
            "CustomProperties": {
              "_Sys_Nav_IsNavigationTermSet": "True"
            },
            "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
            "IsAvailableForTagging": true,
            "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Contact": "",
            "Description": "",
            "IsOpenForTermCreation": false,
            "Names": {
              "1033": "PnP-CollabFooter-SharedLinks>"
            },
            "Stakeholders": []
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'PnP-CollabFooter-SharedLinks>', termGroupName: termGroupName } });
    assert(loggerLogSpy.calledWith({
      "CreatedDate": "2018-09-13T11:52:53.337Z",
      "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed",
      "LastModifiedDate": "2018-09-13T12:13:46.883Z",
      "Name": "PnP-CollabFooter-SharedLinks",
      "CustomProperties": {
        "_Sys_Nav_IsNavigationTermSet": "True"
      },
      "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f",
      "IsAvailableForTagging": true,
      "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
      "Contact": "",
      "Description": "",
      "IsOpenForTermCreation": false,
      "Names": {
        "1033": "PnP-CollabFooter-SharedLinks>"
      },
      "Stakeholders": []
    }));
  });

  it('correctly handles term group not found via id', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
          "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
        }, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e"
      }
    ]));

    await assert.rejects(command.action(logger, {
      options: {
        id: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed',
        termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term group not found via name', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
          "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
        }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
      }
    ]));

    await assert.rejects(command.action(logger, {
      options: {
        name: name,
        termGroupName: termGroupName
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term set not found via id', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
          "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
        }, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d"
      }
    ]));

    await assert.rejects(command.action(logger, {
      options: {
        id: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed',
        termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb'
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term set not found via name', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
          "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
        }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
      }
    ]));

    await assert.rejects(command.action(logger, {
      options: {
        name: name,
        termGroupName: termGroupName
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles error when retrieving taxonomy term set', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
          "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
        }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
      }
    ]));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('File Not Found.'));
  });

  it('fails validation if neither id nor name specified', async () => {
    const actual = await command.validate({ options: { termGroupName: termGroupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name specified', async () => {
    const actual = await command.validate({ options: { id: id, name: name, termGroupName: termGroupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid', termGroupName: termGroupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither termGroupId nor termGroupName specified', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termGroupId and termGroupName specified', async () => {
    const actual = await command.validate({ options: { id: id, termGroupId: termGroupId, termGroupName: termGroupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: id, termGroupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and termGroupName specified', async () => {
    const actual = await command.validate({ options: { id: id, termGroupName: termGroupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and termGroupName specified', async () => {
    const actual = await command.validate({ options: { name: 'People', termGroupName: termGroupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when id and termGroupId specified', async () => {
    const actual = await command.validate({ options: { id: id, termGroupId: termGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and termGroupId specified', async () => {
    const actual = await command.validate({ options: { name: name, termGroupId: termGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when webUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: id, termGroupId: termGroupId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid url', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, termGroupId: termGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').rejects(new Error('getRequestDigest error'));

    await assert.rejects(command.action(logger, {
      options: {
        name: name,
        termGroupName: termGroupName
      }
    } as any), new CommandError('getRequestDigest error'));
  });
});