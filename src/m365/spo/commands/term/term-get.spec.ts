import * as assert from 'assert';
import * as os from 'os';
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
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./term-get');

describe(commands.TERM_GET, () => {
  const termId = '334d792b-0026-4486-a593-a2031b564104';
  const termName = 'IT';
  const termGroupId = '5c928151-c140-4d48-aab9-54da901c7fef';
  const termGroupName = 'People';
  const termSetId = '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f';
  const termSetName = 'Department';

  //#region Mocked Responses
  const formattedResponse = {
    "CreatedDate": "2023-02-07T17:24:44.037Z",
    "Id": "334d792b-0026-4486-a593-a2031b564104",
    "LastModifiedDate": "2023-02-07T17:24:44.037Z",
    "Name": "Test Term",
    "CustomProperties": {},
    "CustomSortOrder": null,
    "IsAvailableForTagging": true,
    "Owner": "i:0#.f|membership|joe@contoso.com",
    "Description": "",
    "IsDeprecated": false,
    "IsKeyword": false,
    "IsPinned": false,
    "IsPinnedRoot": false,
    "IsReused": false,
    "IsRoot": true,
    "IsSourceTerm": true,
    "LocalCustomProperties": {},
    "MergedTermIds": [],
    "PathOfTerm": "Test Term",
    "TermsCount": 1
  };
  const csomResponseById = [{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.23325.12003", "ErrorInfo": null, "TraceCorrelationId": "8fff93a0-2037-6000-110c-b9eb6618f070" }, 14, { "IsNull": false }, 15, { "_ObjectIdentity_": "8fff93a0-2037-6000-110c-b9eb6618f070|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmtK3lNMyYAhkSlk6IDG1ZBBA==" }, 16, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "8fff93a0-2037-6000-110c-b9eb6618f070|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmtK3lNMyYAhkSlk6IDG1ZBBA==", "CreatedDate": "\/Date(1675790684037)\/", "Id": "\/Guid(334d792b-0026-4486-a593-a2031b564104)\/", "LastModifiedDate": "\/Date(1675790684037)\/", "Name": "Test Term", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|joe@contoso.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Test Term", "TermsCount": 1 }];
  const csomResponseByName = [{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.23325.12003", "ErrorInfo": null, "TraceCorrelationId": "47ff93a0-c06c-6000-110c-be1cdced2038" }, 2, { "IsNull": false }, 3, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 5, { "IsNull": false }, 6, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:st:kTm3XibpGUiE5nxBtVMTfw==" }, 8, { "IsNull": false }, 10, { "IsNull": false }, 11, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:gr:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfg=" }, 13, { "IsNull": false }, 15, { "IsNull": false }, 16, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:se:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmt" }, 18, { "IsNull": false }, 22, { "IsNull": false }, 23, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [{ "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmtK3lNMyYAhkSlk6IDG1ZBBA==", "CreatedDate": "\/Date(1675790684037)\/", "Id": "\/Guid(334d792b-0026-4486-a593-a2031b564104)\/", "LastModifiedDate": "\/Date(1675790684037)\/", "Name": "Test Term", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|joe@contoso.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Test Term", "TermsCount": 1 }] }];
  //#endregion

  let log: string[];
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
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TERM_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets taxonomy term by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${termId}}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(csomResponseById);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: termId } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('gets taxonomy term by name, term group by id, term set by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(csomResponseByName);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: termName, termGroupId: termGroupId, termSetId: termSetId } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('gets taxonomy term by name, term group by id, term set by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(termSetName)}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(csomResponseByName);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: termName, termGroupId: termGroupId, termSetName: termSetName } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('gets taxonomy term by name, term group by name, term set by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(termGroupName)}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(csomResponseByName);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: termName, termGroupName: termGroupName, termSetId: termSetId } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('gets taxonomy term by name, term group by name, term set by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(termGroupName)}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(termSetName)}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify(csomResponseByName);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: termName, termGroupName: termGroupName, termSetName: termSetName } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('correctly handles term not found by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.23325.12003", "ErrorInfo": null, "TraceCorrelationId": "47ff93a0-c06c-6000-110c-be1cdced2038" }, 2, { "IsNull": false }, 3, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 5, { "IsNull": false }, 6, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:st:kTm3XibpGUiE5nxBtVMTfw==" }, 8, { "IsNull": false }, 10, { "IsNull": false }, 11, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:gr:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfg=" }, 13, { "IsNull": false }, 15, { "IsNull": false }, 16, { "_ObjectIdentity_": "47ff93a0-c06c-6000-110c-be1cdced2038|fec14c62-7c3b-481b-851b-c80d7802b224:se:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmt" }, 18, { "IsNull": false }, 22, { "IsNull": false }, 23, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [] }]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: termName,
        termGroupId: termGroupId,
        termSetId: termSetId
      }
    } as any), new CommandError(`Term with name '${termName}' could not be found.`));
  });

  it('correctly handles multiple terms not found by name', async () => {
    const childItems = [{ "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmtEP0Wk4LJvk+y0GrLwtClew==", "CreatedDate": "\/Date(1675790717780)\/", "Id": "\/Guid(9316fd10-c982-4fbe-b2d0-6acbc2d0a57b)\/", "LastModifiedDate": "\/Date(1675790717780)\/", "Name": "Test Child Term", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|joe@contoso.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": false, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Test Term;Test Child Term", "TermsCount": 0 }, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmtNXuY\u002fJxmJ0m9jOekcLeh2w==", "CreatedDate": "\/Date(1675795608853)\/", "Id": "\/Guid(fc987b35-669c-4927-bd8c-e7a470b7a1db)\/", "LastModifiedDate": "\/Date(1675795608853)\/", "Name": "Test Child Term", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|joe@contoso.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Test Child Term", "TermsCount": 0 }];
    const disambiguationText = childItems.map(c => {
      return `- ${c.Id.replace('/Guid(', '').replace(')/', '')} - ${c.PathOfTerm}`;
    }).join(os.EOL);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.23325.12003", "ErrorInfo": null, "TraceCorrelationId": "b50094a0-80a4-6000-110c-b074a0d4c336" }, 2, { "IsNull": false }, 3, { "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 5, { "IsNull": false }, 6, { "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:st:kTm3XibpGUiE5nxBtVMTfw==" }, 8, { "IsNull": false }, 10, { "IsNull": false }, 11, { "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:gr:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfg=" }, 13, { "IsNull": false }, 15, { "IsNull": false }, 16, { "_ObjectIdentity_": "b50094a0-80a4-6000-110c-b074a0d4c336|fec14c62-7c3b-481b-851b-c80d7802b224:se:kTm3XibpGUiE5nxBtVMTf25aOnte4ElDn7uvWBPvXfjuQ1jPsltwT78ny15SLpmt" }, 18, { "IsNull": false }, 22, { "IsNull": false }, 23, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": childItems }]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: termName,
        termGroupId: termGroupId,
        termSetId: termSetId
      }
    } as any), new CommandError(`Multiple terms with the specific term name found. Please disambiguate:${os.EOL}${disambiguationText}`));
  });

  it('correctly handles term not found by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${termId}}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.23318.12003", "ErrorInfo": null, "TraceCorrelationId": "b5ff93a0-b05f-6000-2394-3fb22b8af563" }, 14, { "IsNull": true }, 15, null, 16, null]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: termId } } as any),
      new CommandError(`Term with id '${termId}' could not be found.`));
  });

  it('correctly handles term group not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": { "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException" }, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b" }]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: termName,
        termGroupId: termGroupId,
        termSetId: termSetId
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('correctly handles term set not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(termName)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="GetById"><Parameters><Parameter Type="String">${termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="GetById"><Parameters><Parameter Type="String">${termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": { "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException" }, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b" }]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: termName,
        termGroupId: termGroupId,
        termSetId: termSetId
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index'));
  });

  it('escapes the xml when special characters are passed', async () => {
    const weirdCharacterTermName = 'IT>';
    const weirdCharacterTermGroupName = 'Department>';
    const weirdCharacterTermSetName = 'People>';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery' &&
        opts.headers &&
        opts.headers['X-RequestDigest']) {
        return JSON.stringify(csomResponseByName);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: weirdCharacterTermName, termGroupName: weirdCharacterTermGroupName, termSetName: weirdCharacterTermSetName } });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only id specified', async () => {
    const actual = await command.validate({ options: { id: termId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when only name specified', async () => {
    const actual = await command.validate({ options: { name: termName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: termName, termGroupId: 'invalid', termSetName: termSetName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termSetId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: termName, termSetId: 'invalid', termGroupName: termGroupId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name, termGroupName and termSetName specified', async () => {
    const actual = await command.validate({ options: { name: termName, termGroupName: termGroupName, termSetName: termSetName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name, termGroupId and termSetId specified', async () => {
    const actual = await command.validate({ options: { name: termName, termGroupId: termGroupId, termSetId: termSetId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').callsFake(async () => { throw 'getRequestDigest error'; });

    await assert.rejects(command.action(logger, { options: { name: termName, termGroupName: termGroupName, termSetName: termSetName } } as any), new CommandError('getRequestDigest error'));
  });

});
