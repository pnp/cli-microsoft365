import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-list');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TERM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets taxonomy terms from term set by id, term group by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term set by name, term group by id (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, termSetName: 'PnPTermSets', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term set by id, term group by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{7a167c47-2b37-41d0-94d0-e962c1a4f2ed}</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term set by name, term group by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all properties for output json', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "10ca969e-3062-0000-2cdb-e38e5b6fba03" }, 70, { "IsNull": false }, 71, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 73, { "IsNull": false }, 74, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh/fzgFZGpUQ==" }, 76, { "IsNull": false }, 78, { "IsNull": false }, 79, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+s=" }, 81, { "IsNull": false }, 83, { "IsNull": false }, 84, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7" }, 86, { "IsNull": false }, 87, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [{ "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "/Date(1536839575320)/", "Id": "/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)/", "LastModifiedDate": "/Date(1536839575337)/", "Name": "HR", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "HR", "TermsCount": 0 }, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "/Date(1536839575477)/", "Id": "/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)/", "LastModifiedDate": "/Date(1536839575490)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD/1ASKE2ziDgfrY1GAg==", "CreatedDate": "/Date(1536839575600)/", "Id": "/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)/", "LastModifiedDate": "/Date(1536839575617)/", "Name": "Leadership", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Leadership", "TermsCount": 2 }] }]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "2018-09-13T11:52:55.320Z", "Id": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b", "LastModifiedDate": "2018-09-13T11:52:55.337Z", "Name": "HR", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "HR", "TermsCount": 0 }, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "2018-09-13T11:52:55.477Z", "Id": "247543b6-45f2-4232-b9e8-66c5bf53c31e", "LastModifiedDate": "2018-09-13T11:52:55.490Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD/1ASKE2ziDgfrY1GAg==", "CreatedDate": "2018-09-13T11:52:55.600Z", "Id": "ffc3608f-1250-4d28-b388-381fad8d4602", "LastModifiedDate": "2018-09-13T11:52:55.617Z", "Name": "Leadership", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Leadership", "TermsCount": 2 }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in term group name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in term set name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations&gt;</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "1e1e969e-7056-0000-2cdb-ea009f6c99c8"
          }, 70, {
            "IsNull": false
          }, 71, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 73, {
            "IsNull": false
          }, 74, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 76, {
            "IsNull": false
          }, 78, {
            "IsNull": false
          }, 79, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 81, {
            "IsNull": false
          }, 83, {
            "IsNull": false
          }, 84, {
            "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7"
          }, 86, {
            "IsNull": false
          }, 87, {
            "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7niHPAumMhU6sBKkTpEpdKw==", "CreatedDate": "\/Date(1536839575320)\/", "Id": "\/Guid(02cf219e-8ce9-4e85-ac04-a913a44a5d2b)\/", "LastModifiedDate": "\/Date(1536839575337)\/", "Name": "HR", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "HR", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7tkN1JPJFMkK56GbFv1PDHg==", "CreatedDate": "\/Date(1536839575477)\/", "Id": "\/Guid(247543b6-45f2-4232-b9e8-66c5bf53c31e)\/", "LastModifiedDate": "\/Date(1536839575490)\/", "Name": "IT", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "IT", "TermsCount": 0
              }, {
                "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1e1e969e-7056-0000-2cdb-ea009f6c99c8|fec14c62-7c3b-481b-851b-c80d7802b224:te:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7j2DD\u002f1ASKE2ziDgfrY1GAg==", "CreatedDate": "\/Date(1536839575600)\/", "Id": "\/Guid(ffc3608f-1250-4d28-b388-381fad8d4602)\/", "LastModifiedDate": "\/Date(1536839575617)\/", "Name": "Leadership", "CustomProperties": {

                }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {

                }, "MergedTermIds": [

                ], "PathOfTerm": "Leadership", "TermsCount": 0
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations>', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Id: "02cf219e-8ce9-4e85-ac04-a913a44a5d2b",
            Name: "HR"
          },
          {
            Id: "247543b6-45f2-4232-b9e8-66c5bf53c31e",
            Name: "IT"
          },
          {
            Id: "ffc3608f-1250-4d28-b388-381fad8d4602",
            Name: "Leadership"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term group not found via id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "8092929e-e06a-0000-2cdb-e217ce4a986e"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term group not found via name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-CollabFooter-SharedLinks', termGroupName: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term set not found via id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7192929e-70ad-0000-2cdb-e0f1f8d0326d"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termSetId: '7a167c47-2b37-41d0-94d0-e962c1a4f2ed', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term set not found via name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1218", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "7992929e-a0f1-0000-2cdb-e3c8b27b1f34"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-CollabFooter-SharedLinks', termGroupName: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving taxonomy terms', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no terms found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" /><Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="80" ParentId="77" Name="TermSets" /><Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">PnP-Organizations</Parameter></Parameters></Method><Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8126.1225", "ErrorInfo": null, "TraceCorrelationId": "10ca969e-3062-0000-2cdb-e38e5b6fba03" }, 70, { "IsNull": false }, 71, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 73, { "IsNull": false }, 74, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh/fzgFZGpUQ==" }, 76, { "IsNull": false }, 78, { "IsNull": false }, 79, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+s=" }, 81, { "IsNull": false }, 83, { "IsNull": false }, 84, { "_ObjectIdentity_": "10ca969e-3062-0000-2cdb-e38e5b6fba03|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+ts4nkUgBOoQZGDcrxallG7" }, 86, { "IsNull": false }, 87, { "_ObjectType_": "SP.Taxonomy.TermCollection", "_Child_Items_": [] }]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither termSetId nor termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupName: 'PnPTermSets' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termSetId and termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termSetName: 'PnP-CollabFooter-SharedLinks', termGroupName: 'PnPTermSets' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termSetId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: 'invalid', termGroupName: 'PnPTermSets' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither termGroupId nor termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termGroupId and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termGroupName: 'PnPTermSets' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'PnPTermSets' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetName and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetName: 'People', termGroupName: 'PnPTermSets' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetId and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when termSetName and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termSetName: 'PnP-CollabFooter-SharedLinks', termGroupId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('handles promise rejection', (done) => {
    Utils.restore((command as any).getRequestDigest);
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.reject('getRequestDigest error'));
    
    cmdInstance.action({
      options: { debug: false, termSetName: 'PnP-Organizations', termGroupName: 'PnPTermSets', output: 'json' }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('getRequestDigest error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});