import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-group-list');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_GROUP_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists taxonomy term groups (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1215",
            "ErrorInfo": null,
            "TraceCorrelationId": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9"
          },
          4,
          {
            "IsNull": false
          },
          5,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          7,
          {
            "IsNull": false
          },
          8,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          10,
          {
            "IsNull": false
          },
          11,
          {
            "_ObjectType_": "SP.Taxonomy.TermGroupCollection",
            "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8=",
                "CreatedDate": "\/Date(1529479401033)\/",
                "Id": "\/Guid(36a62501-17ea-455a-bed4-eff862242def)\/",
                "LastModifiedDate": "\/Date(1529479401033)\/",
                "Name": "People",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s=",
                "CreatedDate": "\/Date(1536839573117)\/",
                "Id": "\/Guid(0e8f395e-ff58-4d45-9ff7-e331ab728beb)\/",
                "LastModifiedDate": "\/Date(1536839573117)\/",
                "Name": "PnPTermSets",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUTdqe9gByDZKkEZiltR3nIc=",
                "CreatedDate": "\/Date(1529479401063)\/",
                "Id": "\/Guid(d87b6a37-c801-4a36-9046-6296d4779c87)\/",
                "LastModifiedDate": "\/Date(1529479401063)\/",
                "Name": "Search Dictionaries",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUdrlarEXoGtNuzIB3A5zZDo=",
                "CreatedDate": "\/Date(1529479400770)\/",
                "Id": "\/Guid(b16ae5da-a017-4d6b-bb32-01dc0e73643a)\/",
                "LastModifiedDate": "\/Date(1529479400770)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-search",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQZhmdVzct1Fj6MAalJ1aHc=",
                "CreatedDate": "\/Date(1529495406027)\/",
                "Id": "\/Guid(d5996106-7273-45dd-8fa3-006a52756877)\/",
                "LastModifiedDate": "\/Date(1529495406027)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-Analytics",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUeAa0tV1fe9PpxZBXc21aYc=",
                "CreatedDate": "\/Date(1536754831887)\/",
                "Id": "\/Guid(d5d21ae0-7d75-4fef-a716-415dcdb56987)\/",
                "LastModifiedDate": "\/Date(1536754831887)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-hr",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUVSux4Ka74dLrn8bmCVuTp0=",
                "CreatedDate": "\/Date(1536754843060)\/",
                "Id": "\/Guid(82c7ae54-ef9a-4b87-ae7f-1b98256e4e9d)\/",
                "LastModifiedDate": "\/Date(1536754843060)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-Marketing",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpURC8Oohu2K5FoLzWJkLCzM0=",
                "CreatedDate": "\/Date(1536754304210)\/",
                "Id": "\/Guid(883abc10-d86e-45ae-a0bc-d62642c2cccd)\/",
                "LastModifiedDate": "\/Date(1536754304210)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-portal",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUYWJl\u002fqvH5hPrfM1Rk4nNTU=",
                "CreatedDate": "\/Date(1529479155453)\/",
                "Id": "\/Guid(fa978985-1faf-4f98-adf3-35464e273535)\/",
                "LastModifiedDate": "\/Date(1529479155453)\/",
                "Name": "System",
                "Description": "These term sets are used by the system itself.",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": true
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "Id": "36a62501-17ea-455a-bed4-eff862242def", "Name": "People" },
          { "Id": "0e8f395e-ff58-4d45-9ff7-e331ab728beb", "Name": "PnPTermSets" },
          { "Id": "d87b6a37-c801-4a36-9046-6296d4779c87", "Name": "Search Dictionaries" },
          { "Id": "b16ae5da-a017-4d6b-bb32-01dc0e73643a", "Name": "Site Collection - m365x035040.sharepoint.com-search" },
          { "Id": "d5996106-7273-45dd-8fa3-006a52756877", "Name": "Site Collection - m365x035040.sharepoint.com-sites-Analytics" },
          { "Id": "d5d21ae0-7d75-4fef-a716-415dcdb56987", "Name": "Site Collection - m365x035040.sharepoint.com-sites-hr" },
          { "Id": "82c7ae54-ef9a-4b87-ae7f-1b98256e4e9d", "Name": "Site Collection - m365x035040.sharepoint.com-sites-Marketing" },
          { "Id": "883abc10-d86e-45ae-a0bc-d62642c2cccd", "Name": "Site Collection - m365x035040.sharepoint.com-sites-portal" },
          { "Id": "fa978985-1faf-4f98-adf3-35464e273535", "Name": "System" }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists taxonomy term groups with all properties when output is JSON', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1215",
            "ErrorInfo": null,
            "TraceCorrelationId": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9"
          },
          4,
          {
            "IsNull": false
          },
          5,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          7,
          {
            "IsNull": false
          },
          8,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          10,
          {
            "IsNull": false
          },
          11,
          {
            "_ObjectType_": "SP.Taxonomy.TermGroupCollection",
            "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8=",
                "CreatedDate": "\/Date(1529479401033)\/",
                "Id": "\/Guid(36a62501-17ea-455a-bed4-eff862242def)\/",
                "LastModifiedDate": "\/Date(1529479401033)\/",
                "Name": "People",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s=",
                "CreatedDate": "\/Date(1536839573117)\/",
                "Id": "\/Guid(0e8f395e-ff58-4d45-9ff7-e331ab728beb)\/",
                "LastModifiedDate": "\/Date(1536839573117)\/",
                "Name": "PnPTermSets",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUTdqe9gByDZKkEZiltR3nIc=",
                "CreatedDate": "\/Date(1529479401063)\/",
                "Id": "\/Guid(d87b6a37-c801-4a36-9046-6296d4779c87)\/",
                "LastModifiedDate": "\/Date(1529479401063)\/",
                "Name": "Search Dictionaries",
                "Description": "",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUdrlarEXoGtNuzIB3A5zZDo=",
                "CreatedDate": "\/Date(1529479400770)\/",
                "Id": "\/Guid(b16ae5da-a017-4d6b-bb32-01dc0e73643a)\/",
                "LastModifiedDate": "\/Date(1529479400770)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-search",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQZhmdVzct1Fj6MAalJ1aHc=",
                "CreatedDate": "\/Date(1529495406027)\/",
                "Id": "\/Guid(d5996106-7273-45dd-8fa3-006a52756877)\/",
                "LastModifiedDate": "\/Date(1529495406027)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-Analytics",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUeAa0tV1fe9PpxZBXc21aYc=",
                "CreatedDate": "\/Date(1536754831887)\/",
                "Id": "\/Guid(d5d21ae0-7d75-4fef-a716-415dcdb56987)\/",
                "LastModifiedDate": "\/Date(1536754831887)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-hr",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUVSux4Ka74dLrn8bmCVuTp0=",
                "CreatedDate": "\/Date(1536754843060)\/",
                "Id": "\/Guid(82c7ae54-ef9a-4b87-ae7f-1b98256e4e9d)\/",
                "LastModifiedDate": "\/Date(1536754843060)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-Marketing",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpURC8Oohu2K5FoLzWJkLCzM0=",
                "CreatedDate": "\/Date(1536754304210)\/",
                "Id": "\/Guid(883abc10-d86e-45ae-a0bc-d62642c2cccd)\/",
                "LastModifiedDate": "\/Date(1536754304210)\/",
                "Name": "Site Collection - m365x035040.sharepoint.com-sites-portal",
                "Description": "",
                "IsSiteCollectionGroup": true,
                "IsSystemGroup": false
              },
              {
                "_ObjectType_": "SP.Taxonomy.TermGroup",
                "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUYWJl\u002fqvH5hPrfM1Rk4nNTU=",
                "CreatedDate": "\/Date(1529479155453)\/",
                "Id": "\/Guid(fa978985-1faf-4f98-adf3-35464e273535)\/",
                "LastModifiedDate": "\/Date(1529479155453)\/",
                "Name": "System",
                "Description": "These term sets are used by the system itself.",
                "IsSiteCollectionGroup": false,
                "IsSystemGroup": true
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUQElpjbqF1pFvtTv+GIkLe8=", "CreatedDate": "2018-06-20T07:23:21.033Z", "Id": "36a62501-17ea-455a-bed4-eff862242def", "LastModifiedDate": "2018-06-20T07:23:21.033Z", "Name": "People", "Description": "", "IsSiteCollectionGroup": false, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUV45jw5Y/0VNn/fjMatyi+s=", "CreatedDate": "2018-09-13T11:52:53.117Z", "Id": "0e8f395e-ff58-4d45-9ff7-e331ab728beb", "LastModifiedDate": "2018-09-13T11:52:53.117Z", "Name": "PnPTermSets", "Description": "", "IsSiteCollectionGroup": false, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUTdqe9gByDZKkEZiltR3nIc=", "CreatedDate": "2018-06-20T07:23:21.063Z", "Id": "d87b6a37-c801-4a36-9046-6296d4779c87", "LastModifiedDate": "2018-06-20T07:23:21.063Z", "Name": "Search Dictionaries", "Description": "", "IsSiteCollectionGroup": false, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUdrlarEXoGtNuzIB3A5zZDo=", "CreatedDate": "2018-06-20T07:23:20.770Z", "Id": "b16ae5da-a017-4d6b-bb32-01dc0e73643a", "LastModifiedDate": "2018-06-20T07:23:20.770Z", "Name": "Site Collection - m365x035040.sharepoint.com-search", "Description": "", "IsSiteCollectionGroup": true, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUQZhmdVzct1Fj6MAalJ1aHc=", "CreatedDate": "2018-06-20T11:50:06.027Z", "Id": "d5996106-7273-45dd-8fa3-006a52756877", "LastModifiedDate": "2018-06-20T11:50:06.027Z", "Name": "Site Collection - m365x035040.sharepoint.com-sites-Analytics", "Description": "", "IsSiteCollectionGroup": true, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUeAa0tV1fe9PpxZBXc21aYc=", "CreatedDate": "2018-09-12T12:20:31.887Z", "Id": "d5d21ae0-7d75-4fef-a716-415dcdb56987", "LastModifiedDate": "2018-09-12T12:20:31.887Z", "Name": "Site Collection - m365x035040.sharepoint.com-sites-hr", "Description": "", "IsSiteCollectionGroup": true, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUVSux4Ka74dLrn8bmCVuTp0=", "CreatedDate": "2018-09-12T12:20:43.060Z", "Id": "82c7ae54-ef9a-4b87-ae7f-1b98256e4e9d", "LastModifiedDate": "2018-09-12T12:20:43.060Z", "Name": "Site Collection - m365x035040.sharepoint.com-sites-Marketing", "Description": "", "IsSiteCollectionGroup": true, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpURC8Oohu2K5FoLzWJkLCzM0=", "CreatedDate": "2018-09-12T12:11:44.210Z", "Id": "883abc10-d86e-45ae-a0bc-d62642c2cccd", "LastModifiedDate": "2018-09-12T12:11:44.210Z", "Name": "Site Collection - m365x035040.sharepoint.com-sites-portal", "Description": "", "IsSiteCollectionGroup": true, "IsSystemGroup": false }, { "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "dfa8909e-402d-0000-2cdb-e7b0f4389f1c|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh/fzgFZGpUYWJl/qvH5hPrfM1Rk4nNTU=", "CreatedDate": "2018-06-20T07:19:15.453Z", "Id": "fa978985-1faf-4f98-adf3-35464e273535", "LastModifiedDate": "2018-06-20T07:19:15.453Z", "Name": "System", "Description": "These term sets are used by the system itself.", "IsSiteCollectionGroup": false, "IsSystemGroup": true }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no term groups found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1215",
            "ErrorInfo": null,
            "TraceCorrelationId": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9"
          },
          4,
          {
            "IsNull": false
          },
          5,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          7,
          {
            "IsNull": false
          },
          8,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          10,
          {
            "IsNull": false
          },
          11,
          {
            "_ObjectType_": "SP.Taxonomy.TermGroupCollection",
            "_Child_Items_": []
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving taxonomy term groups', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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
      options: { debug: false }
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