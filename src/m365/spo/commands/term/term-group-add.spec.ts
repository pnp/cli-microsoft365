import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-group-add');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_GROUP_ADD, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_GROUP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds term group by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
          "Description": ""
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term group by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8' } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a8",
          "Description": ""
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term group by name with description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets', description: 'Term sets for PnP' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
          "Description": "Term sets for PnP"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term group by id with description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8', description: 'Term sets for PnP' } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a8",
          "Description": "Term sets for PnP"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving the term store', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified name already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "Group names must be unique.", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Group names must be unique.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified id already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.", "ErrorValue": null, "TraceCorrelationId": "3f4b919e-5077-0000-29c7-0c2eabf41bf3", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Taxonomy.TermStoreOperationException"
              }, "TraceCorrelationId": "3f4b919e-5077-0000-29c7-0c2eabf41bf3"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when setting the description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8', description: 'Term sets for PnP' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term group name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets>", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets>",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
          "Description": ""
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term group description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP&gt;</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnPTermSets', description: 'Term sets for PnP>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Name": "PnPTermSets",
          "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
          "Description": "Term sets for PnP>"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnPTermSets', id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnPTermSets', id: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'People' } });
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
});