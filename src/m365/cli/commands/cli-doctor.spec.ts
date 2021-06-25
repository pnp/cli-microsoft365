import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Logger } from '../../../cli';
import Command from '../../../Command';
import Utils from '../../../Utils';
import commands from '../commands';
import auth from '../../../Auth';
import * as os from 'os';
const packageJSON = require('../../../../package.json');

const command: Command = require('./cli-doctor');

describe(commands.DOCTOR, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
  });

  afterEach(() => {
    Utils.restore([
      Utils.getRolesFromAccessToken,
      Utils.getScopesFromAccessToken,
      os.platform,
      os.version,
      os.release,
      process.env
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
    assert.strictEqual(command.name.startsWith(commands.DOCTOR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fetches roles and scopes from access token', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });

    const getRolesFromAccessTokenSpy = sinon.spy(Utils,'getRolesFromAccessToken');
    const getScopesFromAccessTokenSpy = sinon.spy(Utils,'getScopesFromAccessToken');

    command.action(logger,{options:{}},() => {
      try{
        assert(getRolesFromAccessTokenSpy.called && getScopesFromAccessTokenSpy.called);
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('fetches roles and scopes from multiple access tokens', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."},
      "https://mydev.sharepoint.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."},
      "https://something.sharepoint.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    
    const getRolesFromAccessTokenSpy = sinon.spy(Utils,'getRolesFromAccessToken');
    const getScopesFromAccessTokenSpy = sinon.spy(Utils,'getScopesFromAccessToken');
    
    command.action(logger,{options:{}},() => {
      try{
        assert(getRolesFromAccessTokenSpy.calledThrice && getScopesFromAccessTokenSpy.calledThrice);
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves "scopes" in the diagnostic information about the current environment', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(0);
    
    sinon.stub(Utils,'getRolesFromAccessToken').returns([]);
    sinon.stub(Utils,'getScopesFromAccessToken').returns(["AllSites.FullControl","AppCatalog.ReadWrite.All"]);
  
    command.action(logger,{options:{}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "DeviceCode",
          CliAadAppId:"31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant:"common",
          CliEnvironment:"",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: [],
          Scopes: ["AllSites.FullControl","AppCatalog.ReadWrite.All"],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves "roles" in the diagnostic information about the current environment', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(0);

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
    
  
    command.action(logger,{options:{}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "DeviceCode",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: [],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves "roles" and "scopes" in the diagnostic information about the current environment', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(0);

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });
  
    command.action(logger,{options:{}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "DeviceCode",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: ["Sites.Read.All","Files.ReadWrite.All"],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves diagnostic information about the current environment when there are no roles or scopes available', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(0);

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return [];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
  
    command.action(logger,{options:{}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "DeviceCode",
          CliAadAppId:"31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant:"common",
          CliEnvironment:"",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: [],
          Scopes: [],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves diagnostic information about the current environment with auth type "Certificate"', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
 
    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });

    command.action(logger,{options:{}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: [],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves tenant information as "single" when TenantID is a GUID', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("923d42f0-6d23-41eb-b68d-c036d242654f");
    sinon.stub(auth.service,'authType').value(2);

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
  
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "single",
          CliEnvironment:"",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: [],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves diagnostic information about the current environment (debug)', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
  
    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
  
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId:"31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant:"common",
          CliEnvironment :"",
          CliVersion     : "3.11.0",
          NodeVersion    : "v14.17.0",
          OS             : {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles          : ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes         : [],
          Shell          : ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves diagnostic information of the current environment when executing in docker', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': 'docker' });

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
  
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "docker",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: [],
          Shell: ""
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });

  it('retrieves the "shell" in diagnostic information of the current environment', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"eyJ0..."}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
    sinon.stub(process, 'env').value({ 'SHELL': 'bash' });

    sinon.stub(Utils,'getRolesFromAccessToken').callsFake(()=>{
      return ["Sites.Read.All","Files.ReadWrite.All"];
    });

    sinon.stub(Utils,'getScopesFromAccessToken').callsFake(()=>{
      return [];
    });
  
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: ["Sites.Read.All","Files.ReadWrite.All"],
          Scopes: [],
          Shell: "bash"
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });



  it('returns empty roles and scopes in diagnostic information when access token is empty', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":""}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
    sinon.stub(process, 'env').value({ 'SHELL': 'bash' });
    
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: [],
          Scopes: [],
          Shell: "bash"
        }));
        done();
      }
      catch(e){
        done(e);
      }
    });
  });


  it('returns empty roles and scopes in diagnostic information when access token is invalid', (done) => {
    sinon.stub(auth.service,'accessTokens').value({
      "https://graph.microsoft.com":{"expiresOn":"2021-07-04T09:52:18.000Z","accessToken":"a.b.c.d"}
    });
    sinon.stub(os,'platform').returns("win32");
    sinon.stub(os,'version').returns("Windows 10 Pro");
    sinon.stub(os,'release').returns("10.0.19043");
    sinon.stub(packageJSON,'version').value("3.11.0");
    sinon.stub(process,'version').value("v14.17.0");
    sinon.stub(auth.service,'appId').value("31359c7f-bd7e-475c-86db-fdb8c937548e");
    sinon.stub(auth.service,'tenant').value("common");
    sinon.stub(auth.service,'authType').value(2);
    sinon.stub(process, 'env').value({ 'SHELL': 'bash' });
  
    command.action(logger,{options:{debug:true}},() => {
      try{
        assert(loggerLogSpy.calledWith({
          AuthMode: "Certificate",
          CliAadAppId: "31359c7f-bd7e-475c-86db-fdb8c937548e",
          CliAadAppTenant: "common",
          CliEnvironment: "",
          CliVersion: "3.11.0",
          NodeVersion: "v14.17.0",
          OS: {"Platform":"win32","Version":"Windows 10 Pro","Release":"10.0.19043"},
          Roles: [],
          Scopes: [],
          Shell: "bash"
        }));
        done();
      }
      catch(e){
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