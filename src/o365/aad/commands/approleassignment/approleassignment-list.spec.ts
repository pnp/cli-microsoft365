import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./approleassignment-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as TestRequests from './approleassignment-list-requeststub.spec';
import * as TestConstants from './approleassignment-list-constants.spec';

describe(commands.APPROLEASSIGNMENT_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let textOutput = [
    {
      "resourceDisplayName": "Microsoft Graph",
      "roleName": "User.Read.All"
    },
    {
      "resourceDisplayName": "Contoso Product Catalog service",
      "roleName": "access_as_application"
    }
  ];
  let jsonOutput = [
    {
      "appRoleId": "df021288-bdef-4463-88db-98f22de89214",
      "resourceDisplayName": "Microsoft Graph",
      "resourceId": "b1ce2d04-5502-4142-ba53-819327b74b5b",
      "roleId": "df021288-bdef-4463-88db-98f22de89214",
      "roleName": "User.Read.All"
    },
    {
      "appRoleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "resourceDisplayName": "Contoso Product Catalog service",
      "resourceId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f",
      "roleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "roleName": "access_as_application"
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
      request.get
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
    assert.equal(command.name.startsWith(commands.APPROLEASSIGNMENT_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('retrieves App Role assignments for the specified displayName', (done) => {
    sinon.stub(request, 'get').callsFake(TestRequests.requestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'json', displayName: TestConstants.CommandActionParameters.appNameWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('retrieves App Role assignments for the specified appId', (done) => {
    sinon.stub(request, 'get').callsFake(TestRequests.requestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'json', appId: TestConstants.CommandActionParameters.appIdWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves App Role assignments for the specified appId and outputs text', (done) => {
    sinon.stub(request, 'get').callsFake(TestRequests.requestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'text', appId: TestConstants.CommandActionParameters.appIdWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles an appId that does not exist', (done) => {
    sinon.stub(request, 'get').callsFake(TestRequests.requestStub.retrieveAppRoles);

    cmdInstance.action({ options: { appId: TestConstants.CommandActionParameters.invalidAppId } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('app registration not found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('correctly handles no app role assignments for the specified app', (done) => {
    sinon.stub(request, 'get').callsFake(TestRequests.requestStub.retrieveAppRoles);

    cmdInstance.action({ options: { appId: TestConstants.CommandActionParameters.appIdWithNoRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('app registration not found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, appId: '36e3a540-6f25-4483-9542-9f5fa00bb633' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither appId nor displayName are not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: TestConstants.CommandActionParameters.appIdWithNoRoleAssignments, displayName: TestConstants.CommandActionParameters.appNameWithRoleAssignments } });
    assert.notEqual(actual, true);
  })

  it('passes validation when the appId option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: TestConstants.CommandActionParameters.appIdWithNoRoleAssignments } });
    assert.equal(actual, true);
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

  it('supports specifying appId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.APPROLEASSIGNMENT_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});

