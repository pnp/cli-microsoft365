import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
const command: Command = require('./userprofile-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.USERPROFILE_SET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  const spoUrl = 'https://contoso.sharepoint.com';

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = spoUrl;
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
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.USERPROFILE_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if the userName is not provided.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the propertyName is not provided.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyValue: 'Senior Developer'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the propertyValue is not provided.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when the input is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer'
      }
    });
    assert.equal(actual, true);
  });

  it('updates single valued profile property', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }
      return Promise.reject('Invalid request');
    });

    const body: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-JobTitle',
      'propertyValue': 'Senior Developer'
    };

    cmdInstance.action({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer',
        debug: true
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.equal(JSON.stringify(lastCall.body), JSON.stringify(body));
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('updates single valued profile property (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('updates multi valued profile property', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }
      return Promise.reject('Invalid request');
    });

    const body: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-Skills',
      'propertyValues': ['CSS', 'HTML']
    };

    cmdInstance.action({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-Skills',
        propertyValue: 'CSS, HTML'
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.equal(JSON.stringify(lastCall.body), JSON.stringify(body));
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('updates multi valued profile property (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-Skills',
        propertyValue: 'CSS, HTML',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('correctly handles error while updating profile property', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.USERPROFILE_SET));
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