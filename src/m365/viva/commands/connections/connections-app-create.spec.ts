import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import * as Sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./connections-app-create');

const admZipMock = {
  // we need these unused params so that they can be properly mocked with sinon
  /* eslint-disable @typescript-eslint/no-unused-vars */
  addFile: (entryName: string, data: Buffer, comment?: string, attr?: number) => { },
  addLocalFile: (localPath: string, zipPath?: string, zipName?: string) => { },
  writeZip: (targetFileName?: string, callback?: (error: Error | null) => void) => { }
  /* eslint-enable @typescript-eslint/no-unused-vars */
};

describe(commands.CONNECTIONS_APP_CREATE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: Sinon.SinonSpy;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    (command as any).archive = admZipMock;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      Cli.executeCommandWithOutput,
      admZipMock.addFile,
      admZipMock.addLocalFile,
      admZipMock.writeZip
    ]);
  });

  after(() => {
    (command as any).archive = undefined;
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONNECTIONS_APP_CREATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app package for the specified communication site (root site)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify({
        Configuration: 0,
        WebTemplate: 'SITEPAGEPUBLISHING'
      }),
      stderr: ''
    }));
    const admZipMockAddFileSpy = sinon.spy(admZipMock, 'addFile');
    const admZipMockAddLocalFileSpy = sinon.spy(admZipMock, 'addLocalFile');
    const admZipMockWriteZipSpy = sinon.spy(admZipMock, 'writeZip');

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err);
        assert(admZipMockAddFileSpy.calledWith('manifest.json'), 'manifest not added to the zip');
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-color.png'), undefined, 'icon-color.png'));
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-outline.png'), undefined, 'icon-outline.png'));
        assert(admZipMockWriteZipSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app package for the specified communication site (/sites)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify({
        Configuration: 0,
        WebTemplate: 'SITEPAGEPUBLISHING'
      }),
      stderr: ''
    }));
    const admZipMockAddFileSpy = sinon.spy(admZipMock, 'addFile');
    const admZipMockAddLocalFileSpy = sinon.spy(admZipMock, 'addLocalFile');
    const admZipMockWriteZipSpy = sinon.spy(admZipMock, 'writeZip');

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com/sites/contoso',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png',
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err);
        assert(admZipMockAddFileSpy.calledWith('manifest.json'), 'manifest not added to the zip');
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-color.png'), undefined, 'icon-color.png'));
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-outline.png'), undefined, 'icon-outline.png'));
        assert(admZipMockWriteZipSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app package for the specified communication site (/teams + query string)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify({
        Configuration: 0,
        WebTemplate: 'SITEPAGEPUBLISHING'
      }),
      stderr: ''
    }));
    const admZipMockAddFileSpy = sinon.spy(admZipMock, 'addFile');
    const admZipMockAddLocalFileSpy = sinon.spy(admZipMock, 'addLocalFile');
    const admZipMockWriteZipSpy = sinon.spy(admZipMock, 'writeZip');

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com/teams/contoso?param=value',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err);
        assert(admZipMockAddFileSpy.calledWith('manifest.json'), 'manifest not added to the zip');
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-color.png'), undefined, 'icon-color.png'));
        assert(admZipMockAddLocalFileSpy.calledWithExactly(path.resolve('icon-outline.png'), undefined, 'icon-outline.png'));
        assert(admZipMockWriteZipSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails with an error if the specified site doesn't exist`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.reject({
      error: '404 - FILE NOT FOUND',
      stderr: '404 - FILE NOT FOUND'
    }));
    const admZipMockWriteZipSpy = sinon.spy(admZipMock, 'writeZip');

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, '404 - FILE NOT FOUND');
        assert(admZipMockWriteZipSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails with an error if the specified site doesn't exist (debug)`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.reject({
      error: '404 - FILE NOT FOUND',
      stderr: '404 - FILE NOT FOUND stderr'
    }));

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png',
        debug: true
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('404 - FILE NOT FOUND stderr'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with an error if the specified site is not a communication site', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify({
        Configuration: 0,
        WebTemplate: 'TEAM'
      }),
      stderr: ''
    }));

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com/sites/contoso',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png',
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, 'Site https://contoso.sharepoint.com/sites/contoso is not a Communication Site. Please specify a different site and try again.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails with an error if creating the zip file failed`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify({
        Configuration: 0,
        WebTemplate: 'SITEPAGEPUBLISHING'
      }),
      stderr: ''
    }));
    sinon.stub(admZipMock, 'writeZip').callsFake(() => {
      throw new Error('An error has occurred');
    });

    command.action(logger, {
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails validation if the specified app name is longer than 30 chars`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: `Stay on top of what's happening at Contoso`,
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if the specified description is longer than 80 chars`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: `Stay on top of what's happening at Contoso Stay on top of what's happening at Contoso`,
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if the specified long description is longer than 4000 chars`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `

      Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque magna turpis, sollicitudin vitae dui non, rutrum tincidunt ipsum. Vestibulum finibus, lectus vel fermentum pretium, urna lectus fermentum nulla, eu condimentum lectus justo in elit. Cras et pretium nibh. Suspendisse et placerat enim, a convallis odio. Donec elementum efficitur leo, quis semper nisi venenatis sit amet. Integer pellentesque tellus sit amet mattis cursus. Vivamus at viverra elit, vel lobortis sem.
      
      Integer rutrum efficitur rutrum. Nam malesuada malesuada purus quis viverra. In ac mauris quis tortor pellentesque convallis. Integer vel posuere leo. Proin fermentum nunc eget turpis feugiat vestibulum. Quisque ac urna est. Quisque porttitor, nunc id efficitur maximus, purus enim molestie velit, in molestie ante orci non tellus. Fusce feugiat, velit ut feugiat aliquet, justo odio pretium nulla, vel auctor elit leo elementum lorem.
      
      Sed vitae elit viverra, dignissim enim et, pretium nisi. Donec vel orci quam. Nulla a ex velit. Ut id diam a elit gravida eleifend at et erat. Sed eu lectus a libero sodales pharetra aliquet at tortor. Donec condimentum sed nulla venenatis sollicitudin. Sed ac lacus at sapien placerat porttitor sed sit amet risus. Fusce ac velit risus. Curabitur et metus tellus. Aenean ac molestie dolor. Nulla facilisi. Aliquam eu cursus metus, quis tincidunt justo. Vivamus consectetur ultrices lorem.
      
      Etiam nec ultricies nulla, et iaculis tellus. Fusce convallis et dolor sed rhoncus. Pellentesque nulla tortor, rhoncus eu nibh et, molestie rhoncus leo. Ut sit amet mattis nisi. Ut accumsan placerat ipsum sed tincidunt. Fusce ut efficitur enim. Aenean ornare quis sapien vitae eleifend. Praesent ultrices sed ex ut placerat.
      
      Sed ex massa, eleifend eget orci in, elementum auctor metus. Nulla at augue consectetur, luctus nulla sed, dictum nunc. Maecenas vel est laoreet, lobortis nibh in, mollis sem. Donec hendrerit dolor et velit efficitur feugiat. Fusce tristique, ex a porta luctus, orci mi sodales nibh, ac semper metus lacus gravida felis. Suspendisse in fringilla dui. Phasellus id nunc at orci sollicitudin laoreet ut in mi. Vestibulum leo ex, mattis vel augue eget, mattis efficitur magna.
      
      Integer et diam ipsum. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec pharetra pharetra lacinia. Aenean vitae ex ac nisi pharetra semper. Nam bibendum lobortis quam, nec venenatis metus efficitur at. Nunc dictum nulla eu tincidunt semper. Suspendisse porta, nisi quis cursus rhoncus, justo arcu dapibus elit, ut euismod odio sem in quam. Morbi quis est mauris. Sed laoreet vel velit eget convallis. Proin porta sed quam vitae volutpat. Aenean scelerisque dui sem, a ullamcorper magna ultrices sit amet.
      
      Nunc vehicula quis lectus sed tristique. Nullam consequat auctor libero vel mattis. Praesent dapibus ornare faucibus. Proin id viverra eros. Nunc diam dui, aliquam sed nisi id, faucibus semper orci. Quisque lacinia purus non porta sollicitudin. Nullam sit amet eros interdum, pharetra tellus vel, auctor sapien. Suspendisse et augue imperdiet ante pellentesque bibendum eu vel arcu. Etiam arcu nulla, finibus vitae porta vitae, tempus nec sapien. Nunc vitae aliquam nunc. Proin nec congue dolor, eu congue tortor. Mauris sed turpis sed mauris fringilla faucibus. Integer neque libero, venenatis quis fringilla commodo, tempus quis leo. Maecenas rhoncus tellus et molestie iaculis. Nulla quis feugiat nibh, maximus imperdiet enim. Nam congue a justo quis blandit.
      
      Integer dignissim vitae leo vel sagittis. Vivamus interdum, ipsum sed dictum aliquam, est nisl euismod nisl, vel luctus tellus nibh a ante. Curabitur posuere sapien a ullamcorper pharetra. Etiam consectetur, nunc vitae ullamcorper consequat, enim quam vulputate diam, non tempus mauris ligula quis justo. Fusce porta dui dignissim mauris ullamcorper mollis. Aliquam eget tempus libero. Nam eget purus sit amet lacus commodo commodo. Cras faucibus tortor vel odio varius, nec dignissim sapien commodo. Ut a lacus eu donec. `,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if a file with the app name already exists and no force flag specified`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if the specified colored icon doesn't exist`, () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      const p = path.toString();
      if (p.indexOf('.zip') > -1) {
        return false;
      }
      return p.indexOf('color') < 0;
    });
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if the specified outline icon doesn't exist`, () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      const p = path.toString();
      if (p.indexOf('.zip') > -1) {
        return false;
      }
      return p.indexOf('outline') < 0;
    });
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if a file with the app name already exists and force flag specified`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png',
        force: true
      }
    });
    assert.strictEqual(actual, true);
  });

  it(`passes validation if all arguments are correct`, () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.zip') < 0);
    const actual = command.validate({
      options: {
        portalUrl: 'https://contoso.sharepoint.com',
        appName: 'Contoso',
        description: 'Contoso',
        longDescription: `Stay on top of what's happening at Contoso`,
        companyName: 'Contoso',
        companyWebsiteUrl: 'https://contoso.com',
        coloredIconPath: 'icon-color.png',
        outlineIconPath: 'icon-outline.png'
      }
    });
    assert.strictEqual(actual, true);
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
