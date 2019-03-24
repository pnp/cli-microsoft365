import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./page-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.PAGE_REMOVE, () => {
	let vorpal: Vorpal;
	let log: string[];
	let cmdInstance: any;
	let cmdInstanceLogSpy: sinon.SinonSpy;
	let trackEvent: any;
	let telemetry: any;
	let promptOptions: any;

	const fakeRestCalls: (pageName?: string) => sinon.SinonStub = (pageName: string='page.aspx') => {
		return sinon.stub(request, 'post').callsFake((opts) => {
			if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/${pageName}')`) > -1) {
				return Promise.resolve();
			}

			return Promise.reject('Invalid request');
		});
	};

	before(() => {
		sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
		sinon.stub(auth, 'getAccessToken').callsFake(() => {
			return Promise.resolve('ABC');
		});
		sinon
			.stub(command as any, 'getRequestDigestForSite')
			.callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
		trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
			telemetry = t;
		});
	});

	beforeEach(() => {
		vorpal = require('../../../../vorpal-init');
		log = [];
		cmdInstance = {
			log: (msg: string) => {
				log.push(msg);
			},
			prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
				promptOptions = options;
				cb({ continue: false });
			}
		};
		cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
		auth.site = new Site();
		telemetry = null;
	});

	afterEach(() => {
		Utils.restore([ vorpal.find, request.post ]);
	});

	after(() => {
		Utils.restore([
			appInsights.trackEvent,
			auth.getAccessToken,
			auth.restoreAuth,
			(command as any).getRequestDigestForSite
		]);
	});

	it('has correct name', () => {
		assert.equal(command.name.startsWith(commands.PAGE_REMOVE), true);
	});

	it('has a description', () => {
		assert.notEqual(command.description, null);
	});

	it('calls telemetry', (done) => {
		cmdInstance.action = command.action();
		cmdInstance.action({ options: {} }, () => {
			try {
				assert(trackEvent.called);
				done();
			} catch (e) {
				done(e);
			}
		});
	});

	it('logs correct telemetry event', (done) => {
		cmdInstance.action = command.action();
		cmdInstance.action({ options: {} }, () => {
			try {
				assert.equal(telemetry.name, commands.PAGE_REMOVE);
				done();
			} catch (e) {
				done(e);
			}
		});
	});

	it('aborts when not logged in to a SharePoint site', (done) => {
		auth.site = new Site();
		auth.site.connected = false;
		cmdInstance.action = command.action();
		cmdInstance.action({ options: { debug: true } }, (err?: any) => {
			try {
				assert.equal(
					JSON.stringify(err),
					JSON.stringify(new CommandError('Log in to a SharePoint Online site first'))
				);
				done();
			} catch (e) {
				done(e);
			}
		});
	});

	it('removes a modern page without confirm prompt', (done) => {
		fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.action(
			{
				options: {
					debug: false,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a',
					confirm: true
				}
			},
			() => {
				try {
					assert(cmdInstanceLogSpy.notCalled);
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('removes a modern page (debug) without confirm prompt', (done) => {
    fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.action(
			{
				options: {
					debug: true,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a',
					confirm: true
				}
			},
			() => {
				try {
					assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('removes a modern page with confirm prompt', (done) => {
		fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
			promptOptions = options;
			cb({ continue: true });
		};
		cmdInstance.action(
			{
				options: {
					debug: false,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a'
				}
			},
			() => {
				try {
					assert(cmdInstanceLogSpy.notCalled);
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('removes a modern page (debug) with confirm prompt', (done) => {
		fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
			promptOptions = options;
			cb({ continue: true });
		};
		cmdInstance.action(
			{
				options: {
					debug: true,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a'
				}
			},
			() => {
				try {
					assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('should prompt before removing page when confirmation argument not passed', (done) => {
		fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.action(
			{
				options: {
					debug: true,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a'
				}
			},
			() => {
				let promptIssued = false;

				if (promptOptions && promptOptions.type === 'confirm') {
					promptIssued = true;
				}

				try {
					assert(promptIssued);
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('should abort page removal when prompt not confirmed', (done) => {
		let postCallSpy = fakeRestCalls();
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso-admin.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
			cb({ continue: false });
		};
		cmdInstance.action(
			{
				options: {
					debug: true,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a'
				}
			},
			() => {
				try {
					assert(postCallSpy.notCalled === true);
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('automatically appends the .aspx extension', (done) => {
		fakeRestCalls();
		auth.site = new Site();
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		auth.site.connected = true;
		cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
			cb({ continue: false });
		};
		cmdInstance.action(
			{
				options: {
					debug: false,
					name: 'page',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a',
					confirm: true
				}
			},
			() => {
				try {
					assert(cmdInstanceLogSpy.notCalled);
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('correctly handles OData error when removing modern page', (done) => {
		sinon.stub(request, 'post').callsFake((opts) => {
			return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
		});

		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
			cb({ continue: false });
		};
		cmdInstance.action(
			{
				options: {
					debug: false,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a',
					confirm: true
				}
			},
			(err?: any) => {
				try {
					assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});

	it('supports debug mode', () => {
		const options = command.options() as CommandOption[];
		let containsOption = false;
		options.forEach((o) => {
			if (o.option === '--debug') {
				containsOption = true;
			}
		});
		assert(containsOption);
	});

	it('supports specifying name', () => {
		const options = command.options() as CommandOption[];
		let containsOption = false;
		options.forEach((o) => {
			if (o.option.indexOf('--name') > -1) {
				containsOption = true;
			}
		});
		assert(containsOption);
	});

	it('supports specifying webUrl', () => {
		const options = command.options() as CommandOption[];
		let containsOption = false;
		options.forEach((o) => {
			if (o.option.indexOf('--webUrl') > -1) {
				containsOption = true;
			}
		});
		assert(containsOption);
	});

	it('supports specifying confirm', () => {
		const options = command.options() as CommandOption[];
		let containsOption = false;
		options.forEach((o) => {
			if (o.option.indexOf('--confirm') > -1) {
				containsOption = true;
			}
		});
		assert(containsOption);
	});

	it('fails validation if name not specified', () => {
		const actual = (command.validate() as CommandValidate)({
			options: { webUrl: 'https://contoso.sharepoint.com' }
		});
		assert.notEqual(actual, true);
	});

	it('fails validation if webUrl not specified', () => {
		const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx' } });
		assert.notEqual(actual, true);
	});

	it('fails validation if webUrl is not an absolute URL', () => {
		const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'foo' } });
		assert.notEqual(actual, true);
	});

	it('fails validation if webUrl is not a valid SharePoint URL', () => {
		const actual = (command.validate() as CommandValidate)({
			options: { name: 'page.aspx', webUrl: 'http://foo' }
		});
		assert.notEqual(actual, true);
	});

	it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
		const actual = (command.validate() as CommandValidate)({
			options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' }
		});
		assert.equal(actual, true);
	});

	it('passes validation when name has no extension', () => {
		const actual = (command.validate() as CommandValidate)({
			options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' }
		});
		assert.equal(actual, true);
	});

	it('has help referring to the right command', () => {
		const cmd: any = {
			log: (msg: string) => {},
			prompt: () => {},
			helpInformation: () => {}
		};
		const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
		cmd.help = command.help();
		cmd.help({}, () => {});
		assert(find.calledWith(commands.PAGE_REMOVE));
	});

	it('has help with examples', () => {
		const _log: string[] = [];
		const cmd: any = {
			log: (msg: string) => {
				_log.push(msg);
			},
			prompt: () => {},
			helpInformation: () => {}
		};
		sinon.stub(vorpal, 'find').callsFake(() => cmd);
		cmd.help = command.help();
		cmd.help({}, () => {});
		let containsExamples: boolean = false;
		_log.forEach((l) => {
			if (l && l.indexOf('Examples:') > -1) {
				containsExamples = true;
			}
		});
		Utils.restore(vorpal.find);
		assert(containsExamples);
	});

	it('correctly handles lack of valid access token', (done) => {
		Utils.restore(auth.getAccessToken);
		sinon.stub(auth, 'getAccessToken').callsFake(() => {
			return Promise.reject(new Error('Error getting access token'));
		});
		auth.site = new Site();
		auth.site.connected = true;
		auth.site.url = 'https://contoso.sharepoint.com';
		cmdInstance.action = command.action();
		cmdInstance.action(
			{
				options: {
					debug: false,
					name: 'page.aspx',
					webUrl: 'https://contoso.sharepoint.com/sites/team-a',
					confirm: true
				}
			},
			(err?: any) => {
				try {
					assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
					done();
				} catch (e) {
					done(e);
				}
			}
		);
	});
});
