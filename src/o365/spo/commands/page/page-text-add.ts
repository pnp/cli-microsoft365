import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { ContextInfo } from '../../spo';
import { isNumber } from 'util';
import { ClientSidePage } from './clientsidepages';
import ClientSidePageCommandHelper, { ICommandContext } from './ClientSidePageCommandHelper';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
	options: Options;
}

interface Options extends GlobalOptions {
	pageName: string;
	webUrl: string;
	text: string;
	section?: number;
	column?: number;
	order?: number;
}

class SpoPageClientSideWebPartAddCommand extends SpoCommand {
	public get name(): string {
		return `${commands.PAGE_TEXT_ADD}`;
	}

	public get description(): string {
		return 'Adds a client side text to a modern page';
	}

	public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
		const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
		let commandContext: ICommandContext = {
			requestContext: {
				accessToken: '',
				requestDigest: ''
			},
			debug: this.debug,
			verbose: this.verbose,
			webUrl: args.options.webUrl,
			pageName: args.options.pageName,
			log: (msg: any) => cmd.log(msg)
		};
		let clientSidePage: ClientSidePage | null = null;

		if (commandContext.pageName.indexOf('.aspx') < 0) {
			commandContext.pageName += '.aspx';
		}

		if (this.debug) {
			cmd.log(`Retrieving access token for ${resource}...`);
		}

		auth
			.getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
			.then((accessToken: string): request.RequestPromise => {
				if (this.debug) {
					cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
				}

				commandContext.requestContext.accessToken = accessToken;

				if (this.verbose) {
					cmd.log(`Retrieving request digest...`);
				}

				return this.getRequestDigestForSite(
					args.options.webUrl,
					commandContext.requestContext.accessToken,
					cmd,
					this.debug
				);
			})
			.then((res: ContextInfo): Promise<ClientSidePage> => {
				if (this.debug) {
					cmd.log('Response:');
					cmd.log(res);
					cmd.log('');
				}

				// Keep the reference of request digest for subsequent requests
				commandContext.requestContext.requestDigest = res.FormDigestValue;

				if (this.verbose) {
					cmd.log(`Retrieving Client Side Page ${commandContext.pageName}...`);
				}
				// Get Client Side Page
				return ClientSidePageCommandHelper.getClientSidePage(commandContext);
			})
			.then((csPage) => {
				if (this.debug) {
					cmd.log(`Retrieved Client Side Page:`);
					// cmd.log(csPage); // Cannot log because of circular structure on JSON stringifying
					cmd.log('');
				}

				// Keep the reference of client side page for subsequent requests
				clientSidePage = csPage;

				if (this.debug) {
					cmd.log(`Adding text to page ${args.options.pageName} on section ${args.options.section} column ${args.options.column} with order ${args.options.order}..`);
        }
        if (this.verbose) {
					cmd.log(`Adding text to page ${args.options.pageName}...`);
				}

				// Add the text to page and set layout
				ClientSidePageCommandHelper.addTextToPage(
          clientSidePage as ClientSidePage,
          args.options.text,
					commandContext,
					args.options.section,
					args.options.column,
					args.options.order
				);

				if (this.verbose) {
					cmd.log(`Saving Client Side Page...`);
				}
				// Save the Client Side Page with updated information
				return ClientSidePageCommandHelper.saveClientSidePage(commandContext, clientSidePage as ClientSidePage);
			})
			.then(() => {
				if (this.verbose) {
					cmd.log(`Text ${args.options.text} is added to page ${commandContext.pageName}`);
				}
				cb();
			})
			.catch((err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
	}

	public options(): CommandOption[] {
		const options: CommandOption[] = [
			{
				option: '-n, --pageName <pageName>',
				description: 'Name of the page where the control is located'
			},
			{
				option: '-u, --webUrl <webUrl>',
				description: 'URL of the site where the page to retrieve is located'
			},
			{
				option: '--text <text>',
				description: 'text to add to the page'
			},
			{
				option: '--section [section]',
				description: 'number of section to which the text should be added (1 or higher)'
			},
			{
				option: '--column [column]',
				description: 'number of column in which the text should be added (1 or higher)'
			},
			{
				option: '--order [order]',
				description: 'order of the text control in the column'
			}
		];

		const parentOptions: CommandOption[] = super.options();
		return options.concat(parentOptions);
	}

	public validate(): CommandValidate {
		return (args: CommandArgs): boolean | string => {
			if (!args.options.webUrl) {
				return 'Required parameter webUrl missing';
			}

			if (!args.options.pageName) {
				return 'Required option pageName is missing';
			}

			if (!args.options.text) {
				return 'Required option text is missing';
			}

			if (args.options.section && (!isNumber(args.options.section) || args.options.section < 1)) {
				return 'The value of parameter section must be 1 or higher';
			}

			if (args.options.column && (!isNumber(args.options.column) || args.options.column < 1)) {
				return 'The value of parameter column must be 1 or higher';
			}

			return SpoCommand.isValidSharePointUrl(args.options.webUrl);
		};
	}

	public commandHelp(args: {}, log: (help: string) => void): void {
		const chalk = vorpal.chalk;
		log(vorpal.find(this.name).helpInformation());
		log(
			`  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To add a client-side web part to a modern page, you have to first connect
    to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:

    Add the text 'Hello World!' to a modern page in the first available location on the page
      ${chalk.grey(config.delimiter)} ${this
				.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text "Hello World!"

    Add the text 'Hello World!' to a modern page in the third column of the second section
      ${chalk.grey(config.delimiter)} ${this
				.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text "Hello World!" --section 2 --column 3      `
		);
	}
}

module.exports = new SpoPageClientSideWebPartAddCommand();
