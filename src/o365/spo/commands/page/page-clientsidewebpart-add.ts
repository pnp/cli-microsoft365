import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import {
	ClientSidePage,
	CanvasSection,
	ClientSideWebpart,
	ClientSidePageComponent,
	CanvasColumn
} from './clientsidepages';
import { StandardWebPart, StandartWebPartUtils } from '../../common/StandardWebPartTypes';
import { ContextInfo } from '../../spo';
import { isNumber } from 'util';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
	options: Options;
}

interface Options extends GlobalOptions {
	pageName: string;
	webUrl: string;
	standardWebPart: StandardWebPart;
	webPartId: string;
	webPartProperties: string;
	section: number;
	column: number;
	order: number;
}

class SpoPageClientSideWebPartAddCommand extends SpoCommand {
	public get name(): string {
		return `${commands.PAGE_CLIENTSIDEWEBPART_ADD}`;
	}

	public get description(): string {
		return 'Adds a client side WebPart to a modern page';
	}

	public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
		const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
		let siteAccessToken: string = '';
		let requestDigest: string = '';
		let clientSidePage: ClientSidePage = null as any;

		let pageName: string = args.options.pageName;
		if (args.options.pageName.indexOf('.aspx') < 0) {
			pageName += '.aspx';
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

				siteAccessToken = accessToken;

				if (this.verbose) {
					cmd.log(`Retrieving request digest...`);
				}

				return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
			})
			.then((res: ContextInfo): Promise<ClientSidePage> => {
				if (this.debug) {
					cmd.log('Response:');
					cmd.log(res);
					cmd.log('');
				}

				// Keep the reference of request digest for subsequent requests
				requestDigest = res.FormDigestValue;

				if (this.verbose) {
					cmd.log(`Retrieving Client Side Page ${pageName}...`);
				}
				// Get Client Side Page
				return this._getClientSidePage(pageName, cmd, args, siteAccessToken);
			})
			.then((csPage) => {
				// Keep the reference of client side page for subsequent requests
				clientSidePage = csPage;

				if (this.verbose) {
					cmd.log(
						`Retrieving Client Side definition for WebPart ${args.options.webPartId ||
							args.options.standardWebPart}...`
					);
				}
				// Get the WebPart according to arguments
				return this._getWebPart(cmd, args, siteAccessToken);
			})
			.then((webPart) => {
				if (this.debug) {
					cmd.log(`Retrieved WebPart definition:`);
					cmd.log(webPart);
					cmd.log('');
				}

				if (this.verbose) {
					cmd.log(`Setting Client Side WebPart layout and properties...`);
				}
				// Set the WebPart properties and layout (section, column and order)
				this._setWebPartPropertiesAndLayout(clientSidePage, webPart, cmd, args);
				if (this.verbose) {
					cmd.log(`Saving Client Side Page...`);
				}
				// Save the Client Side Page with updated information
				return this._saveClientSidePage(clientSidePage, cmd, args, pageName, siteAccessToken, requestDigest);
			})
			.then(() => {
				if (this.verbose) {
					cmd.log(
						`Client Side WebPart ${args.options.webPartId ||
							args.options.standardWebPart} is added to page ${pageName}`
					);
				}
				cb();
			})
			.catch((err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
	}

	private _getClientSidePage(
		pageName: string,
		cmd: CommandInstance,
		args: CommandArgs,
		accessToken: string
	): Promise<ClientSidePage> {
		return new Promise<ClientSidePage>((resolve, reject) => {
			if (this.verbose) {
				cmd.log(`Retrieving information about the page...`);
			}

			const webUrl = args.options.webUrl;
			const requestOptions: any = {
				url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${webUrl.substr(
					webUrl.indexOf('/', 8)
				)}/sitepages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
				headers: Utils.getRequestHeaders({
					authorization: `Bearer ${accessToken}`,
					accept: 'application/json;odata=nometadata'
				}),
				json: true
			};

			if (this.debug) {
				cmd.log('Executing web request...');
				cmd.log(requestOptions);
				cmd.log('');
			}

			request
				.get(requestOptions)
				.then((res) => {
					if (this.debug) {
						cmd.log('Response:');
						cmd.log(res);
						cmd.log('');
					}

					if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
						reject(new Error(`Page ${args.options.pageName} is not a modern page.`));
						return;
					}

					const clientSidePage: ClientSidePage = ClientSidePage.fromHtml(
						res.ListItemAllFields.CanvasContent1
					);

					resolve(clientSidePage);
				})
				.catch((error) => {
					cmd.log(error);
					reject(error);
				});
		});
	}

	private _getWebPart(cmd: CommandInstance, args: CommandArgs, accessToken: string): Promise<ClientSideWebpart> {
		return new Promise<ClientSideWebpart>((resolve, reject) => {
			const standardWebPart = args.options.standardWebPart;

			const webPartId = standardWebPart
				? StandartWebPartUtils.getWebPartId(standardWebPart)
				: args.options.webPartId;

			if (this.debug) {
				cmd.log(`StandardWebPart: ${standardWebPart}`);
				cmd.log(`WebPartId: ${webPartId}`);
			}

			const webUrl = args.options.webUrl;
			const requestOptions: any = {
				url: `${webUrl}/_api/web/getclientsidewebparts()`,
				headers: Utils.getRequestHeaders({
					authorization: `Bearer ${accessToken}`,
					accept: 'application/json;odata=nometadata'
				}),
				json: true
			};

			if (this.debug) {
				cmd.log('Executing web request...');
				cmd.log(requestOptions);
				cmd.log('');
			}

			request
				.get(requestOptions)
				.then((res: { value: ClientSidePageComponent[] }) => {
					if (this.debug) {
						cmd.log('Response:');
						cmd.log(res);
						cmd.log('');
					}

					const webPartDefinition = res.value.filter((c) => c.Id === webPartId);
					if (webPartDefinition.length === 0) {
						reject(new Error(`There is no available WebPart with Id ${webPartId}.`));
						return;
					}

					if (this.debug) {
						cmd.log('WebPart definition:');
						cmd.log(webPartDefinition);
						cmd.log('');
					}

					if (this.verbose) {
						cmd.log(`Creating instance from definition of WebPart ${webPartId}...`);
					}
					const webPart = ClientSideWebpart.fromComponentDef(webPartDefinition[0]);
					resolve(webPart);
				})
				.catch((error) => {
					reject(error);
				});
		});
	}

	private _setWebPartPropertiesAndLayout(
		clientSidePage: ClientSidePage,
		webPart: ClientSideWebpart,
		cmd: CommandInstance,
		args: CommandArgs
	) {
		let actualSectionIndex: number | null = args.options.section && args.options.section - 1;
		// If the section arg is not specified, set to first section
		if (actualSectionIndex == null) {
			if (this.debug) {
				cmd.log(`No section argument specified, The component will be added to default section`);
			}
			actualSectionIndex = 0;
		}
		// Make sure the section is in the range, other add a new section
		if (actualSectionIndex >= clientSidePage.sections.length) {
			throw new Error(`Invalid Section '${args.options.section}'`);
		}

		// Get the target section
		const section: CanvasSection = clientSidePage.sections[actualSectionIndex];

		let actualColumnIndex: number | null = args.options.column && args.options.column - 1;
		// If the column arg is not specified, set to first column
		if (actualColumnIndex == null) {
			if (this.debug) {
				cmd.log(`No column argument specified, The component will be added to default column`);
			}
			actualColumnIndex = 0;
		}
		// Make sure the column is in the range of the current section
		if (actualColumnIndex >= section.columns.length) {
			throw new Error(`Invalid Column '${args.options.column}'`);
		}

		// Get the target column
		const column: CanvasColumn = section.columns[actualColumnIndex];

		if (args.options.webPartProperties) {
			if (this.debug) {
				cmd.log('WebPart properties: ');
				cmd.log(args.options.webPartProperties);
				cmd.log('');
			}

			try {
				let properties = JSON.parse(args.options.webPartProperties);
				webPart.setProperties(properties);
			} catch (e) {
			}
		}

		// Add the WebPart to to appropriate location
		column.addControl(webPart);
	}

	private _saveClientSidePage(
		clientSidePage: ClientSidePage,
		cmd: CommandInstance,
		args: CommandArgs,
		pageName: string,
		accessToken: string,
		requestDigest: string
	): request.RequestPromise {
		const serverRelativeSiteUrl: string = `${args.options.webUrl.substr(
			args.options.webUrl.indexOf('/', 8)
		)}/sitepages/${pageName}`;

		const updatedContent = clientSidePage.toHtml();

		if (this.debug) {
			cmd.log('Updated canvas content: ');
			cmd.log(updatedContent);
			cmd.log('');
		}

		const requestOptions: any = {
			url: `${args.options
				.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields`,
			headers: Utils.getRequestHeaders({
				authorization: `Bearer ${accessToken}`,
				'X-RequestDigest': requestDigest,
				'content-type': 'application/json;odata=nometadata',
				'X-HTTP-Method': 'MERGE',
				'IF-MATCH': '*',
				accept: 'application/json;odata=nometadata'
			}),
			body: {
				CanvasContent1: updatedContent
			},
			json: true
		};

		if (this.debug) {
			cmd.log('Executing web request...');
			cmd.log(requestOptions);
			cmd.log('');
		}

		return request.post(requestOptions);
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
				option: '--standardWebPart [standardWebPart]',
				description: `set to add one of the standard SharePoint web parts.
                      Available values: ContentRollup | BingMap | ContentEmbed | DocumentEmbed | Image | ImageGallery | LinkPreview | NewsFeed | NewsReel 
                      | PowerBIReportEmbed | QuickChart | SiteActivity | VideoEmbed | YammerEmbed | Events | GroupCalendar | Hero | List 
                      | PageTitle | People | QuickLinks | CustomMessageRegion | Divider | MicrosoftForms | Spacer`
			},
			{
				option: '--webPartId [webPartId]',
				description: 'Set to add a custom web part'
			},
			{
				option: '--webPartProperties [webPartProperties]',
				description: 'JSON string with web part properties to set on the web part'
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
				description: 'order of the section'
			}
		];

		const parentOptions: CommandOption[] = super.options();
		return options.concat(parentOptions);
	}

	public validate(): CommandValidate {
		return (args: CommandArgs): boolean | string => {
			if (!args.options.pageName) {
				return 'Required parameter pageName is missing';
			}

			if (!args.options.standardWebPart && !args.options.webPartId) {
				return 'Either standardWebPart or webPartId parameter must be specified';
			}

			if (args.options.standardWebPart && args.options.webPartId) {
				return 'standardWebPart and webPartId parameters cannot be specified at the same time';
			}

			if (!args.options.standardWebPart) {
				if (!Utils.isValidGuid(args.options.webPartId)) {
					return `The WebPart Id ${args.options.webPartId} is not a valid GUID`;
				}
			} else {
				if (!StandartWebPartUtils.isValidStandardWebPartType(args.options.standardWebPart)) {
					return `${args.options.standardWebPart} is not a valid standard WebPart type`;
				}
			}

			if (args.options.webPartProperties) {
				try {
					JSON.parse(args.options.webPartProperties);
				} catch (e) {
					return `Specified webPartProperties is not a valid JSON string. Input: ${args.options
						.webPartProperties}. Error: ${e}`;
				}
			}

			if (args.options.section && (!isNumber(args.options.section) || args.options.section < 1)) {
				return 'The value of parameter section must be 1 or above';
			}

			if (args.options.column && (!isNumber(args.options.column) || args.options.column < 1)) {
				return 'The value of parameter column must be 1 or above';
			}

			if (!args.options.webUrl) {
				return 'Required parameter webUrl missing';
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

    To add a Client Side WebPart to a modern page, you have to first
    connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('pageName')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey("File doesn't exists")} error.

  Examples:
  
    Add a WebPart with ID
    ${chalk.grey('3ede60d3-dc2c-438b-b5bf-cc40bb2351e1')} onto a modern page
    with name ${chalk.grey('home.aspx')} on the third column of the second section
      ${chalk.grey(config.delimiter)} ${this
				.name} --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --section 2 --column 3

    Add the Bing Map WebPart onto a modern page
    with name ${chalk.grey('home.aspx')} on the default column of the first section
      ${chalk.grey(config.delimiter)} ${this
				.name} --standardWebPart BingMap --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx
      `
		);
	}
}

module.exports = new SpoPageClientSideWebPartAddCommand();
