import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { StandardWebPart, StandardWebPartUtils } from '../../common/StandardWebPartTypes';
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
  standardWebPart?: StandardWebPart;
  webPartId?: string;
  webPartProperties?: string;
  section?: number;
  column?: number;
  order?: number;
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
		let commandContext: ICommandContext = {
			requestContext: {
				accessToken: '',
				requestDigest: ''
			},
			debug: this.debug,
			verbose: this.verbose,
			webUrl: args.options.webUrl,
			pageName: args.options.pageName,
			log: (msg:any) => cmd.log(msg)
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

				if (this.verbose) {
					cmd.log(
						`Retrieving Client Side definition for WebPart...`
					);
				}

				if (args.options.standardWebPart) {
					// Get the standard WebPart according to arguments
					return ClientSidePageCommandHelper.getStandardWebPartInstance(
						commandContext,
						args.options.standardWebPart
					);
				} else {
					// Get the WebPart according to arguments
					return ClientSidePageCommandHelper.getWebPartInstance(commandContext, args.options.webPartId as string);
				}
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
				// Add the WebPart to page, set layout and properties
				ClientSidePageCommandHelper.addWebPartToPage(
					clientSidePage as ClientSidePage,
					webPart,
					commandContext,
					args.options.section,
					args.options.column,
					args.options.order,
					args.options.webPartProperties
				);
			})
			.then(() => {
				if (this.verbose) {
					cmd.log(`Saving Client Side Page...`);
				}
				// Save the Client Side Page with updated information
				return ClientSidePageCommandHelper.saveClientSidePage(commandContext, clientSidePage as ClientSidePage);
			})
			.then(() => {
				if (this.verbose) {
					cmd.log(
						`Client Side WebPart ${args.options.webPartId ||
							args.options.standardWebPart} is added to page ${commandContext.pageName}`
					);
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
				description: 'order of the WebPart in the column'
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

      if (!args.options.standardWebPart && !args.options.webPartId) {
        return 'Specify either the standardWebPart or the webPartId option';
      }

      if (args.options.standardWebPart && args.options.webPartId) {
        return 'Specify either the standardWebPart or the webPartId option but not both';
      }

      if (args.options.webPartId && !Utils.isValidGuid(args.options.webPartId)) {
        return `The webPartId '${args.options.webPartId}' is not a valid GUID`;
      }

      if (args.options.standardWebPart && !StandardWebPartUtils.isValidStandardWebPartType(args.options.standardWebPart)) {
        return `${args.options.standardWebPart} is not a valid standard WebPart type`;
      }

      if (args.options.webPartProperties) {
        try {
          JSON.parse(args.options.webPartProperties);
        }
        catch (e) {
          return `Specified webPartProperties is not a valid JSON string. Input: ${args.options
            .webPartProperties}. Error: ${e}`;
        }
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

    If the specified ${chalk.grey('pageName')} doesn't refer to an existing modern page,
    you will get a ${chalk.grey("File doesn't exists")} error.

    To add a standard web part to the page, specify one of the following values:
    ${chalk.grey("ContentRollup, BingMap, ContentEmbed, DocumentEmbed, Image,")}
    ${chalk.grey("ImageGallery, LinkPreview, NewsFeed, NewsReel, PowerBIReportEmbed,")}
    ${chalk.grey("QuickChart, SiteActivity, VideoEmbed, YammerEmbed, Events,")}
    ${chalk.grey("GroupCalendar, Hero, List, PageTitle, People, QuickLinks,")}
    ${chalk.grey("CustomMessageRegion, Divider, MicrosoftForms, Spacer")}.

    When specifying the JSON string with web part properties on Windows, you
    have to escape double quotes in a specific way. Considering the following
    value for the _webPartProperties_ option: {"Foo":"Bar"},
    you should specify the value as \`"{""Foo"":""Bar""}"\`. In addition,
    when using PowerShell, you should use the --% argument.

  Examples:

    Add the standard Bing Map web part to a modern page in the first available location on the page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap

    Add the standard Bing Map web part to a modern page in the third column of the second section
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --section 2 --column 3

    Add a custom web part to the modern page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1

    Using PowerShell, add the standard Bing Map web part with the specific properties to a modern page
      ${chalk.grey(config.delimiter)} --% ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties \`"{""Title"":""Foo location""}"\`

    Using Windows command line, add the standard Bing Map web part with the specific properties to a modern page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties \`"{""Title"":""Foo location""}"\`
      `
    );
  }
}

module.exports = new SpoPageClientSideWebPartAddCommand();
