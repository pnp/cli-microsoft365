import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListInstance } from './ListInstance';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
}

class SpoListSiteScriptGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SITESCRIPT_GET;
  }

  public get description(): string {
    return 'Extracts a site script from a SharePoint list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
      cmd.log(`Extracting Site Script from list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      if (this.debug) {
        cmd.log(`Retrieving List Url from Id '${args.options.listId}'...`);
      }
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')?$expand=RootFolder`;
    }
    else {
      if (this.debug) {
        cmd.log(`Retrieving List Url from Title '${args.options.listTitle}'...`);
      }
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')?$expand=RootFolder`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): Promise<any> => {
        const listAbsoluteUrl = Utils.getAbsoluteUrl(args.options.webUrl, listInstance.RootFolder.ServerRelativeUrl);
        const requestUrl = `${args.options.webUrl}/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`;
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          },
          json: true,
          body: {
            listUrl: listAbsoluteUrl
          }
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        const siteScript: string | null = res.value;
        if (!siteScript) {
          cb(new CommandError(`An error has occurred, the site script could not be extracted from list '${args.options.listId || args.options.listTitle}'`));
          return;
        }

        cmd.log(siteScript);
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to extract the site script from is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list to extract the site script from. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list to extract the site script from. Specify either listId or listTitle but not both'
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

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Extract a site script from an existing SharePoint list with title
    ${chalk.grey('ContosoList')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_SITESCRIPT_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle ContosoList

    Extract a site script from an existing SharePoint list with id
    ${chalk.grey('cc27a922-8224-4296-90a5-ebbc54da2e85')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_SITESCRIPT_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --listId cc27a922-8224-4296-90a5-ebbc54da2e85

      `);
  }
}

module.exports = new SpoListSiteScriptGetCommand();