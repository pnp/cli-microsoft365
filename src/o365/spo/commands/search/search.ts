import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { Auth } from '../../../../Auth';
import Utils from '../../../../Utils';
import { SearchResult } from './datatypes/SearchResult';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  query: string;
  selectProperties:string;
}

class SearchCommand extends SpoCommand {
  public get name(): string {
    return commands.SEARCH;
  }

  public get description(): string {
    return 'Executes a search query';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.query = (!(!args.options.query)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl = auth.site.url;
    const resource: string = Auth.getResourceFromUrl(webUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Executing search query '${args.options.query}' on site at ${webUrl}...`);
        }

        const requestOptions: any = {
          url: this.getRequestUrl(webUrl,cmd,args),
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((searchResult: SearchResult): void => {
        if (this.debug) {
          cmd.log(`${searchResult.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates} Results found (including duplicates) :`);
          cmd.log('');
          cmd.log(searchResult);
          cmd.log('');
        }

        if (args.options.output === 'json') {
          cmd.log(searchResult);
        }
        else {
          cmd.log(this.getTextOutput(args,searchResult));
        }


        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getTextOutput(args: CommandArgs,searchResult: SearchResult) {
    var selectProperties = this.getSelectPropertiesArray(args);
    var outputData = searchResult.PrimaryQueryResult.RelevantResults.Table.Rows.map(row => {
      var rowOutput:any = {};

      row.Cells.map(cell => { 
        if(selectProperties.filter(prop => { return cell.Key.toUpperCase() === prop.toUpperCase() }).length > 0) {
          rowOutput[cell.Key] = cell.Value;
        }
      })

      return rowOutput;
    });
    return outputData;
  }

  private getRequestUrl(webUrl:string, cmd:CommandInstance, args: CommandArgs): string {
    //Get arg data
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    //Transform arg data to requeststrings
    const propertySelect: string = selectPropertiesArray.length > 0 ?
        `&selectproperties='${encodeURIComponent(selectPropertiesArray.join(","))}'` :
        ``;

    //Construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.query}'${propertySelect}`

    if(this.debug) {
      cmd.log(`RequestURL: ${requestUrl}`);
    }
    return requestUrl;
  }

  private getSelectPropertiesArray(args: CommandArgs) {
    return args.options.selectProperties 
      ? args.options.selectProperties.split(",")
      : (!args.options.output || args.options.output === "text") 
        ? ["Rank","DocId","OriginalPath"/*,"PartitionId","UrlZone","Culture","ResultTypeId","RenderTemplateId"*/]
        : [];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-q, --query <query>',
        description: 'Query to be executed in KQL format'
      },
      {
        option: '-p, --selectProperties <selectProperties>',
        description: 'Comma separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.query) {
        return 'Required parameter query missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
      using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To execute a search query you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Execute search query to retrieve all Document Sets (ContentTypeId = '${chalk.grey('0x0120D520')}')
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'ContentTypeId:0x0120D520'

    Retrieve all documents. For each document, retrieve the Path, Author and FileType.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'IsDocument:1' --selectProperties 'Path,Author,FileType'
      `);
  }
}

module.exports = new SearchCommand();