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
import { ResultTableRow } from './datatypes/ResultTableRow';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  query: string;
  rowLimit:number;
  selectProperties:string;
  allResults?:boolean;
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
    telemetryProps.query = args.options.query;
    telemetryProps.selectproperties = args.options.selectProperties;
    telemetryProps.allResults = args.options.allResults;
    telemetryProps.rowLimit = args.options.rowLimit;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl = auth.site.url;
    const resource: string = Auth.getResourceFromUrl(webUrl);

    if (this.debug) {
      cmd.log("Calling Search API on = " + auth.site.url);
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): string => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Executing search query '${args.options.query}' on site at ${webUrl}...`);
        }

        return accessToken;
      })
      .then((accessToken:string):Promise<any[]> => {
        return this.executeSearchQuery(cmd,args,accessToken,webUrl,[],0);
      })
      .then((results:SearchResult[]) => { 
        this.printResults(cmd,args,results);
      })
      .then(() => {
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private executeSearchQuery(cmd:CommandInstance,args:CommandArgs,accessToken:string,webUrl:string,resultSet:SearchResult[],startRow:number):Promise<SearchResult[]> {
    return (():request.RequestPromise => { 
        const requestUrl:string = this.getRequestUrl(webUrl,cmd,args,startRow);
        const requestOptions: any = {
          url: requestUrl,
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
      })()
      .then((searchResult: SearchResult): SearchResult => {
        if (this.debug) {
          cmd.log(`${searchResult.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates} Results found (including duplicates) :`);
          cmd.log('');
          cmd.log(searchResult);
          cmd.log('');
        }

        resultSet.push(searchResult);

        return searchResult;
      })
      .then((searchResult:SearchResult):Promise<SearchResult[]> => {
        if(args.options.allResults) {
          if(startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount < searchResult.PrimaryQueryResult.RelevantResults.TotalRows) {
            const nextStartRow = startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount;
            return this.executeSearchQuery(cmd,args,accessToken,webUrl,resultSet,nextStartRow);
          }
        } 
        return new Promise<SearchResult[]>((resolve) => { resolve(resultSet); });
      })
      .then(() => { return resultSet });
  }

  private getRequestUrl(webUrl:string, cmd:CommandInstance, args: CommandArgs,startRow:number): string {
    //Get arg data
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    //Transform arg data to requeststrings
    const propertySelectRequestString: string = selectPropertiesArray.length > 0 ?
        `&selectproperties='${encodeURIComponent(selectPropertiesArray.join(","))}'` :
        ``;
    const startRowRequestString = `&startrow=${startRow ? startRow : 0}`;
    const rowLimitRequestString = args.options.rowLimit ? `&rowlimit=${args.options.rowLimit}` : ``;

    //Construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.query}'${propertySelectRequestString}${startRowRequestString}${rowLimitRequestString}`

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
      },
      {
        option: '--allResults',
        description: 'Get all results of the search query, not only the amount specified by the rowlimit (default: 10)'
      },
      {
        option: '--rowLimit <rowLimit>',
        description: 'Sets the number of rows to be returned. When the \'allResults\' parameter is enabled, it will determines the size of the batches being retrieved'
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

  private printResults(cmd:CommandInstance,args:CommandArgs,results:SearchResult[]) {
    if(results.length === 1) {
      if(args.options.output === 'json') {
        cmd.log(results[0]);
      } else {
        cmd.log(this.getTextOutput(args,results[0].PrimaryQueryResult.RelevantResults.Table.Rows));
      }
    } else {
      if(args.options.output === 'json') {
        cmd.log(results);
      } else {
        let allRows:ResultTableRow[] = [];
        for(let i = 0;i < results.length;i++) {
          allRows.push(...results[i].PrimaryQueryResult.RelevantResults.Table.Rows);
        }
        cmd.log(this.getTextOutput(args,allRows));
      }
    }
  }

  private getTextOutput(args: CommandArgs,searchResultRows: ResultTableRow[]) {
    var selectProperties = this.getSelectPropertiesArray(args);
    var outputData = searchResultRows.map(row => {
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
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'IsDocument:1' --selectProperties 'Path,Author,FileType' --allResults
    
    Return the top 50 items of which the title starts with 'Marketing'.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'Title:Marketing*' --rowLimit=50
      `);
  }
}

module.exports = new SearchCommand();