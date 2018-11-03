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
  sourceId:string;
  trimDuplicates?:boolean;
  enableStemming?:boolean;
  culture?:number;
  refinementFilters:string;
  queryTemplate:string;
  sortList:string;
  rankingModelId:string;
  startRow?:number;
  properties:string;
  sourceName:string;
  refiners:string;
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
    telemetryProps.sourceId = args.options.sourceId;
    telemetryProps.trimDuplicates = args.options.trimDuplicates;
    telemetryProps.enableStemming = args.options.enableStemming;
    telemetryProps.culture = args.options.culture;
    telemetryProps.refinementFilters = args.options.refinementFilters;
    telemetryProps.queryTemplate = args.options.queryTemplate;
    telemetryProps.sortList = args.options.sortList;
    telemetryProps.rankingModelId = args.options.rankingModelId;
    telemetryProps.startRow = args.options.startRow;
    telemetryProps.properties = args.options.properties;
    telemetryProps.sourceName = args.options.sourceName;
    telemetryProps.refiners = args.options.refiners;
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
        const startRow = args.options.startRow ? args.options.startRow : 0;

        return this.executeSearchQuery(cmd,args,accessToken,webUrl,[],startRow);
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
    const sourceIdRequestString = args.options.sourceId ? `&sourceid='${args.options.sourceId}'` : ``;
    const trimDuplicatesRequestString = `&trimduplicates=${args.options.trimDuplicates ? args.options.trimDuplicates : "false"}`;
    const enableStemmingRequestString = `&enablestemming=${typeof(args.options.enableStemming) === 'undefined' ? "true" : args.options.enableStemming}`;
    const cultureRequestString = args.options.culture ? `&culture=${args.options.culture}` : ``;
    const refinementFiltersRequestString = args.options.refinementFilters ? `&refinementfilters='${args.options.refinementFilters}'` : ``;
    const queryTemplateRequestString = args.options.queryTemplate ? `&querytemplate='${args.options.queryTemplate}'` : ``;
    const sortListRequestString = args.options.sortList ? `&sortList='${encodeURIComponent(args.options.sortList)}'` : ``;
    const rankingModelIdRequestString = args.options.rankingModelId ? `&rankingmodelid='${args.options.rankingModelId}'` : ``;
    const propertiesRequestString = this.getPropertiesRequestString(args);
    const refinersRequestString = args.options.refiners ? `&refiners='${args.options.refiners}'` : ``;

    //Construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.query}'`.concat(
      propertySelectRequestString,
      startRowRequestString,
      rowLimitRequestString,
      sourceIdRequestString,
      trimDuplicatesRequestString,
      enableStemmingRequestString,
      cultureRequestString,
      refinementFiltersRequestString,
      queryTemplateRequestString,
      sortListRequestString,
      rankingModelIdRequestString,
      propertiesRequestString,
      refinersRequestString
    );

    if(this.debug) {
      cmd.log(`RequestURL: ${requestUrl}`);
    }
    return requestUrl;
  }

  private getPropertiesRequestString(args: CommandArgs):string {
    let properties = args.options.properties ? args.options.properties : '';
    if(args.options.sourceName) {
      if(properties && !properties.endsWith(",")) { properties += `,`; }
      properties += `SourceName:${args.options.sourceName},SourceLevel:SPSite`;
    }
    return properties ? `&properties='${properties}'` : ``;
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
      },
      {
        option: '--sourceId <sourceId>',
        description: 'Specifies the identifier (GUID) of the result source to be used to run the query.'
      },
      {
        option: '--trimDuplicates',
        description: 'Specifies whether near duplicate items should be removed from the search results.'
      },
      {
        option: '--enableStemming',
        description: 'Specifies whether stemming is enabled.'
      },
      {
        option: '--culture <culture>',
        description: 'The locale for the query.'
      },
      {
        option: '--refinementFilters <refinementFilters>',
        description: 'The set of refinement filters used when issuing a refinement query. For GET requests, the RefinementFilters parameter is specified as an FQL filter. For POST requests, the RefinementFilters parameter is specified as an array of FQL filters.'
      },
      {
        option: '--queryTemplate <queryTemplate>',
        description: 'A string that contains the text that replaces the query text, as part of a query transformation.'
      },
      {
        option:'--sortList <sortList>',
        description: 'The list of properties by which the search results are ordered.'
      },
      {
        option:'--rankingModelId <rankingModelId>',
        description: 'The ID of the ranking model to use for the query.'
      },
      {
        option: '--startRow <startRow>',
        description: 'The first row that is included in the search results that are returned. You use this parameter when you want to implement paging for search results.'
      },
      {
        option: '--properties <properties>',
        description: 'Additional properties for the query. GET requests support only string values. POST requests support values of any type.'
      },
      {
        option: '--sourceName <sourceName>',
        description: 'Specified the name of the result source to be used to run the query.'
      },
      {
        option: '--refiners <refiners>',
        description: 'The set of refiners to return in a search result.'
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
      if (args.options.sourceId && !Utils.isValidGuid(args.options.sourceId)) {
        return `${args.options.sourceId} is not a valid GUID`;
      }
      if (args.options.rankingModelId && !Utils.isValidGuid(args.options.rankingModelId)) {
        return `${args.options.rankingModelId} is not a valid GUID`;
      }
      if(args.options.sortList && !Utils.isRegExMatch(args.options.sortList,"^([a-z0-9_]+:(ascending|descending))(,([a-z0-9_]+:(ascending|descending)))*$")) {
        return `sortlist parameter value '${args.options.sortList}' does not match the required pattern (=comma separated list of '<property>:(ascending|descending)'-pattern)`;
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

    if(!args.options.output || args.options.output == 'text') {
      cmd.log("# Rows: "+results[results.length-1].PrimaryQueryResult.RelevantResults.TotalRows);
      cmd.log("# Rows (Including duplicates): "+results[results.length-1].PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates);
      cmd.log("Elapsed Time: "+this.getElapsedTime(results));
    }
  }

  private getElapsedTime(results:SearchResult[]) {
    var totalTime:number = 0;
    results.forEach(result => {
      totalTime += result.ElapsedTime;
    });
    return totalTime;
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
  
    Execute search query to retrieve all Document Sets (ContentTypeId = '${chalk.grey('0x0120D520')}') for the english locale
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'ContentTypeId:0x0120D520' --culture 1033

    Retrieve all documents. For each document, retrieve the Path, Author and FileType.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'IsDocument:1' --selectProperties 'Path,Author,FileType' --allResults
    
    Return the top 50 items of which the title starts with 'Marketing' while trimming duplicates.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'Title:Marketing*' --rowLimit=50 --trimDuplicates

    Return only items from a specific resultsource (using the source id).
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query '*' --sourceId 6e71030e-5e16-4406-9bff-9c1829843083
      `);
  }
}

module.exports = new SearchCommand();