import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import * as request from 'request-promise-native';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { ContextInfo } from '../../spo';
import SpoCommand from '../../SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  identity: string;
}

class AppRetractCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_RETRACT;
  }

  public get description(): string {
    return 'Retracts/disables an app from the tenant app catalog';
  }

  
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.AppId = args.options.identity;
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const verbose: boolean = args.options.verbose || false;
    let appCatalogUrl: string = '';

    auth
      .ensureAccessToken(auth.service.resource, this, verbose)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (verbose) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/contextinfo`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Retrieving appcatalog url...`);

        const requestOptions: any = {
          url: `${auth.site.url}/_api/search/query?querytext='contentclass:STS_Site%20AND%20SiteTemplate:APPCATALOG'`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json',
            'X-RequestDigest': res.FormDigestValue
          },
          body: ``
        };

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: string): Promise<string> => {
        
        const json: any = JSON.parse(res);

        const cells: any[] = json.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells;

        const cell: any = cells.filter(t => t.Key === "SPWebUrl")[0];

        appCatalogUrl = cell.Value;

        if (verbose) {
          cmd.log('Retrieving request digest for app catalog');
        }

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/contextinfo`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): Promise<string> => {

        const formDigestValue = res;

        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Retracting app...`);

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/tenantappcatalog/AvailableApps/GetByID('${args.options.identity}')/retract`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=verbose',
            'X-RequestDigest': formDigestValue
          },
          body: ``
        };

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {

        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log("App retracted");

        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-id, --identity [guid]',
        description: 'The id of the app to deploy. Needs to be available in the app catalog.'
      }
    ];

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.APP_RETRACT).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
   
  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.APP_RETRACT} --identity 058140e3-0e37-44fc-a1d3-79c487d371a3
      Disables the given app in the tenant app catalog.

`);
    };
  }
}

module.exports = new AppRetractCommand();