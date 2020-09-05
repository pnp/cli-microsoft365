import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption } from '../../../../Command';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  workload?: string;
}

class TenantStatusListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_STATUS_LIST;
  }

  public get description(): string {
    return 'Gets health status of the different services in Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.sharingCapabilities = args.options.workload;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Getting the health status of the different services in Microsoft 365.`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = typeof args.options.workload !== 'undefined' ? `ServiceComms/CurrentStatus?$filter=Workload eq '${encodeURIComponent(args.options.workload)}'` : 'ServiceComms/CurrentStatus';

    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<string> => {
        const tenantIdentifier: string = _spoUrl.replace('https://', '');
        const requestOptions: any = {
          url: `${serviceUrl}/${tenantIdentifier}/${statusEndpoint}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          cmd.log(res.value.map((r: any) => {
            return {
              Name: r.WorkloadDisplayName,
              Status: r.StatusDisplayName
            }
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --workload [workload]',
        description: 'Retrieve service status for the specified service. If not provided, will list the current service status of all services'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    To get the name of the particular workload for use with the ${chalk.grey('workload')}
    option, execute ${chalk.grey(`${commands.TENANT_STATUS_LIST} --output json`)} and get
    the value of the ${chalk.grey('Workload')} property for the particular service.
      
  Examples:
  
    Gets health status of all services in Microsoft 365
      m365 ${this.name}

    Gets health status for SharePoint Online
      m365 ${this.name} --workload "SharePoint"

  More information:
    
    Microsoft 365 Service Communications API reference:
      https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-current-status
  ` );
  }
}

module.exports = new TenantStatusListCommand();