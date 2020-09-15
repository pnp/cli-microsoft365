import auth from '../../../../Auth';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import Command, { CommandOption } from '../../../../Command';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  workload?: string;
}

class TenantServiceReportHistoricalServiceStatusCommand extends Command {
  public get name(): string {
    return `${commands.TENANT_SERVICE_REPORT_HISTORICALSERVICESTATUS}`;
  }

  public get description(): string {
    return 'Gets the historical service status of Microsoft 365 Services of the last 7 days';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.workload = args.options.workload;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Gets the historical service status of Microsoft 365 Services of the last 7 days`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = typeof args.options.workload !== 'undefined' ? `ServiceComms/HistoricalStatus?$filter=Workload eq '${encodeURIComponent(args.options.workload)}'` : 'ServiceComms/HistoricalStatus';
    const tenantId: string = Utils.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);

    const requestOptions: any = {
      url: `${serviceUrl}/${tenantId}/${statusEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request.get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          cmd.log(res.value.map((r: any) => {
            return {
              WorkloadDisplayName: r.WorkloadDisplayName,
              StatusDisplayName: r.StatusDisplayName,
              StatusTime: r.StatusTime
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
        description: 'Retrieve the historical service status for the particular service. If not provided, the historical service status of all services will be returned.'
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
    option, execute ${chalk.grey(`m365 ${this.name} --output json`)}
    and get the value of the ${chalk.grey('Workload')} property for the particular service.
      
  Examples:
  
    Gets the historical service status of Microsoft 365 Services for the last
    7 days
      m365 ${this.name}

    Gets the historical service status of Microsoft Teams for the last 7 days
      m365 ${this.name} --workload "microsoftteams"

  More information:

    Microsoft 365 Service Communications API reference
      https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-historical-status
  ` );
  }
}

module.exports = new TenantServiceReportHistoricalServiceStatusCommand();