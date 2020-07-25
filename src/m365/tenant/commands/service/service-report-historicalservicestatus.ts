import auth from '../../../../Auth';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import Command, { CommandOption } from '../../../../Command';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';
import * as chalk from 'chalk';

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
          cmd.log(chalk.green('DONE'));
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
}

module.exports = new TenantServiceReportHistoricalServiceStatusCommand();