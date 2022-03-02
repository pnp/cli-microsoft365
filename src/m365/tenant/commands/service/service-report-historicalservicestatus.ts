import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  workload?: string;
}

class TenantServiceReportHistoricalServiceStatusCommand extends Command {
  public get name(): string {
    return commands.SERVICE_REPORT_HISTORICALSERVICESTATUS;
  }

  public get description(): string {
    return 'Gets the historical service status of Microsoft 365 Services of the last 7 days';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.workload = args.options.workload;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['WorkloadDisplayName', 'StatusDisplayName', 'StatusTime'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Gets the historical service status of Microsoft 365 Services of the last 7 days`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = typeof args.options.workload !== 'undefined' ? `ServiceComms/HistoricalStatus?$filter=Workload eq '${encodeURIComponent(args.options.workload)}'` : 'ServiceComms/HistoricalStatus';
    const tenantId: string = accessToken.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken);

    const requestOptions: any = {
      url: `${serviceUrl}/${tenantId}/${statusEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request.get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --workload [workload]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantServiceReportHistoricalServiceStatusCommand();