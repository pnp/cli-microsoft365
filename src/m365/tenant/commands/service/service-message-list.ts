import * as chalk from 'chalk';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  workload: string;
}

class TenantServiceMessageListCommand extends Command {
  public get name(): string {
    return `${commands.TENANT_SERVICE_MESSAGE_LIST}`;
  }

  public get description(): string {
    return 'Gets service messages Microsoft 365';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.log(`Getting service messages...`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = args.options.workload ? `ServiceComms/Messages?$filter=Workload eq '${encodeURIComponent(args.options.workload)}'` : 'ServiceComms/Messages';
    const tenantId: string = Utils.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);

    const requestOptions: any = {
      url: `${serviceUrl}/${tenantId}/${statusEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          logger.log(res);
        }
        else {
          logger.log(res.value.map((r: any) => {
            return {
              Workload: r.Id.startsWith('MC') ? r.AffectedWorkloadDisplayNames.join(', ') : r.Workload,
              Id: r.Id,
              Message: r.Id.startsWith('MC') ? r.Title : r.ImpactDescription
            }
          }));
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --workload [workload]	',
        description: 'Retrieve service messages for the particular workload. If not provided, retrieves messages for all workloads'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantServiceMessageListCommand();