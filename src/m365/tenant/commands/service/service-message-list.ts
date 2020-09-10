import auth from '../../../../Auth';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import Command, { CommandOption } from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return 'Gets the service messages regarding services in Office 365 from the Office 365 Management API';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Getting the service messages regarding services in Office 365 from the Office 365 Management API`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = (typeof args.options.workload != 'undefined' && args.options.workload) ? `ServiceComms/Messages?$filter=Workload eq '${escape(args.options.workload)}'` : 'ServiceComms/Messages';
    
    const requestOptions: any = {
      url: `${serviceUrl}/${auth.service.tenantId}/${statusEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          cmd.log(res.value.map((r: any) => {
            return {
              Workload: r.Workload,
              Id: r.Id,
              ImpactDescription: r.ImpactDescription
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
        option: '-w, --workload [workload]	',
        description: 'Allows retrieval of the service messages for only one particular service. If not provided, the service messages of all services will be returned.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    log(vorpal.find(commands.TENANT_SERVICE_MESSAGE_LIST).helpInformation());
    log(
      `  Examples:

        Get service messages of all services in Microsoft 365
          ${commands.TENANT_SERVICE_MESSAGE_LIST}

        Get service messages of only one particular service in Microsoft 365
          ${commands.TENANT_SERVICE_MESSAGE_LIST} -w SharePoint

        More information:

          Microsoft 365 Service Communications API reference
            https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages
    `);
  }
}

module.exports = new TenantServiceMessageListCommand();