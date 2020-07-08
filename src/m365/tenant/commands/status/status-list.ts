import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class TenantStatusListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_STATUS_LIST;
  }

  public get description(): string {
    return 'Gets health status of the different services in Microsoft 365';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Getting the health status of the different services in Microsoft 365.`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = 'ServiceComms/CurrentStatus';

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

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Gets health status of the different services in Microsoft 365
      ${commands.TENANT_STATUS_LIST}

  More information:
    
    Microsoft 365 Service Communications API reference:
      https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-current-status
  ` );
  }
}

module.exports = new TenantStatusListCommand();