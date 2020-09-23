import auth from '../../../../Auth';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const vorpal: Vorpal = require('../../../../vorpal-init');

class TenantServiceListCommand extends Command {
  public get name(): string {
    return `${commands.TENANT_SERVICE_LIST}`;
  }

  public get description(): string {
    return 'Gets services available in Microsoft 365';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Getting the health status of the different services in Microsoft 365.`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = 'ServiceComms/Services';

    const tenantId = Utils.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);

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
          cmd.log(res);
        }
        else {
          cmd.log(res.value.map((r: any) => {
            return {
              Id: r.Id,
              DisplayName: r.DisplayName
            }
          }));
        }
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    log(vorpal.find(commands.TENANT_SERVICE_LIST).helpInformation());
    log(
      `  Examples:

    Get services available in Microsoft 365
      m365 ${commands.TENANT_SERVICE_LIST}
      
  More information:

    Microsoft 365 Service Communications API reference
      https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages
`);
  }
}

module.exports = new TenantServiceListCommand();