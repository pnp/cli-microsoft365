import commands from '../../commands';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class TenantServiceListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.TENANT_SERVICE_LIST}`;
  }

  public get description(): string {
    return 'Gets the services available in Microsoft 365 from the Microsoft 365 Management API';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Getting the health status of the different services in Microsoft 365.`);
    }

    const serviceUrl: string = 'https://manage.office.com/api/v1.0';
    const statusEndpoint: string = 'ServiceComms/Services';

    this
      .getTenantId(cmd, this.debug)
      .then((tenantId: string): Promise<string> => {
        const pos: number = tenantId.indexOf(':') + 1;
        const tenantIdentifier = tenantId.substr(pos, tenantId.indexOf('&') - pos);

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
      ${commands.TENANT_SERVICE_LIST}
`);
  }
}

module.exports = new TenantServiceListCommand();