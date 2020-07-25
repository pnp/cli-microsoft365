import * as chalk from 'chalk';
import auth from '../../../../Auth';
import { CommandInstance } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

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
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new TenantServiceListCommand();