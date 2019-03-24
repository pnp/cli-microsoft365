import auth from '../../AzmgmtAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../AzmgmtCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class AzmgmtFlowEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<{ value: [{ name: string, properties: { displayName: string } }] }> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving list of Microsoft Flow environments...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: { value: [{ name: string, properties: { displayName: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(res.value);
          }
          else {
            cmd.log(res.value.map(e => {
              return {
                name: e.name,
                displayName: e.properties.displayName
              };
            }));
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_ENVIRONMENT_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Azure Management Service,
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To get information about Microsoft Flow environments, you have to first
    log in to the Azure Management Service using the ${chalk.blue(commands.LOGIN)} command.
   
  Examples:
  
    List Microsoft Flow environments in the current tenant
      ${chalk.grey(config.delimiter)} ${this.getCommandName()}
`);
  }
}

module.exports = new AzmgmtFlowEnvironmentListCommand();