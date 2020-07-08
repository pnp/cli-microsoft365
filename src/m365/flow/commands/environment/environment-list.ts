import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class FlowEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving list of Microsoft Flow environments...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    request
      .get<{ value: [{ name: string, properties: { displayName: string } }] }>(requestOptions)
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
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
  Examples:
  
    List Microsoft Flow environments in the current tenant
      ${this.getCommandName()}
`);
  }
}

module.exports = new FlowEnvironmentListCommand();