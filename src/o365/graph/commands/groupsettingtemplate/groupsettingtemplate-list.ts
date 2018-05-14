import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import { GroupSettingTemplate } from './GroupSettingTemplate';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class GraphGroupSettingTemplateListCommand extends GraphItemsListCommand<GroupSettingTemplate> {
  public get name(): string {
    return `${commands.GROUPSETTINGTEMPLATE_LIST}`;
  }

  public get description(): string {
    return 'Lists Azure AD group settings templates';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${auth.service.resource}/v1.0/groupSettingTemplates`, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(i => {
            return {
              id: i.id,
              displayName: i.displayName
            };
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To list group setting templates, you have to first connect to
    the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

  Examples:
  
    List all group setting templates in the tenant
      ${chalk.grey(config.delimiter)} ${this.name}
`);
  }
}

module.exports = new GraphGroupSettingTemplateListCommand();