import commands from '../../commands';
import aadcommands from '../../../aad/commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import request from '../../../../request';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description: string;
  name: string;
}

class TeamsTeamAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_ADD}`;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/beta/teams`,
      resolveWithFullResponse: true,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata.metadata=none'
      },
      body: {
        'template@odata.bind': 'https://graph.microsoft.com/beta/teamsTemplates/standard',
        displayName: args.options.name,
        description: args.options.description
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Display name for the Microsoft Teams team.'
      },
      {
        option: '-d, --description <description>',
        description: 'Description for the Microsoft Teams team.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return `Required parameter name missing`
      }

      if (!args.options.description) {
        return `Required parameter description missing`
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.

    If you want to add a Team to an existing Office 365 Group use the
    ${chalk.blue(aadcommands.O365GROUP_TEAMIFY)} command instead.

  Examples:
  
    Add a new Microsoft Teams team 
      ${this.name} --name 'Architecture' --description 'Architecture Discussion'

  `);
  }
}

module.exports = new TeamsTeamAddCommand();