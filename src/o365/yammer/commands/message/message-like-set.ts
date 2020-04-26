import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import YammerCommand from "../../../base/YammerCommand";
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  enable?: boolean;
  confirm?: boolean;
}

class YammerMessageLikeSetCommand extends YammerCommand {
  constructor() {
    super();
  }

  public get name(): string {
    return `${commands.YAMMER_MESSAGE_LIKE_SET}`;
  }

  public get description(): string {
    return 'Adds or removes a like from a message';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = args.options.id !== undefined;
    telemetryProps.enable = args.options.enable !== undefined;
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const executeLikeAction: () => void = (): void => {
        let endpoint = `${this.resource}/v1/messages/liked_by/current.json`;
        const requestOptions: any = {
          url: endpoint,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          },
          json: true,
          body: {
            message_id: args.options.id
          }
        };

        if (args.options.enable != false) {
          request
              .post(requestOptions)
              .then((res: any): void => {
                cb();
              }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
        }
        else {
          request
              .delete(requestOptions)
              .then((res: any): void => {
                cb();
              }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
        }
      };

      if (args.options.confirm || args.options.enable != false) {
        executeLikeAction();
      }
      else {
        let messagePrompt = `Are you sure you want to remove the like from the message ${args.options.id}?`;
    
        cmd.prompt({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: messagePrompt,
        }, (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            executeLikeAction();
          }
        });
      }
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The id of the Yammer message'
      },
      {
        option: '--enable [enable]',
        description: 'Add or remove a like from a message. Default true'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation before a like is being removed'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required id value is missing';
      }

      if (args.options.id && typeof args.options.id !== 'number') {
        return `${args.options.id} is not a number`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    ${chalk.yellow('Attention:')} In order to use this command, you need to grant the Azure AD
    application used by the Office 365 CLI the permission to the Yammer API.
    To do this, execute the ${chalk.blue('consent --service yammer')} command.
  
  Examples:
    
    Likes the message with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081
    
    Removes a like from the message with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081 --enable false

    Removes a like from the message with the ID ${chalk.grey('5611239081')} without asking for confirmation
      ${this.name} --id 5611239081 --enable false --confirm
  `);
  }
}

module.exports = new YammerMessageLikeSetCommand();