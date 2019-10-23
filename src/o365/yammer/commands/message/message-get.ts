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
}

class YammerMessageGetCommand extends YammerCommand {
  public get name(): string {
    return `${commands.YAMMER_MESSAGE_GET}`;
  }

  public get description(): string {
    return 'Returns a yammer message';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== undefined;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/messages/${args.options.id}.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      json: true,
      body: {
      }
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          cmd.log({
            id: res.id, 
            sender_id: res.sender_id, 
            replied_to_id: res.replied_to_id, 
            thread_id: res.thread_id, 
            group_id: res.group_id,
            created_at: res.created_at,
            direct_message: res.direct_message,   
            system_message: res.system_message,
            privacy: res.privacy,
            message_type: res.message_type,
            content_excerpt: res.content_excerpt
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The id of the yammer message'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required option id missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Returns the yammer message with the id 1239871123
      ${this.name} --id 1239871123

    Returns the yammer message with the id 1239871123 in JSON format
      ${this.name} --id 1239871123 --output json
    `);
  }
}

module.exports = new YammerMessageGetCommand();