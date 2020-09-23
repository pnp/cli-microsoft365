import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  confirm?: boolean;
}

class YammerMessageRemoveCommand extends YammerCommand {
  public get name(): string {
    return commands.YAMMER_MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Yammer message';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeMessage: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1/messages/${args.options.id}.json`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .delete(requestOptions)
        .then((res: any): void => {
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }

    if (args.options.confirm) {
      removeMessage();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Yammer message ${args.options.id}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeMessage();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The id of the Yammer message'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the Yammer message'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (typeof args.options.id !== 'number') {
      return `${args.options.id} is not a number`;
    }

    return true;
  }
}

module.exports = new YammerMessageRemoveCommand();
