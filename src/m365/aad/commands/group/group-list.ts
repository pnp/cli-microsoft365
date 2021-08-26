import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Group } from './Group';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  output?: string;
}

class AadGroupListCommand extends GraphItemsListCommand<Group>   {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all Azure AD groups in the tenant.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mailNickname','groupTypes', 'securityEnabled','mailEnabled','visibility'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/groups`;
    this
      .getAllItems(endpoint, logger, true)
      .then((): Promise<any> => {
        return Promise.resolve();
      })
      .then((): void => {
        logger.log(this.items);
        if (this.verbose) {
          logger.logToStderr(chalk.green("DONE"));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadGroupListCommand();