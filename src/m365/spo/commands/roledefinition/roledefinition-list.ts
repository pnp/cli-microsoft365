import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoRoleDefinitionListCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_LIST;
  }

  public get description(): string {
    return 'Gets list of role definitions for the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Getting role definitions list from ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/roledefinitions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((response: { value: any[] }): void => {
        logger.log(response.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoRoleDefinitionListCommand();