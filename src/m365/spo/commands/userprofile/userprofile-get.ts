import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
}

class SpoUserProfileGetCommand extends SpoCommand {
  public get name(): string {
    return commands.USERPROFILE_GET;
  }

  public get description(): string {
    return 'Sets user profile property for a SharePoint user';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --userName <userName>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<{ UserProfileProperties: { Key: string; Value: string }[] }> => {
        const userName: string = `i:0#.f|membership|${args.options.userName}`;
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(`${userName}`)}'`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        return request.get<{ UserProfileProperties: { Key: string; Value: string }[] }>(requestOptions);
      })
      .then((res: { UserProfileProperties: { Key: string; Value: string }[] }): void => {
        // in text mode, reformat properties for readability
        if (!args.options.output ||
          args.options.output === 'text') {
          res.UserProfileProperties = JSON.stringify(res.UserProfileProperties) as any;
        }

        logger.log(res);

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}
module.exports = new SpoUserProfileGetCommand();
