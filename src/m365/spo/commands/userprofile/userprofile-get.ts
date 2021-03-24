import { ContextInfo } from '../../spo';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { Logger } from '../../../../cli';
import Utils from '../../../../Utils';
import {
  CommandOption,
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
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
  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoUrl: string = '';
    this
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((response: ContextInfo): Promise<string> => {
        let uName: string = `i:0#.f|membership|${args.options.userName}`
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(`${uName}`)}'`,
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-type": "application/json;odata=verbose",
            "X-RequestDigest": response.FormDigestValue
          },
          responseType: 'json'
        };
        return request.get(requestOptions);
      })
      .then((res: any): void => {
        if(args.options.output === 'json'){
            logger.log(res);
         }
         else{
          logger.log(res.UserProfileProperties.map((property: any) => {
            return {
              key: property.Key,
              Value: property.Value
            };
          }));
         }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --userName <userName>'
      },
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidUserPrincipalName(args.options.userName)) {
      return `${args.options.userName} is not a valid user principal name`;
    }

    return true;
  }
}
module.exports = new SpoUserProfileGetCommand();
