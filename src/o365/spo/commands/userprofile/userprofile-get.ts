import { ContextInfo } from '../../spo';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');
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
  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoUrl: string = '';
    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((response: ContextInfo): Promise<string> => {
        let uName : string = `i:0#.f|membership|${args.options.userName}`
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(`${uName}`)}'`,
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-type": "application/json;odata=verbose",
            "X-RequestDigest": response.FormDigestValue
         },
         json:true
        };
        return request.get(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cmd.log(res);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --userName <userName>',
        description: 'Account name of the user'
      },
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.userName) {
        return 'Required parameter userName missing';
      }
      return true;
    };
  }
  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
    This command requires tenant admin permissions in case of updating properties other than the current logged user.
  Examples:
  
    Get property of user profile properties
      ${commands.USERPROFILE_GET} --userName 'john.doe@mytenant.onmicrosoft.com'
  `);
  }
}
module.exports = new SpoUserProfileGetCommand();