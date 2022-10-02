import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  propertyName: string;
  propertyValue: string;
}

class SpoUserProfileSetCommand extends SpoCommand {
  public get name(): string {
    return commands.USERPROFILE_SET;
  }

  public get description(): string {
    return 'Sets user profile property for a SharePoint user';
  }

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --userName <userName>'
      },
      {
        option: '-n, --propertyName <propertyName>'
      },
      {
        option: '-v, --propertyValue <propertyValue>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);

      const res: ContextInfo = await spo.getRequestDigest(spoUrl);
      const propertyValue: string[] = args.options.propertyValue.split(',').map(o => o.trim());
      let propertyType: string = 'SetSingleValueProfileProperty';
      const data: any = {
        accountName: `i:0#.f|membership|${args.options.userName}`,
        propertyName: args.options.propertyName
      };

      if (propertyValue.length > 1) {
        propertyType = 'SetMultiValuedProfileProperty';
        data.propertyValues = [...propertyValue];
      }
      else {
        data.propertyValue = propertyValue[0];
      }

      const requestOptions: any = {
        url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/${propertyType}`,
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'X-RequestDigest': res.FormDigestValue
        },
        data: data,
        responseType: 'json'
      };

      await request.post(requestOptions);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoUserProfileSetCommand();
