import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userName: z.string().refine(userName => validation.isValidUserPrincipalName(userName), {
    error: e => `${e.input} is not a valid user principal name`
  }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoUserProfileGetCommand extends SpoCommand {
  public get name(): string {
    return commands.USERPROFILE_GET;
  }

  public get description(): string {
    return 'Gets SharePoint user profile properties for the specified user';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const userName: string = `i:0#.f|membership|${args.options.userName}`;
      const requestOptions: any = {
        url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${formatting.encodeQueryParameter(`${userName}`)}'`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res: { UserProfileProperties: { Key: string; Value: string }[] } = await request.get<{ UserProfileProperties: { Key: string; Value: string }[] }>(requestOptions);
      if (!args.options.output || cli.shouldTrimOutput(args.options.output)) {
        res.UserProfileProperties = JSON.stringify(res.UserProfileProperties) as any;
      }

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}
export default new SpoUserProfileGetCommand();
