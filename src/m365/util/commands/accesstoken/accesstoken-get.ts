import { z } from 'zod';
import auth, { Auth } from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import { accessToken } from '../../../../utils/accessToken.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  resource: z.string().alias('r'),
  new: z.boolean().optional(),
  decoded: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class UtilAccessTokenGetCommand extends Command {
  public get name(): string {
    return commands.ACCESSTOKEN_GET;
  }

  public get description(): string {
    return 'Gets access token for the specified resource';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let resource: string = args.options.resource;

    if (resource.toLowerCase() === 'sharepoint') {
      if (auth.connection.spoUrl) {
        resource = auth.connection.spoUrl;
      }
      else {
        throw `SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`;
      }
    }
    else if (resource.toLowerCase() === 'graph') {
      resource = Auth.getEndpointForResource('https://graph.microsoft.com', auth.connection.cloudType);
    }

    try {
      const token: string = await auth.ensureAccessToken(resource, logger, this.debug, args.options.new);

      if (args.options.decoded) {
        const { header, payload } = accessToken.getDecodedAccessToken(token);

        await logger.logRaw(`${JSON.stringify(header, null, 2)}.${JSON.stringify(payload, null, 2)}.[signature]`);
      }
      else {
        await logger.log(token);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new UtilAccessTokenGetCommand();