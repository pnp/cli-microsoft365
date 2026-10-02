import { z } from 'zod';
import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  domainName: z.string().optional().alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantIdGetCommand extends Command {
  public get name(): string {
    return commands.ID_GET;
  }

  public get description(): string {
    return 'Gets Microsoft 365 tenant ID for the specified domain';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let domainName: string | undefined = args.options.domainName;
    if (!domainName) {
      const userName: string = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      domainName = userName.split('@')[1];
    }

    const requestOptions: any = {
      url: `https://login.windows.net/${domainName}/.well-known/openid-configuration`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json',
        'x-anonymous': true
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);

      if (res.error) {
        throw res.error_description;
      }

      if (res.token_endpoint) {
        await logger.log(res.token_endpoint.split('/')[3]);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantIdGetCommand();