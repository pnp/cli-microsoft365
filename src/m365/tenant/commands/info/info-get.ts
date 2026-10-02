import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  domainName: z.string().optional().alias('d'),
  tenantId: z.string().optional().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantInfoGetCommand extends GraphCommand {
  public get name(): string {
    return commands.INFO_GET;
  }

  public get description(): string {
    return 'Gets information about any tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .superRefine((opts, ctx) => {
        if (opts.tenantId && !validation.isValidGuid(opts.tenantId)) {
          ctx.addIssue({
            code: 'custom',
            message: `${opts.tenantId} is not a valid GUID`
          });
        }
      })
      .refine(
        opts => !(opts.tenantId && opts.domainName),
        {
          message: 'Specify either domainName or tenantId but not both',
          params: {
            customCode: 'optionSet',
            options: ['domainName', 'tenantId']
          }
        }
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let domainName: string | undefined = args.options.domainName;
    const tenantId: string | undefined = args.options.tenantId;

    if (!domainName && !tenantId) {
      const userName: string = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      domainName = userName.split('@')[1];
    }

    let requestUrl = `${this.resource}/v1.0/tenantRelationships/`;

    if (tenantId) {
      requestUrl += `findTenantInformationByTenantId(tenantId='${formatting.encodeQueryParameter(tenantId)}')`;
    }
    else {
      requestUrl += `findTenantInformationByDomainName(domainName='${formatting.encodeQueryParameter(domainName!)}')`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantInfoGetCommand();