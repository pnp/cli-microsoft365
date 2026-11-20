import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import type GlobalOptions from '../../../../GlobalOptions.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoServicePrincipalPermissionRequestListCommand from './serviceprincipal-permissionrequest-list.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  all: z.boolean().optional(),
  resource: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoServicePrincipalPermissionRequestApproveCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE;
  }

  public get description(): string {
    return 'Approves the specified permission request';
  }

  public alias(): string[] | undefined {
    return [commands.SP_PERMISSIONREQUEST_APPROVE];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.all, opts.resource].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of id, all, or resource',
        params: {
          customCode: 'optionSet',
          options: ['id', 'all', 'resource']
        }
      })
      .superRefine((opts, ctx) => {
        if (opts.id && !validation.isValidGuid(opts.id)) {
          ctx.addIssue({
            code: 'custom',
            message: `The value '${opts.id}' is not a valid GUID.`,
            path: ['id']
          });
        }
      }) as z.ZodObject<any>;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving request digest...`);
      }

      const permissionRequestIds = await this.getAllPendingPermissionRequests(args);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      const response: any = [];

      await permissionRequestIds.reduce(async (previousPromise, nextPermissionRequestId) => {
        return previousPromise.then(() => {
          return this.approvePermissionRequest(nextPermissionRequestId, reqDigest, spoAdminUrl).then(result => response.push(result));
        });
      }, Promise.resolve());

      await logger.log(response.length === 1 ? response[0] : response);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getAllPendingPermissionRequests(args: CommandArgs): Promise<string[]> {
    if (args.options.id) {
      return [args.options.id];
    }
    else {
      const options: GlobalOptions = {
        debug: this.debug,
        verbose: this.verbose
      };

      const output = await cli.executeCommandWithOutput(spoServicePrincipalPermissionRequestListCommand as Command, { options: { ...options, _: [] } });
      const getPermissionRequestsOutput = JSON.parse(output.stdout);
      if (args.options.resource) {
        return getPermissionRequestsOutput.filter((x: any) => x.Resource === args.options.resource).map((x: any) => { return x.Id; });
      }
      return getPermissionRequestsOutput.map((x: any) => { return x.Id; });
    }
  }

  private async approvePermissionRequest(permissionRequestId: string, reqDigest: FormDigestInfo, spoAdminUrl: string): Promise<any> {
    const requestOptions: any = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': reqDigest.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(permissionRequestId)}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`
    };

    const res = await request.post<string>(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];
    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }
    else {
      const output: any = json[json.length - 1];
      delete output._ObjectType_;
      return output;
    }
  }
}

export default new SpoServicePrincipalPermissionRequestApproveCommand();