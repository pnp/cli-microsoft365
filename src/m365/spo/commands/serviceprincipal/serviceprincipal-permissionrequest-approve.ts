import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoServicePrincipalPermissionRequestListCommand from './serviceprincipal-permissionrequest-list';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  all?: boolean;
  resource?: string;
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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTelemetry();
    this.#initOptionSets();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '--all'
      },
      {
        option: '--resource [resource]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        resource: typeof args.options.resource !== 'undefined',
        all: !!args.options.all
      });
    });
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'all', 'resource'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      const permissionRequestIds = await this.getAllPendingPermissionRequests(args);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      const response: any = [];

      await permissionRequestIds.reduce(async (previousPromise, nextPermissionRequestId) => {
        return previousPromise.then(() => {
          return this.approvePermissionRequest(nextPermissionRequestId, reqDigest, spoAdminUrl).then(result => response.push(result));
        });
      }, Promise.resolve());

      logger.log(response.length === 1 ? response[0] : response);
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

      const output = await Cli.executeCommandWithOutput(SpoServicePrincipalPermissionRequestListCommand as Command, { options: { ...options, _: [] } });
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

module.exports = new SpoServicePrincipalPermissionRequestApproveCommand();