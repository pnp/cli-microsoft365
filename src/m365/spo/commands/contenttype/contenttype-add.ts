import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoContentTypeGetCommand, { Options as SpoContentTypeGetCommandOptions } from './contenttype-get.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listTitle: z.string().optional().alias('l'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    }).optional(),
  listUrl: z.string().optional(),
  id: z.string().alias('i'),
  name: z.string().alias('n'),
  description: z.string().optional().alias('d'),
  group: z.string().optional().alias('g')
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeAddCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_ADD;
  }

  public get description(): string {
    return 'Adds a new list or site content type';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let parentInfo = '';

      if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
        parentInfo = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';
      }
      else {
        parentInfo = await this.getParentInfo(args.options, logger);
      }

      if (this.verbose) {
        await logger.logToStderr(`Retrieving request digest...`);
      }

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const description: string = args.options.description ?
        `<Property Name="Description" Type="String">${formatting.escapeXml(args.options.description)}</Property>` :
        '<Property Name="Description" Type="Null" />';
      const group: string = args.options.group ?
        `<Property Name="Group" Type="String">${formatting.escapeXml(args.options.group)}</Property>` :
        '<Property Name="Group" Type="Null" />';

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}">${description}${group}<Property Name="Id" Type="String">${formatting.escapeXml(args.options.id)}</Property><Property Name="Name" Type="String">${formatting.escapeXml(args.options.name)}</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method>${parentInfo}</ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const options: SpoContentTypeGetCommandOptions = {
        webUrl: args.options.webUrl,
        listTitle: args.options.listTitle,
        listUrl: args.options.listUrl,
        listId: args.options.listId,
        id: args.options.id,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      try {
        const output = await cli.executeCommandWithOutput(spoContentTypeGetCommand as Command, { options: { ...options, _: [] } });
        if (this.debug) {
          await logger.logToStderr(output.stderr);
        }

        await logger.log(JSON.parse(output.stdout));
      }
      catch (cmdError: any) {
        throw cmdError.error;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getParentInfo(options: Options, logger: Logger): Promise<string> {
    const siteId: string = await spo.getSiteIdBySPApi(options.webUrl, logger, this.verbose);
    const webId: string = await spo.getWebId(options.webUrl, logger, this.verbose);
    const listId: string = options.listId ? options.listId : await spo.getListId(options.webUrl, options.listTitle, options.listUrl, logger, this.verbose);
    return `<Identity Id="5" Name="1a48869e-c092-0000-1f61-81ec89809537|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}" />`;
  }
}

export default new SpoContentTypeAddCommand();