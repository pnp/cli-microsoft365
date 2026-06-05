import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const contentTypes = ['text', 'html'] as const;

export const options = z.object({
  ...globalOptionsZod.shape,
  id: z.string(),
  externalConnectionId: z.string(),
  content: z.string(),
  contentType: z.enum(contentTypes).optional(),
  acls: z.string()
    .refine(val => {
      const acls = val.split(';');
      return acls.every(acl => acl.split(',').length === 3);
    }, {
      message: 'The value for option acls is not in the correct format. The correct format is "accessType,type,value", eg. "grant,everyone,everyone"'
    })
    .refine(val => {
      const acls = val.split(';');
      const accessTypeValues = ['grant', 'deny'];
      return acls.every(acl => accessTypeValues.includes(acl.split(',')[0]));
    }, {
      message: 'The accessType value for option acls is not valid. Allowed values are grant, deny'
    })
    .refine(val => {
      const acls = val.split(';');
      const aclTypeValues = ['user', 'group', 'everyone', 'everyoneExceptGuests', 'externalGroup'];
      return acls.every(acl => {
        const parts = acl.split(',');
        return parts.length >= 2 && aclTypeValues.includes(parts[1]);
      });
    }, {
      message: 'The type value for option acls is not valid. Allowed values are user, group, everyone, everyoneExceptGuests, externalGroup'
    })
}).passthrough();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExternalItemAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ITEM_ADD;
  }

  public get description(): string {
    return 'Creates external item';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const acls: ExternalConnectors.Acl[] = args.options.acls
      .split(';')
      .map(acl => {
        const aclParts: string[] = acl.split(',');
        return {
          accessType: aclParts[0] as any,
          type: aclParts[1] as any,
          value: aclParts[2]
        };
      });

    const requestBody: ExternalConnectors.ExternalItem = {
      id: args.options.id,
      content: {
        value: args.options.content,
        type: args.options.contentType as any ?? 'text'
      },
      acl: acls,
      properties: {}
    };

    // we need to rewrite the @odata properties to the correct format
    // to extract multiple values for collections into arrays
    this.rewriteCollectionProperties(args.options);
    this.addUnknownOptionsToPayloadZod(requestBody.properties, args.options);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}/items/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: requestBody
    };

    try {
      const externalItem: any = await request.put(requestOptions);

      if (args.options.output === 'csv' || args.options.output === 'md') {
        // for CSV and md, we need to bring the properties to the main object
        // and convert arrays to comma-separated strings or they will be dropped
        // from the output
        Object.getOwnPropertyNames(externalItem.properties).forEach(name => {
          if (Array.isArray(externalItem.properties[name])) {
            externalItem[name] = externalItem.properties[name].join(', ');
          }
          else {
            externalItem[name] = externalItem.properties[name];
          }
        });
      }

      await logger.log(externalItem);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private rewriteCollectionProperties(options: any): void {
    Object.getOwnPropertyNames(options).forEach(name => {
      if (!name.includes('@odata')) {
        return;
      }

      // convert the value of a collection to an array
      const nameWithoutOData: string = name.substring(0, name.indexOf('@odata'));
      if (options[nameWithoutOData]) {
        options[nameWithoutOData] = options[nameWithoutOData].split(';#');
      }
    });
  }
}

export default new ExternalItemAddCommand();
