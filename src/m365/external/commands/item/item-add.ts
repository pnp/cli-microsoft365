import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  externalConnectionId: string;
  content: string;
  contentType?: string;
  acls: string;
}

class ExternalItemAddCommand extends GraphCommand {
  private static contentType: string[] = ['text', 'html'];

  public get name(): string {
    return commands.ITEM_ADD;
  }

  public get description(): string {
    return 'Creates external item';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        contentType: typeof args.options.contentType
      });
      this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--externalConnectionId <externalConnectionId>'
      },
      {
        option: '--content <content>'
      },
      {
        option: '--contentType [contentType]',
        autocomplete: ExternalItemAddCommand.contentType
      },
      {
        option: '--acls <acls>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.contentType &&
          ExternalItemAddCommand.contentType.indexOf(args.options.contentType) < 0) {
          return `${args.options.contentType} is not a valid value for contentType. Allowed values are ${ExternalItemAddCommand.contentType.join(', ')}`;
        }

        // verify that each value for ACLs consists of three parts
        // and that the values are correct
        const acls: string[] = args.options.acls.split(';');
        for (let i = 0; i < acls.length; i++) {
          const acl: string[] = acls[i].split(',');
          if (acl.length !== 3) {
            return `The value ${acls[i]} for option acls is not in the correct format. The correct format is "accessType,type,value", eg. "grant,everyone,everyone"`;
          }

          const accessTypeValues = ['grant', 'deny'];
          if (accessTypeValues.indexOf(acl[0]) < 0) {
            return `The value ${acl[0]} for option acls is not valid. Allowed values are ${accessTypeValues.join(', ')}}`;
          }

          const aclTypeValues = ['user', 'group', 'everyone', 'everyoneExceptGuests', 'externalGroup'];
          if (aclTypeValues.indexOf(acl[1]) < 0) {
            return `The value ${acl[1]} for option acls is not valid. Allowed values are ${aclTypeValues.join(', ')}}`;
          }
        }

        return true;
      }
    );
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
    this.addUnknownOptionsToPayload(requestBody.properties, args.options);

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
