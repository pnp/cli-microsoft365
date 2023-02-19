import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  associatedGroup?: string;
}

class SpoGroupGetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_GET;
  }

  public get description(): string {
    return 'Gets site group';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        name: (!(!args.options.name)).toString(),
        associatedGroup: args.options.associatedGroup
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--associatedGroup [associatedGroup]',
        autocomplete: ['Owner', 'Member', 'Visitor']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && isNaN(args.options.id)) {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.associatedGroup && ['owner', 'member', 'visitor'].indexOf(args.options.associatedGroup.toLowerCase()) === -1) {
          return `${args.options.associatedGroup} is not a valid associatedGroup value. Allowed values are Owner|Member|Visitor.`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name', 'associatedGroup'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for group in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetById('${args.options.id}')`;
    }
    else if (args.options.name) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(args.options.name as string)}')`;
    }
    else if (args.options.associatedGroup) {
      switch (args.options.associatedGroup.toLowerCase()) {
        case 'owner':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedOwnerGroup`;
          break;
        case 'member':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedMemberGroup`;
          break;
        case 'visitor':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedVisitorGroup`;
          break;
      }
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const groupInstance = await request.get(requestOptions);
      logger.log(groupInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoGroupGetCommand();
