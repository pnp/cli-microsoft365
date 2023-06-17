import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  name?: string;
  id?: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

class SpoListSensitivityLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SENSITIVITYLABEL_ENSURE;
  }

  public get description(): string {
    return 'Applies a default sensitivity label to the specified document library';
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
        name: typeof args.options.name !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['name', 'id'] },
      { options: ['listId', 'listTitle', 'listUrl'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const sensitivityLabelId: string = await this.getSensitivityLabelId(args, logger);

      if (this.verbose) {
        logger.logToStderr(`Applying a sensitivity label ${sensitivityLabelId} to the document library...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web`;
      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'if-match': '*'
        },
        data: { 'DefaultSensitivityLabelForLibrary': sensitivityLabelId },
        responseType: 'json'
      };

      await request.patch<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSensitivityLabelId(args: CommandArgs, logger: Logger): Promise<string> {
    const { id, name } = args.options;

    if (id) {
      return id;
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving sensitivity label id of ${name}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${formatting.encodeQueryParameter(name!)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);
    const sensitivityLabelItem: { id: string } | undefined = res.value[0];

    if (!sensitivityLabelItem) {
      throw Error(`The specified sensitivity label does not exist`);
    }

    return sensitivityLabelItem.id;
  }
}

module.exports = new SpoListSensitivityLabelEnsureCommand();
