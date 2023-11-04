import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ProfileCardProperty } from './profileCardProperties.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: GlobalOptions;
}

class TenantPeopleProfileCardPropertyListCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_LIST;
  }

  public get description(): string {
    return 'Lists all profile card properties';
  }

  constructor() {
    super();
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Lists all profile card properties...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const result = await request.get<ProfileCardProperty[]>(requestOptions);
      let output: any = result;

      if (args.options.output && args.options.output !== 'json') {
        output = output.value.map((p: ProfileCardProperty) => {
          const propertyAnnotations = p.annotations[0]?.localizations?.map((l) => {
            return { ['displayName ' + l.languageTag]: l.displayName };
          }) ?? [];

          const propertyAnnotationsObject = Object.assign({}, ...propertyAnnotations);

          const result: any = { directoryPropertyName: p.directoryPropertyName };
          if (p.annotations[0]?.displayName) {
            result.displayName = p.annotations[0]?.displayName;
          }

          return {
            ...result,
            ...propertyAnnotationsObject
          };
        });
      }

      await logger.log(output);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantPeopleProfileCardPropertyListCommand();