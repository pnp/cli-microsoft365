import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { ProfileCardProperty } from './profileCardProperties.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';

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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Listing all profile card properties...`);
      }

      const result = await odata.getAllItems<ProfileCardProperty>(`${this.resource}/v1.0/admin/people/profileCardProperties`);
      let output: any = result;

      if (args.options.output && args.options.output !== 'json') {
        output = output.sort((n1: ProfileCardProperty, n2: ProfileCardProperty) => {
          const localizations1 = n1.annotations[0]?.localizations?.length ?? 0;
          const localizations2 = n2.annotations[0]?.localizations?.length ?? 0;
          if (localizations1 > localizations2) {
            return -1;
          }
          if (localizations1 < localizations2) {
            return 1;
          }
          return 0;
        });

        output = output.map((p: ProfileCardProperty) => {
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