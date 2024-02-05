import { Group } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id: string;
  displayName?: string;
  description?: string;
  owners?: string;
  members?: string;
  isPrivate?: boolean;
  logoPath?: string;
}

class EntraM365GroupSetCommand extends GraphCommand {
  private static numRepeat: number = 15;
  private pollingInterval: number = 500;

  public get name(): string {
    return commands.M365GROUP_SET;
  }

  public get description(): string {
    return 'Updates Microsoft 365 Group properties';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_SET];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        owners: typeof args.options.owners !== 'undefined',
        members: typeof args.options.members !== 'undefined',
        isPrivate: args.options.isPrivate,
        logoPath: typeof args.options.logoPath !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--owners [owners]'
      },
      {
        option: '--members [members]'
      },
      {
        option: '--isPrivate [isPrivate]',
        autocomplete: ['true', 'false']
      },
      {
        option: '-l, --logoPath [logoPath]'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('isPrivate');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.displayName &&
          !args.options.description &&
          !args.options.members &&
          !args.options.owners &&
          typeof args.options.isPrivate === 'undefined' &&
          !args.options.logoPath) {
          return 'Specify at least one property to update';
        }

        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.owners) {
          const owners: string[] = args.options.owners.split(',').map(o => o.trim());
          for (let i = 0; i < owners.length; i++) {
            if (owners[i].indexOf('@') < 0) {
              return `${owners[i]} is not a valid userPrincipalName`;
            }
          }
        }

        if (args.options.members) {
          const members: string[] = args.options.members.split(',').map(m => m.trim());
          for (let i = 0; i < members.length; i++) {
            if (members[i].indexOf('@') < 0) {
              return `${members[i]} is not a valid userPrincipalName`;
            }
          }
        }

        if (args.options.logoPath) {
          const fullPath: string = path.resolve(args.options.logoPath);

          if (!fs.existsSync(fullPath)) {
            return `File '${fullPath}' not found`;
          }

          if (fs.lstatSync(fullPath).isDirectory()) {
            return `Path '${fullPath}' points to a directory`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.showDeprecationWarning(logger, aadCommands.M365GROUP_SET, commands.M365GROUP_SET);

    try {
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(args.options.id);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.id}' is not a Microsoft 365 group.`);
      }

      if (args.options.displayName || args.options.description || typeof args.options.isPrivate !== 'undefined') {
        if (this.verbose) {
          await logger.logToStderr(`Updating Microsoft 365 Group ${args.options.id}...`);
        }

        const update: Group = {};
        if (args.options.displayName) {
          update.displayName = args.options.displayName;
        }
        if (args.options.description) {
          update.description = args.options.description;
        }
        if (typeof args.options.isPrivate !== 'undefined') {
          update.visibility = args.options.isPrivate ? 'Private' : 'Public';
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${args.options.id}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: update
        };

        await request.patch(requestOptions);
      }

      if (args.options.logoPath) {
        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          await logger.logToStderr(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${args.options.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        await this.setGroupLogo(requestOptions, EntraM365GroupSetCommand.numRepeat, logger);
      }
      else if (this.debug) {
        await logger.logToStderr('logoPath not set. Skipping');
      }

      if (args.options.owners) {
        const owners: string[] = args.options.owners.split(',').map(o => o.trim());

        if (this.verbose) {
          await logger.logToStderr('Retrieving user information to set group owners...');
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        const res = await request.get<{ value: { id: string; }[] }>(requestOptions);

        await Promise.all(res.value.map(u => request.post({
          url: `${this.resource}/v1.0/groups/${args.options.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      }
      else if (this.debug) {
        await logger.logToStderr('Owners not set. Skipping');
      }

      if (args.options.members) {
        const members: string[] = args.options.members.split(',').map(o => o.trim());

        if (this.verbose) {
          await logger.logToStderr('Retrieving user information to set group members...');
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/users?$filter=${members.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        const res = await request.get<{ value: { id: string; }[] }>(requestOptions);

        await Promise.all(res.value.map(u => request.post({
          url: `${this.resource}/v1.0/groups/${args.options.id}/members/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      }
      else if (this.debug) {
        await logger.logToStderr('Members not set. Skipping');
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async setGroupLogo(requestOptions: any, retryLeft: number, logger: Logger): Promise<void> {
    try {
      await request.put(requestOptions);
    }
    catch (err: any) {
      if (--retryLeft > 0) {
        await setTimeout(this.pollingInterval * (EntraM365GroupSetCommand.numRepeat - retryLeft));
        await this.setGroupLogo(requestOptions, retryLeft, logger);
      }
      else {
        throw err;
      }
    }
  }

  private getImageContentType(imagePath: string): string {
    const extension: string = imagePath.substr(imagePath.lastIndexOf('.')).toLowerCase();

    switch (extension) {
      case '.png':
        return 'image/png';
      case '.gif':
        return 'image/gif';
      default:
        return 'image/jpeg';
    }
  }
}

export default new EntraM365GroupSetCommand();
