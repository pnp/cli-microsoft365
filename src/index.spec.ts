import assert from 'assert';
import fs from 'fs';
import path from 'path';
import url from 'url';
import * as aadCommands from './m365/entra/commands.js';
import * as cliCommands from './m365/cli/commands.js';
import * as globalCommands from './m365/commands/commands.js';
import * as flowCommands from './m365/flow/commands.js';
import * as graphCommands from './m365/graph/commands.js';
import * as oneDriveCommands from './m365/onedrive/commands.js';
import * as outlookCommands from './m365/outlook/commands.js';
import * as paCommands from './m365/pa/commands.js';
import * as ppCommands from './m365/pp/commands.js';
import * as plannerCommands from './m365/planner/commands.js';
import * as externalCommands from './m365/external/commands.js';
import * as spfxCommands from './m365/spfx/commands.js';
import * as spoCommands from './m365/spo/commands.js';
import * as teamsCommands from './m365/teams/commands.js';
import * as tenantCommands from './m365/tenant/commands.js';
import * as utilCommands from './m365/util/commands.js';
import * as yammerCommands from './m365/yammer/commands.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

describe('Lazy loading commands', () => {
  it('has all commands stored in correct paths that allow lazy loading', () => {
    const commandCollections: any[] = [
      globalCommands.default,
      aadCommands.default,
      cliCommands.default,
      flowCommands.default,
      graphCommands.default,
      oneDriveCommands.default,
      outlookCommands.default,
      paCommands.default,
      ppCommands.default,
      plannerCommands.default,
      externalCommands.default,
      spfxCommands.default,
      spoCommands.default,
      teamsCommands.default,
      tenantCommands.default,
      utilCommands.default,
      yammerCommands.default
    ];
    const aliases: string[] = [
      'consent',
      'flow connector export',
      'flow connector list',
      'search externalconnection add',
      'search externalconnection get',
      'search externalconnection list',
      'search externalconnection remove',
      'search externalconnection schema add',
      'spo page template remove',
      'spo sp grant add',
      'spo sp grant list',
      'spo sp grant revoke',
      'spo sp permissionrequest approve',
      'spo sp permissionrequest deny',
      'spo sp permissionrequest list',
      'spo sp set',
      'teams user add',
      'teams user list',
      'teams user remove',
      'teams user set'
    ];
    const allCommandNames: string[] = [];

    commandCollections.forEach(commandsCollection => {
      for (const commandName in commandsCollection) {
        allCommandNames.push(commandsCollection[commandName]);
      }
    });

    const incorrectFilePaths: string[] = [];
    allCommandNames.forEach(commandName => {
      if (aliases.indexOf(commandName) > -1) {
        // aliases can't be resolved to file names
        return;
      }

      const words: string[] = commandName.split(' ');
      let commandFilePath: string = '';
      if (words.length === 1) {
        commandFilePath = path.join('m365', 'commands', `${commandName}.js`);
      }
      else {
        if (words.length === 2) {
          commandFilePath = path.join('m365', words[0], 'commands', `${words.join('-')}.js`);
        }
        else {
          commandFilePath = path.join('m365', words[0], 'commands', words[1], words.slice(1).join('-') + '.js');
        }
      }

      if (!fs.existsSync(path.join(__dirname, commandFilePath))) {
        incorrectFilePaths.push(commandFilePath);
      }
    });

    assert.deepStrictEqual(incorrectFilePaths, []);
  });
});
