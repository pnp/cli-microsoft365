import assert from 'assert';
import fs from 'fs';
import path from 'path';
import url from 'url';
import * as adaptiveCardCommands from './m365/adaptivecard/commands.js';
import * as appCommands from './m365/app/commands.js';
import * as bookingCommands from './m365/booking/commands.js';
import * as cliCommands from './m365/cli/commands.js';
import * as globalCommands from './m365/commands/commands.js';
import * as entraCommands from './m365/entra/commands.js';
import * as fileCommands from './m365/file/commands.js';
import * as flowCommands from './m365/flow/commands.js';
import * as graphCommands from './m365/graph/commands.js';
import * as oneDriveCommands from './m365/onedrive/commands.js';
import * as oneNoteCommands from './m365/onenote/commands.js';
import * as outlookCommands from './m365/outlook/commands.js';
import * as paCommands from './m365/pa/commands.js';
import * as ppCommands from './m365/pp/commands.js';
import * as plannerCommands from './m365/planner/commands.js';
import * as purviewCommands from './m365/purview/commands.js';
import * as externalCommands from './m365/external/commands.js';
import * as skypeCommands from './m365/skype/commands.js';
import * as spfxCommands from './m365/spfx/commands.js';
import * as spoCommands from './m365/spo/commands.js';
import * as teamsCommands from './m365/teams/commands.js';
import * as tenantCommands from './m365/tenant/commands.js';
import * as todoCommands from './m365/todo/commands.js';
import * as vivaCommands from './m365/viva/commands.js';
import * as utilCommands from './m365/util/commands.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

describe('Lazy loading commands', () => {
  it('has all commands stored in correct paths that allow lazy loading', () => {
    const commandCollections: any[] = [
      globalCommands.default,
      adaptiveCardCommands.default,
      appCommands.default,
      bookingCommands.default,
      entraCommands.default,
      cliCommands.default,
      entraCommands.default,
      fileCommands.default,
      flowCommands.default,
      graphCommands.default,
      oneDriveCommands.default,
      oneNoteCommands.default,
      outlookCommands.default,
      paCommands.default,
      ppCommands.default,
      plannerCommands.default,
      purviewCommands.default,
      externalCommands.default,
      skypeCommands.default,
      spfxCommands.default,
      spoCommands.default,
      teamsCommands.default,
      tenantCommands.default,
      todoCommands.default,
      vivaCommands.default,
      utilCommands.default
    ];
    const aliases: string[] = [
      'entra sp add',
      'entra sp get',
      'entra sp list',
      'entra sp remove',
      'entra appregistration add',
      'entra appregistration get',
      'entra appregistration list',
      'entra appregistration remove',
      'entra appregistration set',
      'entra appregistration permission add',
      'entra appregistration role add',
      'entra appregistration role list',
      'entra appregistration role remove',
      'consent',
      'flow connector export',
      'flow connector list',
      'pp chatbot get',
      'pp chatbot list',
      'pp chatbot remove',
      'search externalconnection add',
      'search externalconnection get',
      'search externalconnection list',
      'search externalconnection remove',
      'search externalconnection schema add',
      'spo folder rename',
      'spo site list',
      'spo page template remove',
      'spo tenant homesite list',
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
