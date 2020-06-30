import * as assert from 'assert';
import * as path from 'path';
import * as fs from 'fs';

import * as globalCommands from './o365/commands/commands';
import * as aadCommands from './o365/aad/commands';
import * as cliCommands from './o365/cli/commands';
import * as flowCommands from './o365/flow/commands';
import * as graphCommands from './o365/graph/commands';
import * as oneDriveCommands from './o365/onedrive/commands';
import * as outlookCommands from './o365/outlook/commands';
import * as paCommands from './o365/pa/commands';
import * as plannerCommands from './o365/planner/commands';
import * as spfxCommands from './o365/spfx/commands';
import * as spoCommands from './o365/spo/commands';
import * as teamsCommands from './o365/teams/commands';
import * as tenantCommands from './o365/tenant/commands';
import * as utilCommands from './o365/util/commands';
import * as yammerCommands from './o365/yammer/commands';

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
      plannerCommands.default,
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
      'outlook sendmail',
      'spo site classic remove',
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
      for (var commandName in commandsCollection) {
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
        commandFilePath = path.join('o365', 'commands', `${commandName}.js`);
      }
      else {
        if (words.length === 2) {
          commandFilePath = path.join('o365', words[0], 'commands', `${words.join('-')}.js`);
        }
        else {
          commandFilePath = path.join('o365', words[0], 'commands', words[1], words.slice(1).join('-') + '.js');
        }
      }

      if (!fs.existsSync(path.join(__dirname, commandFilePath))) {
        incorrectFilePaths.push(commandFilePath);
      }
    });

    assert.deepStrictEqual(incorrectFilePaths, []);
  });
});