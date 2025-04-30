#!/usr/bin/env node

import confirm from '@inquirer/confirm';
import select from '@inquirer/select';
import input from '@inquirer/input';
import clipboardy from 'clipboardy';

const RESOURCE_CHOICES = [
  { value: 'Azure Service Management', name: 'Azure Service Management' },
  { value: 'Dynamics CRM', name: 'Dynamics CRM' },
  { value: 'Microsoft Graph', name: 'Microsoft Graph' },
  { value: 'Office 365 Management APIs', name: 'Office 365 Management APIs' },
  { value: 'Power BI Service', name: 'Power BI Service' },
  { value: 'PowerApps Service', name: 'PowerApps Service' },
  { value: 'SharePoint', name: 'SharePoint' },
  { value: '_other', name: 'Other - specify…' },
  { value: '_done', name: '✓  Done adding services' }
];

async function pickResource(message) {
  let resource = await select({ message, choices: RESOURCE_CHOICES, pageSize: RESOURCE_CHOICES.length });
  if (resource === '_done') {
    return null;
  }
  if (resource === '_other') {
    resource = await input({ message: 'Specify the resource name:' });
  }
  return resource;
}

async function collectPermissions(kind /* 'delegated' | 'application' */) {
  const list = [];
  while (true) {
    console.log('');
    const resource = await pickResource(`Choose a resource to add ${kind} permissions for:`);
    if (!resource){
      break;
    }
    let permissions = await input({ message: `Enter the ${kind} permissions for ${resource}:` });
    permissions = permissions.split(',').map(p => p.trim()).join(', ');
    list.push({ resource, permissions });
  }
  return list;
}

function constructMarkdownTable(list) {
  const COLS = ['resource', 'permissions'];

  // Get longest piece of text per column, headers included
  const widths = COLS.map(key =>
    Math.max(
      key.length,
      ...list.map(r => r[key].length)
    )
  );

  // Ensure that cells start with a capital letter and are padded
  const cell = (text, col) => ` ${(text.charAt(0).toUpperCase() + text.slice(1)).padEnd(widths[col])} `;

  const header   = `|${COLS.map((h, i) => cell(h, i)).join('|')}|`;
  const divider  = `  |${widths.map(w => '-'.repeat(w + 2)).join('|')}|`;
  const dataRows = list.map(r => `  |${cell(r.resource, 0)}|${cell(r.permissions, 1)}|`);

  return [header, divider, ...dataRows].join('\n');
}

async function main() {
  console.log('\nThis tool will build a well-formatted Markdown permissions section.\n');

  const permissions = {
    delegated: [],
    application: []
  };

  if (await confirm({ message: 'Does the command support delegated permissions?', default: true })) {
    permissions.delegated = await collectPermissions('delegated');
  }

  if (await confirm({ message: 'Does the command support application permissions?', default: true })) {
    if (permissions.delegated.length) {
      // Re-use the same resource first
      for (const { resource } of permissions.delegated) {
        let scopes = await input({ message: `Enter the application permission for ${resource}:` });
        scopes = scopes.split(',').map(p => p.trim()).join(', ');
        permissions.application.push({ resource, permissions: scopes });
      }
    }
    else {
      // Allow adding extra app-only resources
      permissions.application.push(...await collectPermissions('application'));
    }
  }

  // Generate the Markdown output
  let delegatedTable = '';
  if (permissions.delegated.length === 0) {
    delegatedTable = 'This command does not support delegated permissions.';
  }
  else {
    permissions.delegated.sort((a, b) => a.resource.toLowerCase().localeCompare(b.resource.toLowerCase()));
    delegatedTable = constructMarkdownTable(permissions.delegated);
  }

  let applicationTable = '';
  if (permissions.application.length === 0) {
    applicationTable = 'This command does not support application permissions.';
  }
  else {
    permissions.application.sort((a, b) => a.resource.toLowerCase().localeCompare(b.resource.toLowerCase()));
    applicationTable = constructMarkdownTable(permissions.application);
  }

  const output = `import Tabs from '@theme/Tabs';
import TabItem from '@theme/TabItem';

## Permissions

<Tabs>
  <TabItem value="Delegated">

  ${delegatedTable}

  </TabItem>
  <TabItem value="Application">

  ${applicationTable}

  </TabItem>
</Tabs>`;

  console.log('Your permissions section:\n');
  console.log(output + '\n');
  const copyOutput = await confirm({ message: 'Do you want to copy the output to the clipboard?', default: true });
  if (copyOutput) {
    await clipboardy.write(output);
    console.log('✅ Copied to clipboard!');
  }
}

main()
  .catch(err => {
    console.error('❌  Unexpected error:', err);
    process.exitCode = 1;
  });