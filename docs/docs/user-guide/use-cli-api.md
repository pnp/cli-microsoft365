# Use CLI for Microsoft 365 programmatically

Typically, you'll work with CLI for Microsoft 365 in a command line. You'll either call specific commands or build automation scripts to combine multiple tasks. But if  you're building software, you might want to use CLI for Microsoft 365 from your code.

## Integrate CLI for Microsoft 365 in your app

If you build apps in Node.js, you can integrate CLI for Microsoft 365 using its API. This API lets you call any of CLI's commands. The following example shows how you could call several CLI for Microsoft 365 commands in a Node.js app:

```javascript
const { executeCommand } = require('@pnp/cli-microsoft365');

executeCommand('status', { output: 'text' })
  .then(res => {
    if (res.stdout === 'Logged out') {
      return executeCommand('login', { output: 'text' }, {
        stdout: message => console.log(message)
      });
    }

    return Promise.resolve();
  })
  .then(_ => executeCommand('spo site list', { output: 'json' }))
  .then(res => {
    const sites = JSON.parse(res.stdout);

    if (sites.length === 0) {
      return Promise.reject('No sites found');
    }

    const siteUrl = sites[0].Url;
    return executeCommand('spo web get', { webUrl: siteUrl, output: 'json' });
  })
  .then(res => console.log(res.stdout))
  .catch(err => console.error(err));
```

You start with importing the `executeCommand` function from CLI for Microsoft 365. CLI doesn't expose all of its logic externally, but rather just the function that allows you to run CLI's commands. This could change in the future.

Next, you execute a command by passing the command's name without the `m365` prefix, and its options. After the command completed its execution, it resolves a Promise with the command's output. The `stdout` property contains the main command output. The `stderr` property would contain verbose, debug and error output that in command line would be sent to stderr. The output in the Promise is a string in the format specified in the `output` option passed to `executeCommand`.

In some cases, like when calling the `login` command, you might need to get the command output while it executes. In the case of the `login` command, it will contain the instructions to complete the device login flow. You can get this output by passing to the `executeCommand` function a third argument with listeners attached to stdout and stderr output.

```javascript
executeCommand('login', { output: 'text' }, {
  stdout: message => console.log(message)
});
```

!!! important
    You shouldn't use both listeners and output from Promises. All command output is sent to the registered listeners and exposed in the end through the resolved Promise. If you would send output from both the listener and Promise to the console, you'd end up with the same output printed twice. In the code sample above you see that for all commands you work with the output from Promises but for the `login` command you use a listener because you want to get login instructions while the command is still running.
