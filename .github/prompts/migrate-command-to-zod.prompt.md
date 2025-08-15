---
mode: agent
tools: ['codebase', 'editFiles', 'changes']
---
Update the command file to use ZOD. 
- Don't change the command's functionality. 
- Update tests accordingly. 
- Don't remove or add new tests. 
- For updating test files, check how spec.ts files have been modified in the reference implementations. 
- Use the .mdx file if you need additional information about the command's options. 

Use the below git commits as reference implementations for a command and its tests using ZOD. Do not continue with the task if you do not have access to these commits. 

```diff
commit 824c1ebd2dfb0a1eabb623ed6a22da50f4edf61d
Author: waldekmastykarz <waldek@mastykarz.nl>
Date:   Sun Apr 20 09:02:32 2025 +0200

diff --git a/src/m365/booking/commands/business/business-get.spec.ts b/src/m365/booking/commands/business/business-get.spec.ts
index 40b3c3d2a..9c4f21496 100644
--- a/src/m365/booking/commands/business/business-get.spec.ts
+++ b/src/m365/booking/commands/business/business-get.spec.ts
@@ -1,10 +1,13 @@
 import assert from 'assert';
 import sinon from 'sinon';
+import { z } from 'zod';
 import auth from '../../../../Auth.js';
 import { cli } from '../../../../cli/cli.js';
+import { CommandInfo } from '../../../../cli/CommandInfo.js';
 import { Logger } from '../../../../cli/Logger.js';
 import { CommandError } from '../../../../Command.js';
 import request from '../../../../request.js';
+import { settingsNames } from '../../../../settingsNames.js';
 import { telemetry } from '../../../../telemetry.js';
 import { formatting } from '../../../../utils/formatting.js';
 import { pid } from '../../../../utils/pid.js';
@@ -12,7 +15,6 @@ import { session } from '../../../../utils/session.js';
 import { sinonUtil } from '../../../../utils/sinonUtil.js';
 import commands from '../../commands.js';
 import command from './business-get.js';
-import { settingsNames } from '../../../../settingsNames.js';
 
 describe(commands.BUSINESS_GET, () => {
   const validId = 'mail@contoso.onmicrosoft.com';
@@ -31,6 +33,8 @@ describe(commands.BUSINESS_GET, () => {
   let log: string[];
   let logger: Logger;
   let loggerLogSpy: sinon.SinonSpy;
+  let commandInfo: CommandInfo;
+  let commandOptionsSchema: z.ZodTypeAny;
 
   before(() => {
     sinon.stub(auth, 'restoreAuth').resolves();
@@ -39,6 +43,8 @@ describe(commands.BUSINESS_GET, () => {
     sinon.stub(session, 'getId').returns('');
 
     auth.connection.active = true;
+    commandInfo = cli.getCommandInfo(command);
+    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
   });
 
   beforeEach(() => {
@@ -80,6 +86,25 @@ describe(commands.BUSINESS_GET, () => {
     assert.notStrictEqual(command.description, null);
   });
 
+  it('fails validation when id or name are not specified', () => {
+    const actual = commandOptionsSchema.safeParse({});
+    assert.strictEqual(actual.success, false);
+  });
+
+  it('passes validation when id is specified', () => {
+    const actual = commandOptionsSchema.safeParse({
+      id: validId
+    });
+    assert.strictEqual(actual.success, true);
+  });
+
+  it('passes validation when name is specified', () => {
+    const actual = commandOptionsSchema.safeParse({
+      name: validName
+    });
+    assert.strictEqual(actual.success, true);
+  });
+
   it('gets business by id', async () => {
     sinon.stub(request, 'get').callsFake(async (opts) => {
       if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
@@ -89,7 +114,7 @@ describe(commands.BUSINESS_GET, () => {
       throw 'Invalid request';
     });
 
-    await command.action(logger, { options: { id: validId } });
+    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId }) });
     assert(loggerLogSpy.calledWith(businessResponse));
   });
 
@@ -106,7 +131,7 @@ describe(commands.BUSINESS_GET, () => {
       throw 'Invalid request';
     });
 
-    await command.action(logger, { options: { name: validName } });
+    await command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) });
     assert(loggerLogSpy.calledWith(businessResponse));
   });
 
@@ -127,7 +152,7 @@ describe(commands.BUSINESS_GET, () => {
       throw 'Invalid request';
     });
 
-    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError("Multiple businesses with name 'Valid Business' found. Found: mail@contoso.onmicrosoft.com."));
+    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) }), new CommandError("Multiple businesses with name 'Valid Business' found. Found: mail@contoso.onmicrosoft.com."));
   });
 
   it('handles selecting single result when multiple businesses with the specified name found and cli is set to prompt', async () => {
@@ -153,7 +178,7 @@ describe(commands.BUSINESS_GET, () => {
 
     sinon.stub(cli, 'handleMultipleResultsFound').resolves(businessResponse);
 
-    await command.action(logger, { options: { name: validName } });
+    await command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) });
     assert(loggerLogSpy.calledWith(businessResponse));
   });
 
@@ -166,7 +191,7 @@ describe(commands.BUSINESS_GET, () => {
       throw 'Invalid request';
     });
 
-    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
+    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
   });
 
   it('fails when no business found with name because of an empty displayName', async () => {
@@ -178,13 +203,13 @@ describe(commands.BUSINESS_GET, () => {
       throw 'Invalid request';
     });
 
-    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
+    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
   });
 
   it('correctly handles random API error', async () => {
     sinonUtil.restore(request.get);
     sinon.stub(request, 'get').rejects(new Error('An error has occurred'));
-    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
+    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: validName }) } as any), new CommandError('An error has occurred'));
   });
 });

diff --git a/src/m365/booking/commands/business/business-get.ts b/src/m365/booking/commands/business/business-get.ts
index 5a066ea56..bace8b52c 100644
--- a/src/m365/booking/commands/business/business-get.ts
+++ b/src/m365/booking/commands/business/business-get.ts
@@ -1,21 +1,27 @@
 import { BookingBusiness } from '@microsoft/microsoft-graph-types';
+import { z } from 'zod';
+import { cli } from '../../../../cli/cli.js';
 import { Logger } from '../../../../cli/Logger.js';
-import GlobalOptions from '../../../../GlobalOptions.js';
+import { globalOptionsZod } from '../../../../Command.js';
 import request, { CliRequestOptions } from '../../../../request.js';
 import { formatting } from '../../../../utils/formatting.js';
+import { zod } from '../../../../utils/zod.js';
 import GraphCommand from '../../../base/GraphCommand.js';
 import commands from '../../commands.js';
-import { cli } from '../../../../cli/cli.js';
+
+const options = globalOptionsZod
+  .extend({
+    id: zod.alias('i', z.string().uuid().optional()),
+    name: zod.alias('n', z.string().optional())
+  })
+  .strict();
+
+declare type Options = z.infer<typeof options>;
 
 interface CommandArgs {
   options: Options;
 }
 
-interface Options extends GlobalOptions {
-  id?: string;
-  name?: string;
-}
-
 class BookingBusinessGetCommand extends GraphCommand {
   public get name(): string {
     return commands.BUSINESS_GET;
@@ -25,32 +31,15 @@ class BookingBusinessGetCommand extends GraphCommand {
     return 'Retrieve the specified Microsoft Bookings business.';
   }
 
-  constructor() {
-    super();
-
-    this.#initTelemetry();
-    this.#initOptions();
-    this.#initOptionSets();
+  public get schema(): z.ZodTypeAny | undefined {
+    return options;
   }
 
-  #initTelemetry(): void {
-    this.telemetry.push((args: CommandArgs) => {
-      Object.assign(this.telemetryProperties, {
-        id: typeof args.options.id !== 'undefined',
-        name: typeof args.options.name !== 'undefined'
+  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
+    return schema
+      .refine(options => [options.id, options.name].filter(Boolean).length === 1, {
+        message: 'Specify either id or name'
       });
-    });
-  }
-
-  #initOptions(): void {
-    this.options.unshift(
-      { option: '-i, --id [id]' },
-      { option: '-n, --name [name]' }
-    );
-  }
-
-  #initOptionSets(): void {
-    this.optionSets.push({ options: ['id', 'name'] });
   }
 
   public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
```