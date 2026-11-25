import { z } from 'zod';

const propertiesSchema = z.object({
  message: z.string().describe('A message to display.')
});

export type IHelloWorldCopilotComponentProperties = z.infer<typeof propertiesSchema>;

export default z.toJSONSchema(propertiesSchema, { target: 'draft-7' });
