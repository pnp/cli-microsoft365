import * as crypto from 'crypto';
import { cache } from "./cache";

export const session = {
  getId(pid: number): string {
    const key = `${pid.toString()}_session`;
    let sessionId: string | undefined = cache.getValue(key);
    if (sessionId) {
      return sessionId;
    }

    sessionId = crypto.randomBytes(24).toString('base64');
    cache.setValue(key, sessionId);
    return sessionId;
  }
};