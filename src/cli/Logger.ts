export interface Logger {
  /**
   * Logs message with formatting to stdout
   */
  log: (message: any) => Promise<void>;
  /**
   * Logs message without formatting to stdout
   */
  logRaw?: (message: any) => Promise<void>;
  /**
   * Logs message without formatting to stderr
   */
  logToStderr: (message: any) => Promise<void>;
}