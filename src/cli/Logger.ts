export interface Logger {
  /**
   * Logs message with formatting to stdout
   */
  log: (message: any) => void;
  /**
   * Logs message without formatting to stdout
   */
  logRaw: (message: any) => void;
  /**
   * Logs message without formatting to stderr
   */
  logToStderr: (message: any) => void;
}