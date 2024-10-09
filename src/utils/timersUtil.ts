import { setTimeout } from "timers/promises";

export const timersUtil = {
  /**
   * Timeout for a specific duration.
   * @param duration Duration in milliseconds.
   */
  /* c8 ignore next 4 */
  // Function is created so we can easily mock it in our tests
  async setTimeout(duration: number): Promise<void> {
    return setTimeout(duration);
  }
};