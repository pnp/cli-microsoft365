import appInsights from "./appInsights.js";
import { Cli } from "./cli/Cli.js";
import { settingsNames } from "./settingsNames.js";

class Telemetry {
  public trackEvent(commandName: string, properties: any): void {
    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, false)) {
      return;
    }

    appInsights.trackEvent({
      name: commandName,
      properties
    });
    appInsights.flush();
  }

  public trackException(exception: any): void {
    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, false)) {
      return;
    }

    appInsights.trackException({
      exception
    });
    appInsights.flush();
  }
}

export const telemetry = new Telemetry();