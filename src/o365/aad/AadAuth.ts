import Auth, { Logger, Service } from "../../Auth";
import { CommandError } from "../../Command";

class AadAuth extends Auth {
  private SERVICE: string = 'AAD';

  public restoreAuth(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .getServiceConnectionInfo<Service>(this.SERVICE)
        .then((service: Service): void => {
          this.service = service;
          resolve();
        }, (error: any): void => {
          resolve();
        });
    });
  }

  public ensureAccessToken(resource: string, stdout: Logger, debug: boolean = false): Promise<string> {
    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const now: number = new Date().getTime() / 1000;
      if (this.service.accessToken && this.service.expiresAt > now) {
        resolve(this.service.accessToken);
        return;
      }

      super
        .ensureAccessToken(this.service.resource, stdout, debug)
        .then((accessToken: string): void => {
          this
            .setServiceConnectionInfo(this.SERVICE, this.service)
            .then((): void => {
              resolve(accessToken);
            }, (error: any): void => {
              if (debug) {
                stdout.log(new CommandError(error));
              }
              resolve(accessToken);
            });
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public storeConnectionInfo(): Promise<void> {
    return this.setServiceConnectionInfo(this.SERVICE, this.service);
  }

  public clearConnectionInfo(): Promise<void> {
    return this.clearServiceConnectionInfo(this.SERVICE);
  }
}

export default new AadAuth(new Service('https://graph.windows.net'), '04b07795-8ddb-461a-bbee-02f9e1bf7b46');