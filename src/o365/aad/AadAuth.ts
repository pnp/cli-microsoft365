import { Auth, Service } from "../../Auth";
import config from "../../config";

class AadAuth extends Auth {
  protected serviceId(): string {
    return 'AAD';
  }
}

export default new AadAuth(new Service('https://graph.windows.net'), config.aadAadAppId);