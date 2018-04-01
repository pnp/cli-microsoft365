import { Auth, Service } from "../../Auth";
import config from "../../config";

class AzmgmtAuth extends Auth {
  protected serviceId(): string {
    return 'AzMgmt';
  }
}

export default new AzmgmtAuth(new Service('https://management.azure.com/'), config.aadAzmgmtAppId);