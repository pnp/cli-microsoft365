import { Auth, Service } from "../../Auth";
import config from "../../config";

class GraphAuth extends Auth {
  protected serviceId(): string {
    return 'Graph';
  }
}

export default new GraphAuth(new Service('https://graph.microsoft.com'), config.aadGraphAppId);