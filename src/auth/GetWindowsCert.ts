import * as ca from "win-ca";
import * as caApi from "win-ca/api";
import * as crypto from "crypto";

const useWindowsCertsThumbprint = (id: string) => {
  const thumbprint = (cert: string) => {
    var shasum = crypto.createHash("sha1");
    shasum.update(Buffer.from(cert, "base64"));
    return shasum.digest("hex").toUpperCase();
  };

  const list: string[] = [];

  caApi({
    store: ["My"],
    ondata: list,
  });

  list.forEach((cert) => {
    const certThumbprint = thumbprint(cert);
    if (certThumbprint === id) {
      const toPEM = ca.der2(ca.der2.pem);
      const pem = toPEM(cert);
      return pem;
    }
  });
};

export default useWindowsCertsThumbprint;
