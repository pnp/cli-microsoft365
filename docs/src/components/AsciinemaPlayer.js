"use strict";
var __rest = (this && this.__rest) || function (s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
};
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const useIsBrowser_1 = require("@docusaurus/useIsBrowser");
const AsciinemaPlayer = (_a) => {
    var { src } = _a, asciinemaOptions = __rest(_a, ["src"]);
    if ((0, useIsBrowser_1.default)()) {
        const ref = (0, react_1.useRef)(null);
        const AsciinemaPlayerLibrary = require('asciinema-player');
        (0, react_1.useEffect)(() => {
            const currentRef = ref.current;
            AsciinemaPlayerLibrary.create(src, currentRef, asciinemaOptions);
        }, [src]);
        return <div ref={ref}/>;
    }
    else {
        return <div />;
    }
};
exports.default = AsciinemaPlayer;
//# sourceMappingURL=AsciinemaPlayer.js.map