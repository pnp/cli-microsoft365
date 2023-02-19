import React, { useEffect, useRef } from 'react';
import BrowserOnly from '@docusaurus/BrowserOnly';
const AsciinemaPlayer = ({ src, ...asciinemaOptions }) => {
    const ref = useRef(null);
    useEffect(() => {
        const currentRef = ref.current;
        if (currentRef) {
            const AsciinemaPlayerLibrary = require('asciinema-player');
            AsciinemaPlayerLibrary.create(src, currentRef, asciinemaOptions);
        }
    }, [src, ref.current]);
    return (<BrowserOnly fallback={<div />}>
      {() => <div ref={ref}/>}
    </BrowserOnly>);
};
export default AsciinemaPlayer;
//# sourceMappingURL=AsciinemaPlayer.js.map