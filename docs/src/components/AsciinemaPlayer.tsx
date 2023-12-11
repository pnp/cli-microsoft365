import React, { useEffect, useRef } from 'react';
import BrowserOnly from '@docusaurus/BrowserOnly';

type AsciinemaPlayerProps = {
  src: string;
  // START asciinemaOptions
  cols: string;
  rows: string;
  autoPlay: boolean
  preload: boolean;
  loop: boolean | number;
  startAt: number | string;
  speed: number;
  idleTimeLimit: number;
  theme: string;
  poster: string;
  fit: string;
  fontSize: string;
  // END asciinemaOptions
};

const AsciinemaPlayerComponent: React.FC<AsciinemaPlayerProps> = ({
  src,
  ...asciinemaOptions
}) => {
  const ref = useRef<HTMLDivElement>(null);
  const playerCreated = useRef(false);
  const proxiedSrc = `https://corsproxy.io/?${encodeURIComponent(src)}`;

  useEffect(() => {
    if (ref.current && !playerCreated.current) {
      const AsciinemaPlayerLibrary = require('asciinema-player');
      AsciinemaPlayerLibrary.create(proxiedSrc, ref.current, asciinemaOptions);
      playerCreated.current = true;
    }
  }, [proxiedSrc, asciinemaOptions]);

  return (
    <BrowserOnly fallback={<div/>}>
      {() => <div ref={ref} />}
    </BrowserOnly>
  );
};

export default AsciinemaPlayerComponent;