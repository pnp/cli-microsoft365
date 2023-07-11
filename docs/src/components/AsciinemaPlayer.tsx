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

const AsciinemaPlayer: React.FC<AsciinemaPlayerProps> = ({
  src,
  ...asciinemaOptions
}) => {
  const ref = useRef<HTMLDivElement>(null);
  const playerCreated = useRef(false);

  useEffect(() => {
    if (ref.current && !playerCreated.current) {
      const AsciinemaPlayerLibrary = require('asciinema-player');
      AsciinemaPlayerLibrary.create(src, ref.current, asciinemaOptions);
      playerCreated.current = true;
    }
  }, [src, asciinemaOptions]);

  return (
    <BrowserOnly fallback={<div />}>
      {() => <div ref={ref} />}
    </BrowserOnly>
  );
};

export default AsciinemaPlayer;