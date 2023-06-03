import React, { useEffect, useRef } from 'react';
import useIsBrowser from '@docusaurus/useIsBrowser';

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
  if (useIsBrowser()) {
    const ref = useRef<HTMLDivElement>(null);
    const AsciinemaPlayerLibrary = require('asciinema-player');

    useEffect(() => {
      const currentRef = ref.current;
      AsciinemaPlayerLibrary.create(src, currentRef, asciinemaOptions);
    }, [src]);

    return <div ref={ref} />;
  } 
  else {
    return <div/>;
  }
};

export default AsciinemaPlayer;