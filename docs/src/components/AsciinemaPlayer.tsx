import React, { useEffect, useRef, useState } from 'react';
import BrowserOnly from '@docusaurus/BrowserOnly';
import styles from '../scss/AsciinemaPlayer.module.scss';

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
  const [isMounted, setIsMounted] = useState(false);
  const [isPlayerLoaded, setIsPlayerLoaded] = useState(false);

  useEffect(() => {
    setIsMounted(true);
  }, []);

  useEffect(() => {
    if (ref.current && !playerCreated.current && isMounted) {
      const AsciinemaPlayerLibrary = require('asciinema-player');
      AsciinemaPlayerLibrary.create(proxiedSrc, ref.current, asciinemaOptions);
      playerCreated.current = true;
      setIsPlayerLoaded(true);
    }
  }, [proxiedSrc, asciinemaOptions, isMounted]);

  return (
    <BrowserOnly fallback={<div />}>
      {() => <div ref={ref} className={isPlayerLoaded ? '' : styles.hidden}/>}
    </BrowserOnly>
  );
};

export default AsciinemaPlayerComponent;