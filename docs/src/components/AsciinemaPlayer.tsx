import React, { useEffect, useRef, useState } from 'react';
import BrowserOnly from '@docusaurus/BrowserOnly';

type AsciinemaPlayerProps = {
  src: string;
  // START asciinemaOptions
  cols: string;
  rows: string;
  autoPlay: boolean;
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
  const proxiedSrc = `https://corsproxy.io/?${encodeURIComponent(src)}`;
  const [isMounted, setIsMounted] = useState(false);

  useEffect(() => {
    setIsMounted(true);
  }, []);

  useEffect(() => {
    const loadAsciinemaPlayer = async () => {
      if (ref.current && isMounted) {
        const AsciinemaPlayerLibrary = await import('asciinema-player');
        AsciinemaPlayerLibrary.create(proxiedSrc, ref.current, asciinemaOptions);
      }
    };

    loadAsciinemaPlayer();
  }, [proxiedSrc, asciinemaOptions, isMounted]);

  return (
    <BrowserOnly fallback={<div />}>
      {() => <div ref={ref} />}
    </BrowserOnly>
  );
};

export default AsciinemaPlayerComponent;