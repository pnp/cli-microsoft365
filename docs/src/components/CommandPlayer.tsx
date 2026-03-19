import React, { useEffect, useRef, useState, useCallback, useMemo } from 'react';

type CommandStep = {
  command: string;
  response?: string;
  responseFrom?: string;
};

type CommandPlayerProps = {
  command?: string;
  response?: string;
  commands?: CommandStep[];
  typingSpeed?: number;
  responseDelay?: number;
  loopDelay?: number;
};

function highlightJson(json: string): React.ReactNode[] {
  const nodes: React.ReactNode[] = [];
  const tokenRegex = /("(?:[^"\\]|\\.)*")\s*(?=:)|("(?:[^"\\]|\\.)*")|(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\b|(true|false)\b|(null)\b|([{}[\]:,])/g;

  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = tokenRegex.exec(json)) !== null) {
    if (match.index > lastIndex) {
      nodes.push(<span key={`ws-${lastIndex}`}>{json.slice(lastIndex, match.index)}</span>);
    }
    const [full, key, str, num, bool, nul, punct] = match;
    if (key !== undefined) nodes.push(<span key={`k-${match.index}`} className="cp-tok-key">{key}</span>);
    else if (str !== undefined) nodes.push(<span key={`s-${match.index}`} className="cp-tok-str">{str}</span>);
    else if (num !== undefined) nodes.push(<span key={`n-${match.index}`} className="cp-tok-num">{num}</span>);
    else if (bool !== undefined) nodes.push(<span key={`b-${match.index}`} className="cp-tok-bool">{bool}</span>);
    else if (nul !== undefined) nodes.push(<span key={`nl-${match.index}`} className="cp-tok-null">{nul}</span>);
    else if (punct !== undefined) nodes.push(<span key={`p-${match.index}`} className="cp-tok-punct">{punct}</span>);
    lastIndex = match.index + full.length;
  }
  if (lastIndex < json.length) {
    nodes.push(<span key={`ws-${lastIndex}`}>{json.slice(lastIndex)}</span>);
  }
  return nodes;
}

function prefersReducedMotion(): boolean {
  return typeof window !== 'undefined'
    && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

function naturalDelay(base: number, char: string, prev: string): number {
  let delay = base * (0.3 + Math.random() * 1.5);

  if (char === ' ') delay += base * (0.3 + Math.random() * 1.2);
  else if (char === '-' && prev === '-') delay += base * (0.1 + Math.random() * 0.4);
  else if (char === '-' && prev === ' ') delay += base * (0.5 + Math.random() * 1.5);
  else if (char === '/' || char === '\\') delay += base * (0.2 + Math.random() * 0.5);
  else if (prev && prev === prev.toUpperCase() && prev !== prev.toLowerCase()) delay += base * 0.15;

  if (Math.random() < 0.10) delay += base * (1.5 + Math.random() * 4);
  if (Math.random() < 0.18) delay *= 0.25;
  if (Math.random() < 0.05) return Math.max(12, base * 0.15);

  return Math.max(delay, 12);
}

function isScrolledToBottom(el: HTMLElement): boolean {
  return el.scrollHeight - el.scrollTop - el.clientHeight < 5;
}

type Phase = 'idle' | 'typing-command' | 'showing-response' | 'typing-clear' | 'clearing';

const BETWEEN_STEP_MIN = 1500;
const BETWEEN_STEP_JITTER = 2500;
const SCROLL_DETECT_TIMEOUT = 200;
const RING_RADIUS = 9;
const CLEAR_EXECUTE_DELAY = 300;
const CLEAR_PAUSE_BEFORE_EXECUTE = 200;

function betweenStepDelay(): number {
  return BETWEEN_STEP_MIN + Math.random() * BETWEEN_STEP_JITTER;
}

function randomResponseDelay(base: number): number {
  if (Math.random() < 0.2) return base * (0.15 + Math.random() * 0.25);
  if (Math.random() < 0.15) return base * (1.5 + Math.random() * 1.5);
  return base * (0.5 + Math.random() * 1.0);
}

function initialPause(isFirstStep: boolean): number {
  if (isFirstStep) return 300 + Math.random() * 500;
  return 200 + Math.random() * 400;
}

const AUTO_RESUME_DELAY = 30_000;

const PauseIcon: React.FC = () => (
  <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
    <rect x="2" y="1" width="3" height="10" rx="0.5" />
    <rect x="7" y="1" width="3" height="10" rx="0.5" />
  </svg>
);

const PlayIcon: React.FC = () => (
  <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
    <path d="M3 1.5v9l7-4.5z" />
  </svg>
);

const CommandPlayer: React.FC<CommandPlayerProps> = ({
  command,
  response,
  commands,
  typingSpeed = 45,
  responseDelay = 500,
  loopDelay = 4
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const bodyRef = useRef<HTMLDivElement>(null);
  const [mounted, setMounted] = useState(false);
  const cancelRef = useRef({ cancelled: false });

  const pausedRef = useRef(false);
  const [isPaused, setIsPaused] = useState(false);
  const manualPauseRef = useRef(false);
  const scrollPauseRef = useRef(false);
  const autoResumeTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const ringAnimFrameRef = useRef<number | null>(null);
  const ringRef = useRef<SVGCircleElement>(null);
  const ringStartTimeRef = useRef<number>(0);
  const [scrollPauseActive, setScrollPauseActive] = useState(false);

  const steps = useMemo<CommandStep[]>(() => {
    if (commands && commands.length > 0) return commands;
    if (command) return [{ command, response }];
    return [];
  }, [commands, command, response]);

  const formattedResponses = useMemo(() => {
    return steps.map(step => {
      if (!step.response) return '';
      try { return JSON.stringify(JSON.parse(step.response), null, 2); }
      catch { return step.response.replace(/\\n/g, '\n'); }
    });
  }, [steps]);

  const [displayedSteps, setDisplayedSteps] = useState(0);
  const [typedCommand, setTypedCommand] = useState('');
  const [showCurrentResponse, setShowCurrentResponse] = useState(false);
  const [typedClear, setTypedClear] = useState('');
  const [showClearPrompt, setShowClearPrompt] = useState(false);
  const [phase, setPhase] = useState<Phase>('idle');

  useEffect(() => { setMounted(true); }, []);

  const ringCircumference = 2 * Math.PI * RING_RADIUS;

  const clearAutoResumeTimer = useCallback(() => {
    if (autoResumeTimerRef.current) {
      clearTimeout(autoResumeTimerRef.current);
      autoResumeTimerRef.current = null;
    }
    if (ringAnimFrameRef.current) {
      cancelAnimationFrame(ringAnimFrameRef.current);
      ringAnimFrameRef.current = null;
    }
    setScrollPauseActive(false);
  }, []);

  useEffect(() => {
    if (!scrollPauseActive) return;
    const startTime = ringStartTimeRef.current;
    const animate = (): void => {
      const elapsed = Date.now() - startTime;
      const progress = Math.min(1, elapsed / AUTO_RESUME_DELAY);
      if (ringRef.current) {
        ringRef.current.style.strokeDashoffset = String(ringCircumference * (1 - progress));
      }
      if (progress < 1 && pausedRef.current) {
        ringAnimFrameRef.current = requestAnimationFrame(animate);
      }
    };
    ringAnimFrameRef.current = requestAnimationFrame(animate);
    return () => {
      if (ringAnimFrameRef.current) {
        cancelAnimationFrame(ringAnimFrameRef.current);
        ringAnimFrameRef.current = null;
      }
    };
  }, [scrollPauseActive, ringCircumference]);

  const doPause = useCallback((source: 'manual' | 'scroll') => {
    if (source === 'manual') manualPauseRef.current = true;
    if (source === 'scroll') scrollPauseRef.current = true;
    pausedRef.current = true;
    setIsPaused(true);
    clearAutoResumeTimer();

    if (source === 'scroll') {
      ringStartTimeRef.current = Date.now();
      setScrollPauseActive(true);
      autoResumeTimerRef.current = setTimeout(() => {
        scrollPauseRef.current = false;
        pausedRef.current = false;
        setIsPaused(false);
        setScrollPauseActive(false);
      }, AUTO_RESUME_DELAY);
    }
  }, [clearAutoResumeTimer]);

  const doResume = useCallback(() => {
    manualPauseRef.current = false;
    scrollPauseRef.current = false;
    pausedRef.current = false;
    setIsPaused(false);
    clearAutoResumeTimer();
  }, [clearAutoResumeTimer]);

  const togglePause = useCallback(() => {
    if (pausedRef.current) {
      doResume();
    } else {
      doPause('manual');
    }
  }, [doPause, doResume]);

  useEffect(() => {
    const el = bodyRef.current;
    if (!el || !mounted) return;

    let userScrolling = false;
    let scrollTimeout: ReturnType<typeof setTimeout> | null = null;

    const handleScroll = (): void => {
      if (!userScrolling) return;

      if (isScrolledToBottom(el)) {
        if (scrollPauseRef.current && !manualPauseRef.current) {
          doResume();
        }
      } else {
        if (!pausedRef.current) {
          doPause('scroll');
        }
      }
    };

    const markUserScrolling = (): void => {
      userScrolling = true;
      if (scrollTimeout) clearTimeout(scrollTimeout);
      scrollTimeout = setTimeout(() => { userScrolling = false; }, SCROLL_DETECT_TIMEOUT);
    };

    const handleKeyDown = (e: KeyboardEvent): void => {
      if (['ArrowUp', 'ArrowDown', 'PageUp', 'PageDown', 'Home', 'End', ' '].includes(e.key)) {
        markUserScrolling();
      }
    };

    el.addEventListener('scroll', handleScroll, { passive: true });
    el.addEventListener('wheel', markUserScrolling, { passive: true });
    el.addEventListener('touchmove', markUserScrolling, { passive: true });
    el.addEventListener('pointerdown', markUserScrolling, { passive: true });
    el.addEventListener('keydown', handleKeyDown);

    return () => {
      el.removeEventListener('scroll', handleScroll);
      el.removeEventListener('wheel', markUserScrolling);
      el.removeEventListener('touchmove', markUserScrolling);
      el.removeEventListener('pointerdown', markUserScrolling);
      el.removeEventListener('keydown', handleKeyDown);
      if (scrollTimeout) clearTimeout(scrollTimeout);
    };
  }, [mounted, doPause, doResume]);

  useEffect(() => {
    const el = bodyRef.current;
    if (el && !pausedRef.current) el.scrollTop = el.scrollHeight;
  }, [displayedSteps, typedCommand, showCurrentResponse, typedClear, showClearPrompt]);

  const pauseAwareTimeout = useCallback((fn: () => void, ms: number): void => {
    const token = cancelRef.current;
    let remaining = ms;
    let start = Date.now();

    const check = (): void => {
      if (token.cancelled) return;

      if (pausedRef.current) {
        setTimeout(check, 50);
        start = Date.now();
        return;
      }

      remaining -= (Date.now() - start);
      if (remaining <= 0) {
        fn();
      } else {
        start = Date.now();
        setTimeout(check, Math.min(remaining, 50));
      }
    };

    setTimeout(check, Math.min(ms, 50));
  }, []);

  const runAnimation = useCallback(() => {
    const token = cancelRef.current;

    if (prefersReducedMotion()) {
      setDisplayedSteps(steps.length);
      setTypedCommand('');
      setShowCurrentResponse(false);
      setShowClearPrompt(false);
      setTypedClear('');
      setPhase('showing-response');
      return;
    }

    setDisplayedSteps(0);
    setTypedCommand('');
    setShowCurrentResponse(false);
    setShowClearPrompt(false);
    setTypedClear('');

    const startClearTyping = (): void => {
      if (token.cancelled) return;
      setPhase('typing-clear');

      let clearIdx = 0;
      const clearText = 'clear';

      const typeClear = (): void => {
        if (token.cancelled) return;
        if (clearIdx < clearText.length) {
          clearIdx++;
          setTypedClear(clearText.slice(0, clearIdx));
          pauseAwareTimeout(typeClear, naturalDelay(typingSpeed, clearText[clearIdx - 1], clearIdx > 1 ? clearText[clearIdx - 2] : ''));
        } else {
          pauseAwareTimeout(() => {
            if (token.cancelled) return;
            setPhase('clearing');
            pauseAwareTimeout(() => {
              if (token.cancelled) return;
              runAnimation();
            }, CLEAR_EXECUTE_DELAY);
          }, CLEAR_PAUSE_BEFORE_EXECUTE);
        }
      };

      pauseAwareTimeout(() => { if (!token.cancelled) typeClear(); }, initialPause(false));
    };

    const animateStep = (stepIdx: number, skipInitialPause = false): void => {
      if (token.cancelled) return;

      if (stepIdx >= steps.length) {
        startClearTyping();
        return;
      }

      setTypedCommand('');
      setShowCurrentResponse(false);
      setPhase('typing-command');

      let charIdx = 0;
      const cmd = steps[stepIdx].command;

      const typeCommand = (): void => {
        if (token.cancelled) return;
        if (charIdx < cmd.length) {
          charIdx++;
          setTypedCommand(cmd.slice(0, charIdx));
          pauseAwareTimeout(typeCommand, naturalDelay(typingSpeed, cmd[charIdx - 1], charIdx > 1 ? cmd[charIdx - 2] : ''));
        } else {
          pauseAwareTimeout(() => {
            if (token.cancelled) return;
            setShowCurrentResponse(true);
            setPhase('showing-response');

            const isLast = stepIdx === steps.length - 1;

            if (isLast) {
              setDisplayedSteps(stepIdx + 1);
              setShowClearPrompt(true);
              pauseAwareTimeout(() => {
                if (token.cancelled) return;
                animateStep(stepIdx + 1);
              }, loopDelay * 1000);
            } else {
              setDisplayedSteps(stepIdx + 1);
              setTypedCommand('');
              setShowCurrentResponse(false);
              setPhase('typing-command');
              pauseAwareTimeout(() => {
                if (token.cancelled) return;
                animateStep(stepIdx + 1, true);
              }, betweenStepDelay());
            }
          }, randomResponseDelay(responseDelay));
        }
      };

      if (skipInitialPause) {
        typeCommand();
      } else {
        pauseAwareTimeout(() => {
          if (!token.cancelled) typeCommand();
        }, initialPause(stepIdx === 0));
      }
    };

    animateStep(0);
  }, [steps, formattedResponses, typingSpeed, responseDelay, loopDelay, pauseAwareTimeout]);

  useEffect(() => {
    return () => {
      cancelRef.current.cancelled = true;
      clearAutoResumeTimer();
    };
  }, [clearAutoResumeTimer]);

  useEffect(() => {
    if (!mounted || phase !== 'idle') return;
    const el = containerRef.current;
    if (!el) return;

    if (prefersReducedMotion()) { runAnimation(); return; }

    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            observer.disconnect();
            runAnimation();
            break;
          }
        }
      },
      { threshold: 0, rootMargin: '0px 0px -30px 0px' }
    );

    observer.observe(el);
    return () => { observer.disconnect(); };
  }, [mounted, phase, runAnimation]);

  const currentStepIdx = displayedSteps;

  const highlightedResponses = useMemo(
    () => formattedResponses.map((response) => (response ? highlightJson(response) : null)),
    [formattedResponses]
  );

  return (
    <div className="cp-container" ref={containerRef}>
      <div className="cp-titlebar">
        <span className="cp-titlebar-icon">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
            <path d="M2 3.5L7.5 8 2 12.5V3.5zM8 11.5h6v1H8v-1z" />
          </svg>
        </span>
        <span className="cp-titlebar-text">Terminal</span>
      </div>

      <div className="cp-body" ref={bodyRef}>
        {steps.slice(0, mounted ? displayedSteps : steps.length).map((step, i) => (
          <div key={i}>
            <div className={`cp-prompt${i > 0 ? ' cp-prompt--next' : ''}`}>
              <span className="cp-prompt-symbol">&gt;</span>
              <span className="cp-command">{step.command}</span>
            </div>
            {highlightedResponses[i] && (
              <div className="cp-response cp-response--visible">
                {highlightedResponses[i]}
              </div>
            )}
          </div>
        ))}

        {mounted && currentStepIdx < steps.length && (
          <div>
            <div className={`cp-prompt${currentStepIdx > 0 ? ' cp-prompt--next' : ''}`}>
              <span className="cp-prompt-symbol">&gt;</span>
              <span className="cp-command">
                {typedCommand}
                {phase === 'typing-command' && <span className="cp-cursor" />}
              </span>
            </div>
            {showCurrentResponse && highlightedResponses[currentStepIdx] && (
              <div className="cp-response cp-response--visible">
                {highlightedResponses[currentStepIdx]}
              </div>
            )}
          </div>
        )}

        {showClearPrompt && (
          <div className="cp-prompt cp-prompt--clear">
            <span className="cp-prompt-symbol">&gt;</span>
            <span className="cp-command">
              {typedClear}
              {phase === 'typing-clear' && <span className="cp-cursor" />}
            </span>
          </div>
        )}
      </div>

      {mounted && !prefersReducedMotion() && phase !== 'idle' && (
        <button
          className={`cp-pause-btn${isPaused ? ' cp-pause-btn--paused' : ''}`}
          onClick={togglePause}
          aria-label={isPaused ? 'Resume animation' : 'Pause animation'}
          title={isPaused ? 'Resume' : 'Pause'}
        >
          {isPaused && !manualPauseRef.current && scrollPauseActive && (
            <svg className="cp-pause-ring" width="28" height="28" viewBox="0 0 28 28">
              <circle
                ref={ringRef}
                cx="14" cy="14" r={RING_RADIUS}
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeDasharray={ringCircumference}
                strokeDashoffset={ringCircumference}
                strokeLinecap="round"
                transform="rotate(-90 14 14)"
                opacity="0.4"
              />
            </svg>
          )}
          {isPaused ? <PlayIcon /> : <PauseIcon />}
        </button>
      )}
    </div>
  );
};

export default CommandPlayer;
