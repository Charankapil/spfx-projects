import { useEffect, useState } from 'react';

export function prefersReducedMotion(): boolean {
  return typeof window !== 'undefined' && !!window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

/**
 * True for a moment after `key` changes, then false. Entrance animations are
 * applied only while this is true, so once they have played nothing (a
 * relayout, a style recalculation) can replay them and make tiles blink.
 */
export function useEntrance(key: unknown, durationMs = 1200): boolean {
  const [entering, setEntering] = useState(!prefersReducedMotion());
  useEffect(() => {
    if (prefersReducedMotion()) {
      setEntering(false);
      return undefined;
    }
    setEntering(true);
    const timer = window.setTimeout(() => setEntering(false), durationMs);
    return () => window.clearTimeout(timer);
  }, [key, durationMs]);
  return entering;
}

/** Animates from 0 to `target` once per target change; jumps straight there when motion is reduced. */
export function useCountUp(target: number, durationMs = 700): number {
  const [value, setValue] = useState(prefersReducedMotion() ? target : 0);

  useEffect(() => {
    if (prefersReducedMotion() || target <= 0) {
      setValue(target);
      return undefined;
    }
    let frame = 0;
    const start = Date.now();
    const tick = (): void => {
      const t = Math.min(1, (Date.now() - start) / durationMs);
      const eased = 1 - Math.pow(1 - t, 3);
      setValue(Math.round(target * eased));
      if (t < 1) {
        frame = window.requestAnimationFrame(tick);
      }
    };
    frame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(frame);
  }, [target, durationMs]);

  return value;
}
