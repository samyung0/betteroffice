"use client";

import { useAnimate, useInView, useReducedMotion } from "motion/react";
import { useEffect, type ReactNode } from "react";

const EASE = [0.21, 0.47, 0.32, 0.98] as const;

export function FadeIn({
  children,
  delay = 0,
  className,
}: {
  children: ReactNode;
  delay?: number;
  className?: string;
}) {
  const [scope, animate] = useAnimate<HTMLDivElement>();
  const reduce = useReducedMotion();

  useEffect(() => {
    if (reduce) return;
    const animation = animate(
      scope.current,
      { opacity: [0, 1], y: [14, 0] },
      { duration: 0.6, delay, ease: EASE },
    );
    return () => animation.complete();
  }, [animate, delay, reduce, scope]);

  return (
    <div ref={scope} className={className}>
      {children}
    </div>
  );
}

export function Reveal({
  children,
  delay = 0,
  className,
}: {
  children: ReactNode;
  delay?: number;
  className?: string;
}) {
  const [scope, animate] = useAnimate<HTMLDivElement>();
  const reduce = useReducedMotion();
  const inView = useInView(scope, { once: true, margin: "0px 0px -60px 0px" });

  useEffect(() => {
    if (reduce || !inView) return;
    const animation = animate(
      scope.current,
      { opacity: [0, 1], y: [16, 0] },
      { duration: 0.55, delay, ease: EASE },
    );
    return () => animation.complete();
  }, [animate, delay, inView, reduce, scope]);

  return (
    <div ref={scope} className={className}>
      {children}
    </div>
  );
}
