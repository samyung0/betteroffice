import { describe, expect, test } from 'bun:test';
import {
  createDisplayListQueries,
  isDisplayListQuerySourceDead,
  onDisplayListQuerySourceFailure,
} from './displayListQueries';
import type { DisplayPage } from './displayList';
import type { RustDisplayListQueryEngine } from './rustDisplayList';

function page(pageIndex: number): DisplayPage {
  return { pageIndex, width: 100, height: 100, primitives: [] };
}

function fakeEngine() {
  const calls = {
    open: 0,
    update: 0,
    close: 0,
    rangeByHandle: 0,
    rangeJson: 0,
    verticalByHandle: 0,
  };
  let nextHandle = 1;
  const engine: RustDisplayListQueryEngine = {
    hitTestRegionsJson: () => 'null',
    verticalMoveJson: () => 'null',
    rangeRectsJson: () => {
      calls.rangeJson += 1;
      return '[]';
    },
    hasDisplayListSession: () => true,
    openDisplayList: () => {
      calls.open += 1;
      return nextHandle++;
    },
    closeDisplayList: () => {
      calls.close += 1;
    },
    updateDisplayList: () => {
      calls.update += 1;
    },
    hasDisplayListUpdate: () => true,
    rangeRectsByHandle: () => {
      calls.rangeByHandle += 1;
      return '[]';
    },
    verticalMoveByHandle: () => {
      calls.verticalByHandle += 1;
      return '{"position":2,"goalX":24}';
    },
  };
  return { engine, calls };
}

describe('createDisplayListQueries handle lifecycle', () => {
  test('opens the session handle lazily, on the first query', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    expect(calls.open).toBe(0);
    queries.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.rangeByHandle).toBe(1);
  });

  test('prime() acquires the handle without a query', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    queries.prime();
    expect(calls.open).toBe(1);
    expect(calls.rangeByHandle).toBe(0);
    queries.prime();
    expect(calls.open).toBe(1);
  });

  test('adoption chains across unqueried generations as one page-delta', () => {
    const { engine, calls } = fakeEngine();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    const third = createDisplayListQueries({ pages: [shared] }, engine, second);
    expect(calls.open).toBe(1);
    expect(calls.update).toBe(0);
    third.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.update).toBe(1);
  });

  test('superseded generations fall back to JSON-arg queries, never reopening', () => {
    const { engine, calls } = fakeEngine();
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    second.rangeRects(0, 1);
    expect(calls.update).toBe(1);
    first.rangeRects(0, 1);
    expect(calls.open).toBe(1);
    expect(calls.rangeJson).toBe(1);
  });

  test('keeps no list JSON once the handle holds the parsed list', () => {
    const { engine } = fakeEngine();
    const list = { pages: [page(0)] };
    const stringify = JSON.stringify;
    let serialized = 0;
    JSON.stringify = ((value: unknown, ...rest: []) => {
      if (value === list) serialized += 1;
      return stringify(value, ...rest);
    }) as typeof JSON.stringify;
    try {
      let opened = '';
      let fallback = '';
      engine.openDisplayList = (json) => {
        opened = json;
        return 1;
      };
      engine.hitTestRegionsByHandle = () => {
        throw new Error('unknown display-list handle 1');
      };
      engine.hitTestRegionsJson = (json) => {
        fallback = json;
        return 'null';
      };
      const queries = createDisplayListQueries(list, engine);
      queries.prime();
      expect(serialized).toBe(1);
      // The handle fails, so the query falls back to JSON: the list serializes
      // again because the string the handle opened with was not kept.
      queries.hitTestRegions(0, 1, 1);
      expect(fallback).toBe(opened);
      expect(serialized).toBe(2);
    } finally {
      JSON.stringify = stringify;
    }
  });

  test('routes vertical movement through the retained handle', () => {
    const { engine, calls } = fakeEngine();
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);
    expect(queries.verticalMove(1, 'down')).toEqual({ position: 2, goalX: 24 });
    expect(calls.open).toBe(1);
    expect(calls.verticalByHandle).toBe(1);
  });
});

function wasmTrap(): Error {
  const trap = new Error('unreachable executed');
  trap.name = 'RuntimeError';
  return trap;
}

describe('createDisplayListQueries wasm trap containment', () => {
  test('stops querying an instance whose wasm trapped', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
    expect(queries.sourceState().status).toBe('error');

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(queries.verticalMove(1, 'down')).toBeNull();
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
    expect(calls.verticalByHandle).toBe(0);
    expect(isDisplayListQuerySourceDead(engine)).toBe(true);
  });

  test('a returned Err still falls back to the JSON-arg path', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw new Error('unknown display-list handle 7');
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(1);
    expect(isDisplayListQuerySourceDead(engine)).toBe(false);
  });

  test('a later build over the dead instance never queries it', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const shared = page(0);
    const first = createDisplayListQueries({ pages: [shared] }, engine);
    first.rangeRects(0, 1);
    const openAfterTrap = calls.open;

    const second = createDisplayListQueries({ pages: [shared] }, engine, first);
    expect(second.rangeRects(0, 1)).toEqual([]);
    expect(calls.open).toBe(openAfterTrap);
    expect(calls.update).toBe(0);
    expect(calls.rangeByHandle).toBe(1);
    expect(calls.rangeJson).toBe(0);
  });

  test('notifies failure listeners once so the host can rebuild', () => {
    const { engine, calls } = fakeEngine();
    engine.rangeRectsByHandle = () => {
      calls.rangeByHandle += 1;
      throw wasmTrap();
    };
    const failures: Error[] = [];
    const unsubscribe = onDisplayListQuerySourceFailure((error) => failures.push(error));
    const queries = createDisplayListQueries({ pages: [page(0)] }, engine);

    queries.rangeRects(0, 1);
    queries.rangeRects(0, 1);
    unsubscribe();

    expect(failures).toHaveLength(1);
    expect(failures[0].name).toBe('RuntimeError');
  });

  test('a resident trap stops resident queries', () => {
    let residentCalls = 0;
    const resident = {
      displayHitTestRegionsJson: () => 'null',
      displayVerticalMoveJson: () => {
        residentCalls += 1;
        throw wasmTrap();
      },
      displayRangeRectsJson: () => {
        residentCalls += 1;
        return '[]';
      },
      displayRangeRectsRegionJson: () => '[]',
    };
    const queries = createDisplayListQueries({ pages: [page(0)] }, resident);

    expect(queries.verticalMove(1, 'down')).toBeNull();
    expect(residentCalls).toBe(1);
    expect(queries.rangeRects(0, 1)).toEqual([]);
    expect(residentCalls).toBe(1);
    expect(isDisplayListQuerySourceDead(resident)).toBe(true);
  });
});

describe('createDisplayListQueries unbuilt pages (Q815 C)', () => {
  test('a body range or hit test on a page not built yet asks for that page', () => {
    const asked: number[][] = [];
    const unbuilt = (pageIndex: number, span: [number, number]): DisplayPage => ({
      ...page(pageIndex),
      unbuilt: true,
      positionSpan: span,
    });
    const queries = createDisplayListQueries(
      { pages: [page(0), unbuilt(1, [100, 199]), unbuilt(2, [200, 299])] },
      fakeEngine().engine,
      null,
      (pages) => asked.push(pages)
    );
    queries.rangeRects(10, 20);
    expect(asked).toEqual([]);
    queries.caretRect(150);
    expect(asked[0]).toEqual([1]);
    asked.length = 0;
    queries.rangeRects(150, 250);
    expect(asked).toEqual([[1, 2]]);
    queries.hitTestRegions(2, 5, 5);
    queries.hitTestRegions(0, 5, 5);
    expect(asked).toEqual([[1, 2], [2]]);
  });
});
