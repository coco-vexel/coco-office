import type { NativeWordChange } from './native-commands.js';
import type { WordEditorChangeEvent } from './runtime.js';

export function createWordEditorChangeEvent(
  previous: readonly NativeWordChange[],
  current: readonly NativeWordChange[],
  revision: number,
): WordEditorChangeEvent {
  let commonPrefix = 0;
  while (
    commonPrefix < previous.length &&
    commonPrefix < current.length &&
    previous[commonPrefix] === current[commonPrefix]
  ) {
    commonPrefix += 1;
  }
  return {
    revision,
    changes: current.slice(commonPrefix),
    revertedChanges: previous.slice(commonPrefix),
    totalChangeCount: current.length,
  };
}
