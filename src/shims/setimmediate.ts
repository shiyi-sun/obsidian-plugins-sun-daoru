type ImmediateCallback = (...args: unknown[]) => void;

const globalScope = globalThis as unknown as {
  setImmediate?: (callback: ImmediateCallback, ...args: unknown[]) => number;
  clearImmediate?: (handle: number) => void;
};

if (typeof globalScope.setImmediate !== "function") {
  let nextHandle = 1;
  const pending = new Map<number, ImmediateCallback>();

  globalScope.setImmediate = (callback, ...args) => {
    const handle = nextHandle++;
    pending.set(handle, () => callback(...args));
    setTimeout(() => {
      const task = pending.get(handle);
      if (!task) return;
      pending.delete(handle);
      task();
    }, 0);
    return handle;
  };

  globalScope.clearImmediate = (handle) => {
    pending.delete(handle);
  };
}
