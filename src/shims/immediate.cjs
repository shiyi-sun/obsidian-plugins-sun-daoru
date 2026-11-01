"use strict";

const queue = [];
let scheduled = false;

function drain() {
  scheduled = false;
  const tasks = queue.splice(0);
  for (const task of tasks) task();
}

module.exports = function immediate(task) {
  queue.push(task);
  if (scheduled) return;
  scheduled = true;
  setTimeout(drain, 0);
};
