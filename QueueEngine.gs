// Pure Queue Engine Core
// Process generic dispositions: 'ELIGIBLE', 'SKIP_ONCE', 'EXCLUDED', 'COMPLETE_FOR_PHASE', 'PERMANENT_PASS'

/**
 * Calculates the next pure queue state.
 * @param {object} currentState The current parsed Config state.
 * @param {Array<object>} roster The list of ordered participant objects: { participantId, disposition }
 * @param {object} config Action configuration: { action: 'INIT' | 'READ' | 'COMPLETE' | 'RECONCILE', windowSize: positive_int, phase: str, orderSource: str, movementMode: 'FORWARD_ONLY' | 'SERPENTINE', completedTurnId: str, preserveSkips?: boolean }
 * @returns {object} { nextState: object, queueComplete: boolean, completionReason: string }
 */
function calculateNextQueueState_(currentState, roster, config) {
  if (!config.windowSize || config.windowSize < 1 || !Number.isInteger(config.windowSize)) {
    throw new Error('Invalid window size: must be a positive integer');
  }

  if (['INIT', 'READ', 'COMPLETE', 'RECONCILE'].indexOf(config.action) === -1) {
    throw new Error('Invalid action');
  }

  if (config.action === 'READ') {
      return { nextState: currentState, queueComplete: false, completionReason: "" };
  }

  const orderSource = String(config.orderSource || currentState["Current Queue Order Source"] || '').trim().toLowerCase();
  const orderField = orderSource === 'seniority' ? 'seniority' : (orderSource === 'lottery' ? 'lottery' : null);
  let orderedRoster = roster.slice();
  if (orderField) {
      const seenPositions = {};
      orderedRoster.forEach(participant => {
          if (participant.disposition === 'EXCLUDED' || participant.disposition === 'PERMANENT_PASS') return;
          const position = Number(participant[orderField]);
          if (!Number.isInteger(position) || position <= 0) {
              throw new Error(`Invalid ${orderSource} position for participant ${participant.participantId}. Positions must be positive whole numbers.`);
          }
          if (seenPositions[position]) {
              throw new Error(`Duplicate ${orderSource} position ${position} for participants ${seenPositions[position]} and ${participant.participantId}.`);
          }
          seenPositions[position] = participant.participantId;
      });
      orderedRoster = orderedRoster.map((participant, index) => ({ participant, index })).sort((aEntry, bEntry) => {
          const a = aEntry.participant;
          const b = bEntry.participant;
          const aExcluded = a.disposition === 'EXCLUDED' || a.disposition === 'PERMANENT_PASS';
          const bExcluded = b.disposition === 'EXCLUDED' || b.disposition === 'PERMANENT_PASS';
          if (aExcluded !== bExcluded) return aExcluded ? 1 : -1;
          if (!aExcluded) return Number(a[orderField]) - Number(b[orderField]);
          return aEntry.index - bEntry.index;
      }).map(entry => entry.participant);
  }

  // Clone current state for safe mutation
  const nextState = {
    ...currentState,
    "Current Queue Skip State": { ...currentState["Current Queue Skip State"] },
    "Current Active Window": [...currentState["Current Active Window"]],
    "Current Directional Window": [...currentState["Current Directional Window"]],
    "Current Directional Window Completed": [...currentState["Current Directional Window Completed"]]
  };

  let queueComplete = false;
  let completionReason = "";
  let membershipChanged = false;
  let completedAppearance = null;

  if (config.action === 'INIT') {
    nextState["Current Queue Phase"] = config.phase;
    nextState["Current Queue Order Source"] = config.orderSource;
    nextState["Current Queue Cycle"] = 1;
    nextState["Current Serpentine Direction"] = "FORWARD";
    nextState["Current Queue Cursor"] = 0;
    nextState["Current Active Window"] = [];
    nextState["Current Directional Window"] = [];
    nextState["Current Directional Window Completed"] = [];
    nextState["Current Queue Skip State"] = config.preserveSkips ? { ...currentState["Current Queue Skip State"] } : {};
    membershipChanged = true;
  } else if (config.action === 'COMPLETE') {
    const targetTurnId = config.completedTurnId;
    const activeIdx = nextState["Current Active Window"].findIndex(t => t.turnId === targetTurnId);

    if (activeIdx === -1) {
        // Stale or invalid completion, just return current state without changes
        return { nextState: currentState, queueComplete: false, completionReason: "" };
    }
    completedAppearance = nextState["Current Active Window"][activeIdx];

    // Check if it belongs to frozen directional window
    if (nextState["Current Directional Window"].length > 0) {
        const isInFrozen = nextState["Current Directional Window"].some(t => t.turnId === targetTurnId);
        if (isInFrozen && !nextState["Current Directional Window Completed"].includes(targetTurnId)) {
            nextState["Current Directional Window Completed"].push(targetTurnId);
        }
    }

    nextState["Current Active Window"].splice(activeIdx, 1);
    membershipChanged = true;
  } else if (config.action === 'RECONCILE') {
    // RECONCILE just triggers the window replenishment block logic without forcing a removal.
    // If window size increased, we will fetch new people below.
    // If window size decreased, we don't revoke existing turns.
  }

  // Helper to safely advance cursor
  const advanceCursor = () => {
      nextState["Current Queue Cursor"] += (nextState["Current Serpentine Direction"] === "FORWARD" ? 1 : -1);
  };

  // Loop to refill Active Window and handle reversals
  // Note: We only loop if we are NOT waiting at an endpoint barrier.
  let loops = 0; // prevent infinite loop if all skips
  const maxLoops = orderedRoster.length * 2 + 5;

  while (loops < maxLoops) {
      loops++;

      // 1. Check endpoint reversal state
      if (nextState["Current Directional Window"].length > 0) {
          const allFrozenCompleted = nextState["Current Directional Window"].every(t =>
              nextState["Current Directional Window Completed"].includes(t.turnId)
          );

          if (allFrozenCompleted) {
              // Reversal happens
              nextState["Current Directional Window"] = [];
              nextState["Current Directional Window Completed"] = [];
              nextState["Current Serpentine Direction"] = nextState["Current Serpentine Direction"] === "FORWARD" ? "BACKWARD" : "FORWARD";
              nextState["Current Queue Cycle"]++;
              nextState["Current Queue Cursor"] = nextState["Current Serpentine Direction"] === "FORWARD" ? 0 : orderedRoster.length - 1;
              membershipChanged = true; // New cycle, bump generation
          } else {
              // Barrier is still active, don't admit anyone new from the reverse direction
              break;
          }
      }

      // 2. Check if we need to refill window
      if (nextState["Current Active Window"].length >= config.windowSize) {
          break; // Window full
      }

      // 3. Check if we reached the end of the roster
      if (nextState["Current Queue Cursor"] < 0 || nextState["Current Queue Cursor"] >= orderedRoster.length) {
          if (config.movementMode === 'FORWARD_ONLY') {
              if (nextState["Current Active Window"].length === 0) {
                  queueComplete = true;
                  completionReason = (orderedRoster.length === 0 || orderedRoster.every(r => r.disposition !== 'ELIGIBLE' && r.disposition !== 'SKIP_ONCE')) ? 'EMPTY_ROSTER' : 'FORWARD_PASS_COMPLETE';
              }
              break;
          } else if (config.movementMode === 'SERPENTINE') {
              // Wait for all currently active (which belong to the final window) to finish before reversing
              if (nextState["Current Active Window"].length > 0) {
                  // Freeze the window
                  nextState["Current Directional Window"] = [...nextState["Current Active Window"]];
                  nextState["Current Directional Window Completed"] = [];
                  break; // Stop adding until they finish
              } else {
                  // If everyone finished (active window is 0), flip immediately
                  nextState["Current Serpentine Direction"] = nextState["Current Serpentine Direction"] === "FORWARD" ? "BACKWARD" : "FORWARD";
                  nextState["Current Queue Cycle"]++;
                  nextState["Current Queue Cursor"] = nextState["Current Serpentine Direction"] === "FORWARD" ? 0 : orderedRoster.length - 1;
                  membershipChanged = true;
                  // If roster is completely empty, queue is complete
                  if (orderedRoster.length === 0) {
                      queueComplete = true;
                      completionReason = "EMPTY_ROSTER";
                      break;
                  }
                  continue; // loop again to evaluate reverse direction
              }
          }
      }

      // 4. Try to admit the participant at cursor
      if (orderedRoster.length === 0) break; // safeguard

      const pIndex = nextState["Current Queue Cursor"];
      const p = orderedRoster[pIndex];
      const pDisposition = p.disposition;

      if (pDisposition === 'ELIGIBLE' || pDisposition === 'SKIP_ONCE') {
          // Check if already in active window (e.g., to prevent duplicate across directions or general bugs)
          if (nextState["Current Active Window"].some(a => a.participantId === p.participantId)) {
              // They are already in the window (e.g., duplicate).
              // Wait, a single participant should only be added once.
              // So if they are already in the window, we just advance cursor?
              // Usually they shouldn't be reached again until a reversal.
              advanceCursor();
              continue;
          }

          // Pending skips are durable state. Consume one only when this participant reaches
          // an otherwise-admissible appearance and is not already concurrently ACTIVE.
          let skipCounts = Number(nextState["Current Queue Skip State"][p.participantId] || 0);
          if (skipCounts > 0) {
              const sameAppearance = completedAppearance &&
                  completedAppearance.participantId === p.participantId &&
                  Number(completedAppearance.cycle) === Number(nextState["Current Queue Cycle"]) &&
                  String(completedAppearance.direction) === String(nextState["Current Serpentine Direction"]);
              if (sameAppearance) {
                  advanceCursor();
                  continue;
              }
              skipCounts--;
              if (skipCounts === 0) delete nextState["Current Queue Skip State"][p.participantId];
              else nextState["Current Queue Skip State"][p.participantId] = skipCounts;
              advanceCursor();
              continue;
          }

          // Issue a Turn
          const turnId = `${nextState["Active Year"]}|${nextState["Current Queue Phase"]}|${nextState["Current Queue Cycle"]}|${nextState["Current Serpentine Direction"]}|${p.participantId}`;

          nextState["Current Active Window"].push({
              turnId: turnId,
              participantId: p.participantId,
              cycle: nextState["Current Queue Cycle"],
              direction: nextState["Current Serpentine Direction"],
              position: pIndex
          });
          membershipChanged = true;
          advanceCursor();
      } else {
          // EXCLUDED, COMPLETE_FOR_PHASE, PERMANENT_PASS
          advanceCursor();
      }
  }

  // 5. Final completion checks
  if (!queueComplete && nextState["Current Active Window"].length === 0) {
      if (config.movementMode === 'SERPENTINE' && orderedRoster.length > 0 && orderedRoster.every(r => r.disposition !== 'ELIGIBLE' && r.disposition !== 'SKIP_ONCE')) {
           queueComplete = true;
           completionReason = "NO_ELIGIBLE_PARTICIPANTS";
      } else if (orderedRoster.length === 0) {
           queueComplete = true;
           completionReason = "EMPTY_ROSTER";
      }
  }

  if (membershipChanged) {
      nextState["Active Window Generation"]++;
  }

  return { nextState, queueComplete, completionReason };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    calculateNextQueueState: calculateNextQueueState_
  };
}

// --- Durable Mutation Wrapper ---

/**
 * Normalizes roster reading from Turn Management sheet.
 */
function processReconcile_(config, rosterCallback) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
        throw new Error('Could not obtain lock for queue reconciliation');
    }

    try {
        const _readConfig = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
        const _writeConfig = typeof writeConfigState_ === 'function' ? writeConfigState_ : (typeof global !== 'undefined' && global.writeConfigState_ ? global.writeConfigState_ : (require('./State.gs').writeConfigState));

        const currentState = _readConfig();

        const roster = rosterCallback();

        // Parse config fields
        const parsedState = {
            ...currentState,
            "Current Active Window": JSON.parse(currentState["Current Active Window"] || "[]"),
            "Current Directional Window": JSON.parse(currentState["Current Directional Window"] || "[]"),
            "Current Directional Window Completed": JSON.parse(currentState["Current Directional Window Completed"] || "[]"),
            "Current Queue Skip State": JSON.parse(currentState["Current Queue Skip State"] || "{}"),
            "Current Queue Cursor": Number(currentState["Current Queue Cursor"] || 0),
            "Current Queue Cycle": Number(currentState["Current Queue Cycle"] || 0),
            "Active Window Generation": Number(currentState["Active Window Generation"] || 0)
        };

        // Ensure action is RECONCILE
        config.action = 'RECONCILE';

        const { nextState } = calculateNextQueueState_(parsedState, roster, config);

        // Serialize back
        const updates = {
            "Current Active Window": JSON.stringify(nextState["Current Active Window"]),
            "Current Directional Window": JSON.stringify(nextState["Current Directional Window"]),
            "Current Directional Window Completed": JSON.stringify(nextState["Current Directional Window Completed"]),
            "Current Queue Skip State": JSON.stringify(nextState["Current Queue Skip State"]),
            "Current Queue Cursor": nextState["Current Queue Cursor"].toString(),
            "Current Queue Cycle": nextState["Current Queue Cycle"].toString(),
            "Current Serpentine Direction": nextState["Current Serpentine Direction"],
            "Active Window Generation": nextState["Active Window Generation"].toString()
        };

        _writeConfig(updates);
        return { ok: true };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Processes a participant completion.
 */
function processQueueMutation_(authParticipantId, submittedTurnId, config, rosterCallback) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
        throw new Error('Could not obtain lock for queue mutation');
    }

    try {
        const _readConfig = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
        const _writeConfig = typeof writeConfigState_ === 'function' ? writeConfigState_ : (typeof global !== 'undefined' && global.writeConfigState_ ? global.writeConfigState_ : (require('./State.gs').writeConfigState));

        const currentState = _readConfig();
        const activeWindow = JSON.parse(currentState["Current Active Window"] || "[]");

        // Validate auth
        const targetTurn = activeWindow.find(t => t.turnId === submittedTurnId);
        if (!targetTurn) {
             return { ok: false, message: 'Turn is no longer active' };
        }
        if (targetTurn.participantId !== authParticipantId) {
             return { ok: false, message: 'Wrong participant for turn' };
        }

        const roster = rosterCallback();


        const parsedState = {
            ...currentState,
            "Current Active Window": activeWindow,
            "Current Directional Window": JSON.parse(currentState["Current Directional Window"] || "[]"),
            "Current Directional Window Completed": JSON.parse(currentState["Current Directional Window Completed"] || "[]"),
            "Current Queue Skip State": JSON.parse(currentState["Current Queue Skip State"] || "{}"),
            "Current Queue Cursor": Number(currentState["Current Queue Cursor"] || 0),
            "Current Queue Cycle": Number(currentState["Current Queue Cycle"] || 0),
            "Active Window Generation": Number(currentState["Active Window Generation"] || 0)
        };

        config.action = 'COMPLETE';
        config.completedTurnId = submittedTurnId;

        const { nextState, queueComplete, completionReason } = calculateNextQueueState_(parsedState, roster, config);

        // Serialize back
        const updates = {
            "Current Active Window": JSON.stringify(nextState["Current Active Window"]),
            "Current Directional Window": JSON.stringify(nextState["Current Directional Window"]),
            "Current Directional Window Completed": JSON.stringify(nextState["Current Directional Window Completed"]),
            "Current Queue Skip State": JSON.stringify(nextState["Current Queue Skip State"]),
            "Current Queue Cursor": nextState["Current Queue Cursor"].toString(),
            "Current Queue Cycle": nextState["Current Queue Cycle"].toString(),
            "Current Serpentine Direction": nextState["Current Serpentine Direction"],
            "Active Window Generation": nextState["Active Window Generation"].toString()
        };

        _writeConfig(updates);
        return { ok: true, queueComplete, completionReason };
    } finally {
        lock.releaseLock();
    }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports.processQueueMutation = processQueueMutation_;
  module.exports.processReconcile = processReconcile_;
}
