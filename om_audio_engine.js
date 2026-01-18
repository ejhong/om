/**
 * OM Audio Engine - Shared audio synthesis for OM visualization tools
 * Requires Tone.js to be loaded first
 */

// Optimize for smooth playback over low latency
Tone.context.latencyHint = 'playback';

// White key semitone offsets from C (c1=C, c2=D, c3=E, c4=F, c5=G, c6=A, c7=B, c8=C+octave)
const WHITE_KEY_SEMITONES = {
  1: 0,   // C
  2: 2,   // D
  3: 4,   // E
  4: 5,   // F
  5: 7,   // G
  6: 9,   // A
  7: 11,  // B
  8: 12   // C (next octave)
};

const SEMITONE_TO_NOTE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];

// OM playback defaults
const OM_DEFAULTS = {
  duration: 6,        // Total duration in seconds at 1x speed
  overlapRatio: 1.0,  // Note overlap for blending effect (was 1.8, reduced for CPU)
  defaultVolume: 24   // Default volume in dB
};

// Global tuning offset in semitones (0 = C4 reference, -12 = C3, +12 = C5)
let tuningOffset = 0;

// Global master limiter to prevent clipping/clicking
let masterLimiter = null;
function getMasterLimiter() {
  if (!masterLimiter) {
    masterLimiter = new Tone.Limiter(-3).toDestination();
  }
  return masterLimiter;
}

/**
 * Set the global tuning offset
 * @param {number} semitones - Offset in semitones (-12 to +12)
 */
function setTuning(semitones) {
  tuningOffset = semitones;
}

/**
 * Get the current tuning offset
 * @returns {number} Current offset in semitones
 */
function getTuning() {
  return tuningOffset;
}

// Section volume envelope (M2 silent, others at 1.0)
// Based on analysis: 0.2 -> 1.0 (A2 peak) -> 0.2 (linear decline)
const SECTION_VOLUMES = {
  'A1': 1.00, 'A2': 1.00, 'A3': 1.00, 'A4': 1.00,
  'U1': 1.00, 'U2': 1.00, 'M1': 1.00, 'M2': 0.00
};

/**
 * Calculate velocity based on trajectory magnitude and section envelope
 */
function calculateVelocity(note) {
  const trajMagnitude = Math.max(Math.abs(note.trajStart || 0), Math.abs(note.trajEnd || 0));
  const baseVelocity = 0.6 + trajMagnitude * 0.2;
  const sectionVol = SECTION_VOLUMES[note.group] ?? 1.0;
  if (sectionVol === 0) return 0;
  return Math.max(0.1, Math.min(1, baseVelocity * sectionVol));
}

/**
 * Play a note with Transport scheduling (for precise timing)
 * @param {Object} note - The note object from parseVoiceSpec
 * @param {number} duration - Duration in seconds
 * @param {number} time - Tone.Transport time
 * @param {number} volumeDb - Volume in dB
 * @returns {Object} synthObj for manual disposal if needed
 */
function playNoteScheduled(note, duration, time, volumeDb = 0) {
  const velocity = calculateVelocity(note);
  if (velocity === 0) return null;  // Skip silent notes
  const synthObj = createSynth(note, volumeDb);
  synthObj.synth.triggerAttackRelease(note.startFrequency, duration, time, velocity);

  // Calculate when the note actually starts relative to now (for timeouts)
  const startDelayMs = Math.max(0, (time - Tone.now()) * 1000);

  // Frequency glide if needed
  if (note.startFrequency !== note.endFrequency) {
    setTimeout(() => {
      try {
        synthObj.synth.frequency.rampTo(note.endFrequency, duration * 0.8);
      } catch (e) {}
    }, startDelayMs + 50);
  }

  // Auto-cleanup after note finishes (start delay + duration + small buffer for release)
  const cleanupDelay = startDelayMs + (duration + 0.5) * 1000;
  setTimeout(() => { try { synthObj.dispose(); } catch(e) {} }, cleanupDelay);

  return synthObj;
}

// Formant settings based on analysis of actual OM recordings
const FORMANTS = {
  'A': {  // "aah" - open vowel
    f1: { freq: 500, gain: 6 },
    f2: { freq: 1000, gain: 4 }
  },
  'U': {  // "ooh" - rounded vowel
    f1: { freq: 500, gain: 6 },
    f2: { freq: 1200, gain: -2 }
  },
  'M': {  // "mmm" - nasal hum
    f1: { freq: 300, gain: 4 },
    f2: { freq: 800, gain: -6 }
  }
};

/**
 * Parse a voice spec like v/tlk_nrm/mdl/vbr/ol1/c1/sh8[-0.66]
 */
function parseVoiceSpec(spec) {
  // Try parsing with scale position (sl1-sl8 or sh1-sh8)
  let match = spec.match(/v\/(\w+)_(\w+)\/(\w+)\/(\w+)\/ol(\d)\/(\w)(\d)\/(s[lh])(\d)(?:\[([-\d.:]+)\])?/);

  if (!match) {
    // Fallback: old format with just sh8
    match = spec.match(/v\/(\w+)_(\w+)\/(\w+)\/(\w+)\/ol(\d)\/(\w)(\d)\/sh(\d)(?:\[([-\d.:]+)\])?/);
    if (match) {
      const [_, voiceType, mode, register, articulation, octaveLayer, pitchClass, pitchNum, scaleNum, arg] = match;
      return buildNote(spec, voiceType, mode, register, articulation, octaveLayer, pitchClass, pitchNum, 'sh', scaleNum, arg);
    }
    return null;
  }

  const [_, voiceType, mode, register, articulation, octaveLayer, pitchClass, pitchNum, scaleType, scaleNum, arg] = match;
  return buildNote(spec, voiceType, mode, register, articulation, octaveLayer, pitchClass, pitchNum, scaleType, scaleNum, arg);
}

function buildNote(spec, voiceType, mode, register, articulation, octaveLayer, pitchClass, pitchNum, scaleType, scaleNum, arg) {
  const baseOctave = parseInt(octaveLayer);
  const whiteKeyOffset = WHITE_KEY_SEMITONES[parseInt(pitchNum)] || 0;
  const scaleWhiteKey = WHITE_KEY_SEMITONES[parseInt(scaleNum)] || 0;
  const scaleOffset = scaleType === 'sh' ? scaleWhiteKey + 12 : scaleWhiteKey;

  // Compute KEY (ol + pitch, before scale offset)
  const keyExtraOctaves = Math.floor(whiteKeyOffset / 12);
  const keySemitone = whiteKeyOffset % 12;
  const keyOctave = baseOctave + keyExtraOctaves;
  const keyNoteName = SEMITONE_TO_NOTE[keySemitone];
  const keyNote = `${keyNoteName}${keyOctave}`;

  // Scale index 0-15 (sl1=0, sl8=7, sh1=8, sh8=15)
  const scaleIndex = (scaleType === 'sh' ? 8 : 0) + (parseInt(scaleNum) - 1);

  let totalSemitones = whiteKeyOffset + scaleOffset;

  // Falsetto shifts up an octave
  if (register === 'fls') {
    totalSemitones += 12;
  }

  const extraOctaves = Math.floor(totalSemitones / 12);
  const semitoneInOctave = totalSemitones % 12;
  const finalOctave = baseOctave + extraOctaves;
  const noteName = SEMITONE_TO_NOTE[semitoneInOctave];

  // Parse trajectory (supports range format "start:end")
  let trajStart = 0, trajEnd = 0;
  if (arg) {
    if (arg.includes(':')) {
      const [s, e] = arg.split(':').map(parseFloat);
      trajStart = s;
      trajEnd = e;
    } else {
      trajStart = trajEnd = parseFloat(arg);
    }
  }

  const trajStartSemitones = trajStart * 2;
  const trajEndSemitones = trajEnd * 2;

  // Calculate frequencies (apply global tuning offset)
  const baseSemitonesFromC4 = (finalOctave - 4) * 12 + semitoneInOctave + tuningOffset;
  const startFrequency = 261.63 * Math.pow(2, (baseSemitonesFromC4 + trajStartSemitones) / 12);
  const endFrequency = 261.63 * Math.pow(2, (baseSemitonesFromC4 + trajEndSemitones) / 12);

  return {
    raw: spec,
    voiceType,
    mode,
    register,
    articulation,
    octaveLayer: parseInt(octaveLayer),
    pitchClass: pitchClass.toUpperCase(),
    pitchNum: parseInt(pitchNum),
    scaleType,
    scaleNum: parseInt(scaleNum),
    keyNote,
    scaleIndex,
    trajStart,
    trajEnd,
    trajStartSemitones,
    trajEndSemitones,
    noteName: `${noteName}${finalOctave}`,
    startFrequency,
    endFrequency,
    frequency: startFrequency,
    octave: finalOctave
  };
}

/**
 * Parse OM preset text into groups
 */
function parseOMInput(text) {
  const groups = {};
  const lines = text.trim().split('\n');

  for (const line of lines) {
    const groupMatch = line.match(/^(\w+):\s*(.+)$/);
    if (groupMatch) {
      const [_, groupName, notesStr] = groupMatch;
      const specs = notesStr.split(/\s*\+\s*/).map(s => s.trim()).filter(s => s);
      const notes = specs.map(parseVoiceSpec).filter(n => n);
      notes.forEach(n => {
        n.group = groupName;
        n.section = groupName.charAt(0).toUpperCase();
      });
      if (notes.length > 0) {
        groups[groupName] = notes;
      }
    }
  }
  return groups;
}

/**
 * Create formant filters for A-U-M vowel character
 */
function createFormantFilters(section) {
  const formant = FORMANTS[section] || FORMANTS['A'];

  const f1 = new Tone.Filter({
    frequency: formant.f1.freq,
    type: 'peaking',
    gain: formant.f1.gain,
    Q: 2
  });

  const f2 = new Tone.Filter({
    frequency: formant.f2.freq,
    type: 'peaking',
    gain: formant.f2.gain,
    Q: 2
  });

  f1.connect(f2);

  return {
    input: f1,
    output: f2,
    dispose: () => {
      f1.dispose();
      f2.dispose();
    }
  };
}

/**
 * Create synth based on voice type, mode, register, and articulation
 */
function createSynth(note, volumeDb = -6) {
  const { voiceType, mode, register, articulation, section } = note;
  const isFalsetto = register === 'fls';
  const isOperatic = mode === 'opr';

  // Mode-based synth parameters - THIS is what makes the big difference
  // Operatic: high modulation = rich harmonics, slow attack = legato phrasing
  // Normal: low modulation = simpler sound, faster attack = speech-like
  const isVbr = articulation === 'vbr';
  const isSld = articulation === 'sld';
  // For sld: kill ALL FM modulation to get clean sound
  // For vbr: minimal FM, let Vibrato effect handle it
  const modEnvSus = isSld ? 0 : 0.05;
  const modIndexMult = isSld ? 0.05 : 1.0;  // nearly zero FM for sld
  const modeParams = isOperatic ? {
    modBoost: 1.1 * modIndexMult,       // Richer harmonics
    harmBoost: 1.0,      // More harmonics
    atkMult: 1.1,        // Slower attack (legato)
    susBoost: 0.1,       // More sustain
    oscType: 'triangle', // Triangle for both modes
    volBoost: 3          // Louder
  } : {
    modBoost: 0.4 * modIndexMult,       // Less modulation
    harmBoost: 0.7,      // Fewer harmonics
    atkMult: 0.5,        // Faster attack (speech-like)
    susBoost: -0.1,      // Less sustain
    oscType: 'triangle', // Original triangle for normal
    volBoost: 0          // No volume adjustment
  };

  // Voice type variations combined with mode
  // modEnvSus controls FM wobble: 0 for sld (clean), low for vbr (Vibrato effect handles it)
  let synth;
  switch (voiceType) {
    case 'sng':
      synth = new Tone.FMSynth({
        harmonicity: (isFalsetto ? 4 : 3) * modeParams.harmBoost,
        modulationIndex: (isFalsetto ? 5 : 10) * modeParams.modBoost,
        oscillator: { type: isFalsetto ? 'triangle' : modeParams.oscType },
        envelope: { attack: 1.2 * modeParams.atkMult, decay: 0.3, sustain: Math.min(0.95, 0.85 + modeParams.susBoost), release: 0.8 },
        modulation: { type: 'sine' },
        modulationEnvelope: { attack: 1.0 * modeParams.atkMult, decay: 0.3, sustain: modEnvSus, release: 0.6 }
      });
      break;
    case 'ydl':
      // ydl: use sine oscillator for sld (clean), triangle for vbr (character)
      synth = new Tone.FMSynth({
        harmonicity: (isFalsetto ? 4 : 3) * modeParams.harmBoost,
        modulationIndex: (isFalsetto ? 4 : 8) * (isSld ? 0.1 : modeParams.modBoost),  // very low FM for sld
        oscillator: { type: isSld ? 'sine' : 'triangle' },  // sine for clean sld
        envelope: { attack: 1.0 * modeParams.atkMult, decay: 0.2, sustain: Math.min(0.95, 0.8 + modeParams.susBoost), release: 0.7 },
        modulation: { type: 'sine' },
        modulationEnvelope: { attack: 0.8 * modeParams.atkMult, decay: 0.2, sustain: modEnvSus, release: 0.5 }
      });
      break;
    case 'tlk':
      synth = new Tone.FMSynth({
        harmonicity: (isFalsetto ? 2.5 : 2) * modeParams.harmBoost,
        modulationIndex: (isFalsetto ? 3 : 6) * modeParams.modBoost,
        oscillator: { type: modeParams.oscType },
        envelope: { attack: 0.9 * modeParams.atkMult, decay: 0.2, sustain: Math.min(0.95, 0.75 + modeParams.susBoost), release: 0.6 },
        modulation: { type: isOperatic ? 'sine' : 'triangle' },
        modulationEnvelope: { attack: 0.7 * modeParams.atkMult, decay: 0.15, sustain: modEnvSus, release: 0.5 }
      });
      break;
    case 'rap':
      synth = new Tone.FMSynth({
        harmonicity: (isFalsetto ? 3 : 2) * modeParams.harmBoost,
        modulationIndex: (isFalsetto ? 2.5 : 5) * modeParams.modBoost,
        oscillator: { type: isFalsetto ? 'triangle' : modeParams.oscType },
        envelope: { attack: 0.8 * modeParams.atkMult, decay: 0.2, sustain: Math.min(0.95, 0.7 + modeParams.susBoost), release: 0.5 },
        modulation: { type: 'sine' },
        modulationEnvelope: { attack: 0.6 * modeParams.atkMult, decay: 0.15, sustain: modEnvSus, release: 0.4 }
      });
      break;
    default:
      synth = new Tone.FMSynth({
        harmonicity: 2 * modeParams.harmBoost,
        modulationIndex: 4 * modeParams.modBoost,
        envelope: { attack: 0.8 * modeParams.atkMult, decay: 0.3, sustain: 0.8, release: 0.6 },
        modulationEnvelope: { sustain: modEnvSus }
      });
  }

  const baseVol = isFalsetto ? volumeDb - 6 : volumeDb;
  const volume = new Tone.Volume(baseVol + modeParams.volBoost);
  let vibrato = null;
  let operaticFilters = [];

  // NOTE: Chorus for ydl was removed - too CPU intensive, caused audio cutout

  // Voice mode filters
  if (isOperatic) {
    // Operatic: boost chest resonance, add singer's formant ring
    operaticFilters.push(new Tone.Filter({ frequency: 2800, type: 'peaking', gain: 8, Q: 3 }));  // singer's ring
    operaticFilters.push(new Tone.Filter({ frequency: 400, type: 'peaking', gain: 6, Q: 1 }));  // chest warmth
  } else {
    // Normal: thinner, more nasal, speech-like
    operaticFilters.push(new Tone.Filter({ frequency: 300, type: 'highpass', rolloff: -12 }));  // reduce bass
    operaticFilters.push(new Tone.Filter({ frequency: 3000, type: 'highshelf', gain: 4 }));  // add brightness
  }

  // Formant filters for A-U-M vowels
  let formantFilters = null;
  if (section) {
    formantFilters = createFormantFilters(section);
  }

  // Vibrato for vbr articulation
  if (isVbr) {
    vibrato = new Tone.Vibrato({ frequency: 5.5, depth: 0.5 });
  }

  // Chain: synth -> operatic? -> vibrato? -> formants? -> volume -> limiter
  let chain = synth;
  for (const filter of operaticFilters) {
    chain.connect(filter);
    chain = filter;
  }
  if (vibrato) {
    chain.connect(vibrato);
    chain = vibrato;
  }
  if (formantFilters) {
    chain.connect(formantFilters.input);
    chain = formantFilters.output;
  }
  chain.connect(volume);
  volume.connect(getMasterLimiter());

  return {
    synth,
    effects: [volume, vibrato, ...operaticFilters].filter(Boolean),
    formantFilters,
    dispose: () => {
      try {
        synth.dispose();
        volume.dispose();
        if (vibrato) vibrato.dispose();
        operaticFilters.forEach(f => f.dispose());
        if (formantFilters) formantFilters.dispose();
      } catch (e) {
        // Ignore disposal errors
      }
    }
  };
}

/**
 * Play a single note
 */
async function playNote(note, durationSec, volumeDb = -6, onFinish = null) {
  try {
    await Tone.start();

    const synthObj = createSynth(note, volumeDb);
    const { synth } = synthObj;

    // Base velocity, slightly varied by trajectory magnitude
    const trajMagnitude = Math.max(Math.abs(note.trajStart), Math.abs(note.trajEnd));
    const velocity = 0.6 + trajMagnitude * 0.2;

    synth.triggerAttackRelease(note.startFrequency, durationSec, undefined, Math.max(0.1, Math.min(1, velocity)));

    // Frequency glide if needed
    if (note.startFrequency !== note.endFrequency) {
      setTimeout(() => {
        try {
          synth.frequency.rampTo(note.endFrequency, durationSec * 0.8);
        } catch (e) {}
      }, 50);
    }

    // Cleanup after note
    const cleanupDelay = durationSec * 1000 + 800;
    setTimeout(() => {
      synthObj.dispose();
      if (onFinish) onFinish();
    }, cleanupDelay);

    return synthObj;
  } catch (e) {
    console.error('Error playing note:', e);
    return null;
  }
}

/**
 * Render OM audio offline to a buffer for glitch-free playback
 * @param {Array} notes - Array of parsed note objects
 * @param {number} totalDuration - Total duration in seconds
 * @param {Object} options - { volumeDb, overlapRatio }
 * @returns {Promise<{player: Tone.Player, buffer: Tone.ToneAudioBuffer}>}
 */
async function renderOMOffline(notes, totalDuration, options = {}) {
  const {
    volumeDb = -6,
    overlapRatio = OM_DEFAULTS.overlapRatio
  } = options;

  if (!notes || notes.length === 0) {
    return null;
  }

  const noteCount = notes.length;
  const noteDuration = totalDuration / noteCount;
  const noteSoundDuration = noteDuration * overlapRatio;

  // Add extra time for release tails
  const renderDuration = totalDuration + 3;

  // Render offline - this happens without real-time constraints
  const buffer = await Tone.Offline(() => {
    // Create a limiter for the offline context
    const limiter = new Tone.Limiter(-3).toDestination();

    // Create all synths and schedule their notes upfront
    for (let i = 0; i < noteCount; i++) {
      const note = notes[i];
      const noteTime = i * noteDuration;

      // Calculate velocity
      const velocity = calculateVelocity(note);
      if (velocity === 0) continue; // Skip silent notes (M2)

      // Create synth for this note
      const synthObj = createSynthForOffline(note, volumeDb, limiter);
      if (!synthObj) continue;

      const { synth } = synthObj;

      // Schedule the note directly with absolute time
      synth.triggerAttackRelease(note.startFrequency, noteSoundDuration, noteTime, velocity);

      // Frequency glide if needed
      if (note.startFrequency !== note.endFrequency) {
        synth.frequency.rampTo(note.endFrequency, noteSoundDuration * 0.8, noteTime + 0.05);
      }
    }
  }, renderDuration);

  // Create a player from the buffer
  const player = new Tone.Player(buffer).toDestination();

  return { player, buffer, duration: totalDuration };
}

/**
 * Create synth for offline rendering (simplified for speed)
 */
function createSynthForOffline(note, volumeDb, destination) {
  const { voiceType, mode, register, articulation, section } = note;
  const isFalsetto = register === 'fls';
  const isVbr = articulation === 'vbr';

  // Simplified synth - just vary harmonicity and modulation index by voice type
  const harmonicity = voiceType === 'sng' ? 3 : voiceType === 'ydl' ? 3 : 2;
  const modIndex = isFalsetto ? 4 : 8;

  const synth = new Tone.FMSynth({
    harmonicity,
    modulationIndex: modIndex,
    oscillator: { type: isFalsetto ? 'triangle' : 'sine' },
    envelope: { attack: 1.0, decay: 0.3, sustain: 0.8, release: 0.8 },
    modulation: { type: 'sine' },
    modulationEnvelope: { attack: 0.8, decay: 0.3, sustain: 0.5, release: 0.5 }
  });

  // Minimal effect chain for speed
  const vol = new Tone.Volume(isFalsetto ? volumeDb - 6 : volumeDb);

  if (isVbr) {
    const vibrato = new Tone.Vibrato({ frequency: 5.5, depth: 0.5 });
    synth.connect(vibrato);
    vibrato.connect(vol);
  } else {
    synth.connect(vol);
  }

  vol.connect(destination);

  return { synth };
}

/**
 * OM Offline Player - plays pre-rendered audio with visualization sync
 */
class OMOfflinePlayer {
  constructor() {
    this.player = null;
    this.buffer = null;
    this.duration = 0;
    this.isPlaying = false;
    this.isRendering = false;
    this.startTime = 0;
    this.pausedAt = 0;
    this.animationFrame = null;
    this.onProgress = null; // Callback: (currentTime, noteIndex) => void
    this.onRenderStart = null;
    this.onRenderComplete = null;
    this.onPlaybackEnd = null;
  }

  /**
   * Pre-render audio for given notes
   */
  async render(notes, totalDuration, options = {}) {
    if (this.isRendering) return;

    this.stop();
    this.isRendering = true;

    if (this.onRenderStart) {
      this.onRenderStart();
    }

    try {
      // Dispose old player
      if (this.player) {
        this.player.dispose();
        this.player = null;
      }

      // Render new audio
      const result = await renderOMOffline(notes, totalDuration, options);

      if (result) {
        this.player = result.player;
        this.buffer = result.buffer;
        this.duration = result.duration;
        this.notes = notes;
        this.noteDuration = totalDuration / notes.length;
      }

      if (this.onRenderComplete) {
        this.onRenderComplete();
      }
    } catch (e) {
      console.error('Render error:', e);
    } finally {
      this.isRendering = false;
    }
  }

  /**
   * Start playback from current position
   */
  async play(fromTime = null) {
    if (!this.player || this.isRendering) return;

    await Tone.start();

    const startOffset = fromTime !== null ? fromTime : this.pausedAt;
    this.startTime = Tone.now() - startOffset;
    this.isPlaying = true;

    // Start player from offset
    this.player.start(Tone.now(), startOffset);

    // Start progress animation
    this._animateProgress();

    // Schedule end
    const remaining = this.duration - startOffset;
    setTimeout(() => {
      if (this.isPlaying) {
        this.finish();
      }
    }, remaining * 1000 + 100);
  }

  /**
   * Pause playback
   */
  pause() {
    if (!this.isPlaying) return;

    this.isPlaying = false;
    this.pausedAt = Tone.now() - this.startTime;

    if (this.player) {
      this.player.stop();
    }

    if (this.animationFrame) {
      cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }
  }

  /**
   * Stop and reset
   */
  stop() {
    this.isPlaying = false;
    this.pausedAt = 0;

    if (this.player) {
      try { this.player.stop(); } catch(e) {}
    }

    if (this.animationFrame) {
      cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }
  }

  /**
   * Seek to time position
   */
  seek(time) {
    const wasPlaying = this.isPlaying;
    this.stop();
    this.pausedAt = Math.max(0, Math.min(time, this.duration));

    if (wasPlaying) {
      this.play();
    } else if (this.onProgress) {
      const noteIndex = Math.floor(this.pausedAt / this.noteDuration);
      this.onProgress(this.pausedAt, noteIndex);
    }
  }

  /**
   * Seek to note index
   */
  seekToNote(noteIndex) {
    const time = noteIndex * this.noteDuration;
    this.seek(time);
  }

  /**
   * Get current playback time
   */
  getCurrentTime() {
    if (this.isPlaying) {
      return Tone.now() - this.startTime;
    }
    return this.pausedAt;
  }

  /**
   * Get current note index
   */
  getCurrentNoteIndex() {
    const time = this.getCurrentTime();
    return Math.min(Math.floor(time / this.noteDuration), (this.notes?.length || 1) - 1);
  }

  /**
   * Playback finished
   */
  finish() {
    this.isPlaying = false;
    this.pausedAt = this.duration;

    if (this.player) {
      try { this.player.stop(); } catch(e) {}
    }

    if (this.animationFrame) {
      cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }

    if (this.onPlaybackEnd) {
      this.onPlaybackEnd();
    }
  }

  /**
   * Animation loop for progress updates
   */
  _animateProgress() {
    if (!this.isPlaying) return;

    const currentTime = this.getCurrentTime();
    const noteIndex = this.getCurrentNoteIndex();

    if (this.onProgress) {
      this.onProgress(currentTime, noteIndex);
    }

    this.animationFrame = requestAnimationFrame(() => this._animateProgress());
  }

  /**
   * Dispose resources
   */
  dispose() {
    this.stop();
    if (this.player) {
      this.player.dispose();
      this.player = null;
    }
  }
}

/**
 * OM Audio Player class for managing playback state
 */
class OMAudioPlayer {
  constructor() {
    this.isPlaying = false;
    this.currentSynth = null;
    this.playbackTimeout = null;
    this.notes = [];
    this.currentIndex = 0;
    this.volumeDb = -6;
    this.onNoteStart = null;
    this.onPlaybackEnd = null;
  }

  setVolume(db) {
    this.volumeDb = db;
  }

  setNotes(notes) {
    this.notes = notes;
    this.currentIndex = 0;
  }

  async start(totalDurationSec = 12) {
    if (this.isPlaying || this.notes.length === 0) return;

    await Tone.start();
    this.isPlaying = true;

    const noteDuration = totalDurationSec / this.notes.length;
    this._playNext(noteDuration);
  }

  _playNext(noteDuration) {
    if (!this.isPlaying || this.currentIndex >= this.notes.length) {
      this._finish();
      return;
    }

    const note = this.notes[this.currentIndex];

    if (this.currentSynth) {
      this.currentSynth.dispose();
    }

    this.currentSynth = createSynth(note, this.volumeDb);
    const trajMagnitude = Math.max(Math.abs(note.trajStart), Math.abs(note.trajEnd));
    const velocity = 0.6 + trajMagnitude * 0.2;

    this.currentSynth.synth.triggerAttackRelease(
      note.startFrequency,
      Math.max(0.1, noteDuration - 0.05),
      undefined,
      Math.max(0.1, Math.min(1, velocity))
    );

    if (this.onNoteStart) {
      this.onNoteStart(note, this.currentIndex);
    }

    this.currentIndex++;
    this.playbackTimeout = setTimeout(() => this._playNext(noteDuration), noteDuration * 1000);
  }

  stop() {
    this.isPlaying = false;

    if (this.playbackTimeout) {
      clearTimeout(this.playbackTimeout);
      this.playbackTimeout = null;
    }

    if (this.currentSynth) {
      this.currentSynth.dispose();
      this.currentSynth = null;
    }
  }

  reset() {
    this.stop();
    this.currentIndex = 0;
  }

  _finish() {
    this.isPlaying = false;
    if (this.currentSynth) {
      this.currentSynth.dispose();
      this.currentSynth = null;
    }
    if (this.onPlaybackEnd) {
      this.onPlaybackEnd();
    }
  }

  seekTo(index) {
    const wasPlaying = this.isPlaying;
    this.stop();
    this.currentIndex = Math.max(0, Math.min(index, this.notes.length));
    return wasPlaying;
  }
}

// Export for use as module or global
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseVoiceSpec,
    parseOMInput,
    createSynth,
    createFormantFilters,
    playNote,
    playNoteScheduled,
    calculateVelocity,
    renderOMOffline,
    OMAudioPlayer,
    OMOfflinePlayer,
    FORMANTS,
    WHITE_KEY_SEMITONES,
    SEMITONE_TO_NOTE,
    OM_DEFAULTS,
    SECTION_VOLUMES
  };
}
