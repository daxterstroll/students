/**
 * ASCII Christmas Tree — browser port
 * Drop into static/js/ and include:
 *   <script defer src="{{ url_for('static', filename='js/christmas_tree.js') }}"></script>
 *
 * Optional: place an element with id="christmas-tree" to host the canvas.
 * Controls: Space = pause, R = restart, S = screenshot, click = toggle pause.
 */
(function () {
  "use strict";

  // ---------- math helpers ----------
  const TAU = Math.PI * 2;
  const DEG = 180 / Math.PI;

  function clamp(x, a, b) {
    return Math.max(a, Math.min(b, x));
  }

  function smooth(x) {
    x = clamp(x, 0, 1);
    return x * x * (3 - 2 * x);
  }

  function pop(x) {
    x = clamp(x, 0, 1) - 1;
    return 1 + 2.7 * x * x * x + 1.7 * x * x;
  }

  function unit3(x, y, z) {
    const l = Math.hypot(x, y, z) || 1;
    return [x / l, y / l, z / l];
  }

  function mul3(v, s) {
    return [v[0] * s, v[1] * s, v[2] * s];
  }

  function add3(a, b) {
    return [a[0] + b[0], a[1] + b[1], a[2] + b[2]];
  }

  function hz(w, LOOP) {
    if (!LOOP) return w;
    return Math.max(1, Math.round((w * LOOP) / TAU)) * (TAU / LOOP);
  }

  // ---------- world constants ----------
  const GROUND = -0.98;
  const Y0 = -0.80;
  const Y1 = 0.86;
  const RMAX = 0.58;

  const RED = [0.95, 0.12, 0.1];
  const GOLD = [1.0, 0.76, 0.22];
  const SILVER = [0.86, 0.88, 0.95];
  const BLUE = [0.15, 0.38, 1.0];
  const GREEN_BOX = [0.1, 0.72, 0.25];
  const WARM = [1.0, 0.84, 0.55];
  const WHITE = [1.0, 1.0, 1.0];
  const BARK = [0.45, 0.27, 0.12];

  const T_GROW = [0.4, 6.5];
  const T_GARLAND = [6.5, 8.8];
  const T_BALLS = [8.0, 10.5];
  const T_LIGHTS = [9.0, 11.5];
  const T_GIFTS = [10.0, 12.0];
  const T_STAR = 12.0;
  const T_MSG = 13.0;
  const SPIN = TAU / 20;

  let LOOP = null;

  const SCALE = 28.0;
  const HORIZON = 0.45;
  const CAM = 4.0;
  const TILT = 0.22;
  const LIGHT = unit3(-0.5, 0.6, -0.65);

  const MESSAGE = "Щасливого Різдва 2026".split("").join(" ").replace(/   /g, "     ");

  const NEEDLE = 0,
    SNOW = 1,
    TRUNK = 2,
    GARLAND = 3,
    BALL = 4,
    BULB = 5,
    GIFT = 6,
    RIBBON = 7,
    FLOOR = 8;

  const RAMPS = {
    [NEEDLE]: ".:-~=+*xv%&#@W",
    [SNOW]: "-=+*#W@",
    [TRUNK]: ".:=#%",
    [GARLAND]: "-~=*@",
    [BALL]: ".:o%&@",
    [BULB]: "oO0@",
    [GIFT]: ".-=+#%@",
    [RIBBON]: "+xX#",
    [FLOOR]: "..-",
  };

  const LEVELS = 16;
  const RAMP = new Int32Array(9 * LEVELS);
  for (const k of Object.keys(RAMPS)) {
    const r = RAMPS[k];
    for (let j = 0; j < LEVELS; j++) {
      RAMP[k * LEVELS + j] = r.charCodeAt(Math.floor((j * r.length) / LEVELS));
    }
  }

  const BULB_PTS = 16;

  // ---------- seeded RNG (mulberry32) ----------
  function mulberry32(a) {
    return function () {
      let t = (a += 0x6d2b79f5);
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // ---------- geometry helpers ----------
  function envelope(y) {
    const u = clamp((y - Y0) / (Y1 - Y0), 0, 1);
    return RMAX * Math.pow(1 - u, 1.05);
  }

  function grow_time(y) {
    return T_GROW[0] + (T_GROW[1] - T_GROW[0]) * clamp((y - GROUND) / (Y1 - GROUND), 0, 1);
  }

  function fib_sphere(n) {
    const pts = new Float32Array(n * 3);
    for (let i = 0; i < n; i++) {
      const ii = i + 0.5;
      const phi = Math.acos(1 - (2 * ii) / n);
      const th = Math.PI * (1 + Math.sqrt(5)) * ii;
      const s = Math.sin(phi);
      pts[i * 3] = Math.cos(th) * s;
      pts[i * 3 + 1] = Math.cos(phi);
      pts[i * 3 + 2] = Math.sin(th) * s;
    }
    return pts;
  }

  // ---------- Parts accumulator ----------
  class Parts {
    constructor() {
      this.pos = [];
      this.nrm = [];
      this.col = [];
      this.kind = [];
      this.t0 = [];
      this.gain = [];
      this.centre = [];
      this.phase = [];
      this.pop = [];
    }

    add(pos, nrm, col, kind, t0, gain = 1, centre = null, phase = 0, isPop = false) {
      const n = pos.length / 3;
      for (let i = 0; i < n; i++) {
        this.pos.push(pos[i * 3], pos[i * 3 + 1], pos[i * 3 + 2]);
        this.nrm.push(nrm[i * 3], nrm[i * 3 + 1], nrm[i * 3 + 2]);
        this.col.push(col[i * 3], col[i * 3 + 1], col[i * 3 + 2]);
        this.kind.push(kind[i] !== undefined ? kind[i] : kind);
        this.t0.push(t0[i] !== undefined ? t0[i] : t0);
        this.gain.push(gain[i] !== undefined ? gain[i] : gain);
        if (centre) {
          this.centre.push(centre[i * 3], centre[i * 3 + 1], centre[i * 3 + 2]);
        } else {
          this.centre.push(pos[i * 3], pos[i * 3 + 1], pos[i * 3 + 2]);
        }
        this.phase.push(phase[i] !== undefined ? phase[i] : phase);
        this.pop.push(isPop[i] !== undefined ? isPop[i] : isPop);
      }
    }

    freeze(rng) {
      const n = this.pos.length / 3;
      this.N = n;
      this.pos = new Float32Array(this.pos);
      this.nrm = new Float32Array(this.nrm);
      this.col = new Float32Array(this.col);
      this.kind = new Int32Array(this.kind);
      this.t0 = new Float32Array(this.t0);
      this.gain = new Float32Array(this.gain);
      this.centre = new Float32Array(this.centre);
      this.phase = new Float32Array(this.phase);
      this.pop = new Uint8Array(this.pop.map((v) => (v ? 1 : 0)));
      this.jitter = new Float32Array(n);
      for (let i = 0; i < n; i++) this.jitter[i] = rng() * 2 - 1;
      return this;
    }
  }

  function box(rng, w, h, d, m, rib = 0.022) {
    const areas = [w * d, w * d, h * d, h * d, w * h, w * h];
    let sum = 0;
    for (const a of areas) sum += a;
    const face = new Int32Array(m);
    for (let i = 0; i < m; i++) {
      let r = rng() * sum;
      for (let f = 0; f < 6; f++) {
        r -= areas[f];
        if (r <= 0) {
          face[i] = f;
          break;
        }
      }
    }
    const a = new Float32Array(m);
    const b = new Float32Array(m);
    for (let i = 0; i < m; i++) {
      a[i] = rng() - 0.5;
      b[i] = rng() - 0.5;
    }
    const p = new Float32Array(m * 3);
    const n = new Float32Array(m * 3);
    const ribbon = new Uint8Array(m);

    for (let i = 0; i < m; i++) {
      const f = face[i];
      let px, py, pz, nx, ny, nz;
      if (f === 0) {
        px = a[i] * w;
        py = h;
        pz = b[i] * d;
        nx = 0;
        ny = 1;
        nz = 0;
      } else if (f === 1) {
        px = a[i] * w;
        py = 0;
        pz = b[i] * d;
        nx = 0;
        ny = -1;
        nz = 0;
      } else if (f === 2) {
        px = w / 2;
        py = (a[i] + 0.5) * h;
        pz = b[i] * d;
        nx = 1;
        ny = 0;
        nz = 0;
      } else if (f === 3) {
        px = -w / 2;
        py = (a[i] + 0.5) * h;
        pz = b[i] * d;
        nx = -1;
        ny = 0;
        nz = 0;
      } else if (f === 4) {
        px = a[i] * w;
        py = (b[i] + 0.5) * h;
        pz = d / 2;
        nx = 0;
        ny = 0;
        nz = 1;
      } else {
        px = a[i] * w;
        py = (b[i] + 0.5) * h;
        pz = -d / 2;
        nx = 0;
        ny = 0;
        nz = -1;
      }
      p[i * 3] = px;
      p[i * 3 + 1] = py;
      p[i * 3 + 2] = pz;
      n[i * 3] = nx;
      n[i * 3 + 1] = ny;
      n[i * 3 + 2] = nz;
      ribbon[i] = Math.abs(px) < rib || Math.abs(pz) < rib ? 1 : 0;
    }

    // bow
    const bowN = 120;
    const bowP = new Float32Array(bowN * 3);
    const bowNrm = new Float32Array(bowN * 3);
    const bowRib = new Uint8Array(bowN);
    for (let i = 0; i < bowN; i++) {
      const k = rng();
      const side = k < 0.5 ? -1 : 1;
      const ang = TAU * rng();
      bowP[i * 3] = side * (0.03 + 0.03 * Math.cos(ang));
      bowP[i * 3 + 1] = h + 0.03 + 0.03 * Math.sin(ang);
      bowP[i * 3 + 2] = 0;
      bowNrm[i * 3] = Math.cos(ang) * side;
      bowNrm[i * 3 + 1] = Math.sin(ang);
      bowNrm[i * 3 + 2] = -0.4;
      bowRib[i] = 1;
    }

    const outP = new Float32Array((m + bowN) * 3);
    const outN = new Float32Array((m + bowN) * 3);
    const outR = new Uint8Array(m + bowN);
    outP.set(p);
    outP.set(bowP, m * 3);
    outN.set(n);
    outN.set(bowNrm, m * 3);
    outR.set(ribbon);
    outR.set(bowRib, m);
    return [outP, outN, outR];
  }

  // ---------- build the tree ----------
  function build(cols, rows) {
    const rng = mulberry32(7);
    const q = new Parts();
    const tips = [];

    // needles
    for (let i = 0; i < 24; i++) {
      const y = Y0 + 0.02 + ((Y1 - 0.06 - Y0) * i) / 23;
      const u = (y - Y0) / (Y1 - Y0);
      const R = envelope(y) + 0.05;
      const nb = Math.floor(6 + 12 * (1 - u));
      for (let b = 0; b < nb; b++) {
        const a = (TAU * (b + 0.5 * rng())) / nb + i * 0.77;
        const L = R * (0.85 + 0.23 * rng());
        const n = Math.floor(60 + 380 * L);
        const pos = new Float32Array(n * 3);
        const nrm = new Float32Array(n * 3);
        const col = new Float32Array(n * 3);
        const kind = new Int32Array(n);
        const t0 = new Float32Array(n);
        const gain = new Float32Array(n);
        const d = [Math.cos(a), 0, Math.sin(a)];
        const side = [-Math.sin(a), 0, Math.cos(a)];

        for (let j = 0; j < n; j++) {
          const s = Math.pow(rng(), 0.65);
          const spread = (0.025 + 0.09 * s) * (rng() * 2 - 1);
          const sl = s * L;
          pos[j * 3] = d[0] * sl + side[0] * spread;
          pos[j * 3 + 1] = y + 0.06 * sl - 0.2 * sl * sl + 0.012 * (rng() * 2 - 1);
          pos[j * 3 + 2] = d[2] * sl + side[2] * spread;

          const nx = d[0] * 0.75 + 0.3 * (rng() * 2 - 1);
          const ny = 0.55 + 0.3 * (rng() * 2 - 1);
          const nz = d[2] * 0.75 + 0.3 * (rng() * 2 - 1);
          const nl = Math.hypot(nx, ny, nz) || 1;
          nrm[j * 3] = nx / nl;
          nrm[j * 3 + 1] = ny / nl;
          nrm[j * 3 + 2] = nz / nl;

          const ao = (0.55 + 0.45 * Math.pow(s, 1.3)) * (0.75 + 0.35 * rng());
          const snow = s > 0.5 && rng() < 0.06 * s && nrm[j * 3 + 1] > 0.5;
          if (snow) {
            col[j * 3] = 1;
            col[j * 3 + 1] = 1;
            col[j * 3 + 2] = 1;
            kind[j] = SNOW;
            gain[j] = 1;
          } else {
            col[j * 3] = 0.1 + rng() * 0.12;
            col[j * 3 + 1] = 0.48 + rng() * 0.38;
            col[j * 3 + 2] = 0.18 + rng() * 0.14;
            kind[j] = NEEDLE;
            gain[j] = ao;
          }
          t0[j] = grow_time(pos[j * 3 + 1]);
        }
        q.add(pos, nrm, col, kind, t0, gain);

        const tl = 0.93 * L;
        tips.push(d[0] * tl, y + 0.06 * tl - 0.2 * tl * tl, d[2] * tl);
      }
    }

    // dark inner cone
    {
      const m = 16000;
      const pos = new Float32Array(m * 3);
      const nrm = new Float32Array(m * 3);
      const col = new Float32Array(m * 3);
      const t0 = new Float32Array(m);
      for (let i = 0; i < m; i++) {
        const y = Y0 + (Y1 - 0.1 - Y0) * rng();
        const phi = rng() * TAU;
        const r = 0.8 * envelope(y);
        pos[i * 3] = Math.cos(phi) * r;
        pos[i * 3 + 1] = y;
        pos[i * 3 + 2] = Math.sin(phi) * r;
        nrm[i * 3] = Math.cos(phi);
        nrm[i * 3 + 1] = 0.3;
        nrm[i * 3 + 2] = Math.sin(phi);
        col[i * 3] = 0.12;
        col[i * 3 + 1] = 0.55;
        col[i * 3 + 2] = 0.2;
        t0[i] = grow_time(y);
      }
      q.add(pos, nrm, col, NEEDLE, t0, 0.8);
    }

    // trunk
    {
      const m = 1500;
      const pos = new Float32Array(m * 3);
      const nrm = new Float32Array(m * 3);
      const col = new Float32Array(m * 3);
      const t0 = new Float32Array(m);
      for (let i = 0; i < m; i++) {
        const y = GROUND + rng() * (Y0 + 0.1 - GROUND);
        const phi = rng() * TAU;
        nrm[i * 3] = Math.cos(phi);
        nrm[i * 3 + 1] = 0;
        nrm[i * 3 + 2] = Math.sin(phi);
        pos[i * 3] = nrm[i * 3] * 0.075;
        pos[i * 3 + 1] = y;
        pos[i * 3 + 2] = nrm[i * 3 + 2] * 0.075;
        col[i * 3] = BARK[0];
        col[i * 3 + 1] = BARK[1];
        col[i * 3 + 2] = BARK[2];
        t0[i] = grow_time(y);
      }
      q.add(pos, nrm, col, TRUNK, t0);
    }

    // snowy floor
    {
      const m = 2500;
      const pos = new Float32Array(m * 3);
      const nrm = new Float32Array(m * 3);
      const col = new Float32Array(m * 3);
      for (let i = 0; i < m; i++) {
        const r = 1.25 * Math.sqrt(rng());
        const phi = rng() * TAU;
        pos[i * 3] = Math.cos(phi) * r;
        pos[i * 3 + 1] = GROUND;
        pos[i * 3 + 2] = Math.sin(phi) * r;
        nrm[i * 3] = 0;
        nrm[i * 3 + 1] = 1;
        nrm[i * 3 + 2] = 0;
        col[i * 3] = 0.42;
        col[i * 3 + 1] = 0.45;
        col[i * 3 + 2] = 0.58;
      }
      q.add(pos, nrm, col, FLOOR, 0, 0.45);
    }

    // tinsel garland
    {
      const n = 2600;
      const u = new Float32Array(n);
      for (let i = 0; i < n; i++) u[i] = i / (n - 1);
      const turns = 3.6;
      for (let pass = 0; pass < 4; pass++) {
        const pos = new Float32Array(n * 3);
        const nrm = new Float32Array(n * 3);
        const col = new Float32Array(n * 3);
        const t0 = new Float32Array(n);
        const phase = new Float32Array(n);
        for (let i = 0; i < n; i++) {
          const yg = Y0 + 0.15 + (Y1 - Y0 - 0.45) * u[i];
          const ag = TAU * turns * u[i] + 0.9;
          const swag = -0.07 * Math.abs(Math.sin(ag * 2.5));
          const rg = envelope(yg) + 0.07;
          pos[i * 3] = Math.cos(ag) * rg + 0.016 * (rng() * 2 - 1);
          pos[i * 3 + 1] = yg + swag + 0.016 * (rng() * 2 - 1);
          pos[i * 3 + 2] = Math.sin(ag) * rg + 0.016 * (rng() * 2 - 1);
          nrm[i * 3] = Math.cos(ag) + 0.6 * (rng() * 2 - 1);
          nrm[i * 3 + 1] = 0.3 + 0.6 * (rng() * 2 - 1);
          nrm[i * 3 + 2] = Math.sin(ag) + 0.6 * (rng() * 2 - 1);
          const nl = Math.hypot(nrm[i * 3], nrm[i * 3 + 1], nrm[i * 3 + 2]) || 1;
          nrm[i * 3] /= nl;
          nrm[i * 3 + 1] /= nl;
          nrm[i * 3 + 2] /= nl;
          col[i * 3] = GOLD[0];
          col[i * 3 + 1] = GOLD[1];
          col[i * 3 + 2] = GOLD[2];
          t0[i] = T_GARLAND[0] + u[i] * (T_GARLAND[1] - T_GARLAND[0]);
          phase[i] = rng() * TAU;
        }
        q.add(pos, nrm, col, GARLAND, t0, 1, null, phase);
      }
    }

    // baubles
    {
      const sphere = fib_sphere(160);
      for (let k = 0; k < 34; k++) {
        const u = (k + rng()) / 34;
        const y = Y0 + 0.08 + u * (Y1 - Y0 - 0.38);
        const a = k * 2.39996 + rng() * 0.4;
        const cx = Math.cos(a) * envelope(y);
        const cy = y - 0.03;
        const cz = Math.sin(a) * envelope(y);
        const rad = 0.058 + 0.032 * (1 - u);
        let colour = RED;
        if (k % 2 !== 0) colour = k % 4 === 1 ? GOLD : SILVER;
        const n = 160;
        const pos = new Float32Array(n * 3);
        const nrm = new Float32Array(n * 3);
        const col = new Float32Array(n * 3);
        const centre = new Float32Array(n * 3);
        const t0 = T_BALLS[0] + u * (T_BALLS[1] - T_BALLS[0]) + 0.3 * rng();
        for (let i = 0; i < n; i++) {
          pos[i * 3] = cx + sphere[i * 3] * rad;
          pos[i * 3 + 1] = cy + sphere[i * 3 + 1] * rad;
          pos[i * 3 + 2] = cz + sphere[i * 3 + 2] * rad;
          nrm[i * 3] = sphere[i * 3];
          nrm[i * 3 + 1] = sphere[i * 3 + 1];
          nrm[i * 3 + 2] = sphere[i * 3 + 2];
          col[i * 3] = colour[0];
          col[i * 3 + 1] = colour[1];
          col[i * 3 + 2] = colour[2];
          centre[i * 3] = cx;
          centre[i * 3 + 1] = cy;
          centre[i * 3 + 2] = cz;
        }
        q.add(pos, nrm, col, BALL, t0, 1, centre, 0, true);
      }
    }

    // fairy lights
    {
      const tipCount = tips.length / 3;
      const pickN = Math.min(170, tipCount);
      const pick = new Int32Array(pickN);
      const used = new Uint8Array(tipCount);
      let filled = 0;
      while (filled < pickN) {
        const idx = Math.floor(rng() * tipCount);
        if (!used[idx]) {
          used[idx] = 1;
          pick[filled++] = idx;
        }
      }
      const bulb = fib_sphere(BULB_PTS);
      for (let p = 0; p < pickN; p++) {
        const ti = pick[p];
        const cx = tips[ti * 3];
        const cy = tips[ti * 3 + 1];
        const cz = tips[ti * 3 + 2];
        const t0 = T_LIGHTS[0] + rng() * (T_LIGHTS[1] - T_LIGHTS[0]);
        let colour = WARM;
        if (rng() >= 0.6) {
          const palette = [RED, GOLD, BLUE, GREEN_BOX, WHITE];
          colour = palette[Math.floor(rng() * 5)];
        }
        const n = BULB_PTS;
        const pos = new Float32Array(n * 3);
        const nrm = new Float32Array(n * 3);
        const col = new Float32Array(n * 3);
        const centre = new Float32Array(n * 3);
        const phase = rng() * TAU;
        for (let i = 0; i < n; i++) {
          pos[i * 3] = cx + bulb[i * 3] * 0.03;
          pos[i * 3 + 1] = cy + bulb[i * 3 + 1] * 0.03;
          pos[i * 3 + 2] = cz + bulb[i * 3 + 2] * 0.03;
          nrm[i * 3] = bulb[i * 3];
          nrm[i * 3 + 1] = bulb[i * 3 + 1];
          nrm[i * 3 + 2] = bulb[i * 3 + 2];
          col[i * 3] = colour[0];
          col[i * 3 + 1] = colour[1];
          col[i * 3 + 2] = colour[2];
          centre[i * 3] = cx;
          centre[i * 3 + 1] = cy;
          centre[i * 3 + 2] = cz;
        }
        q.add(pos, nrm, col, BULB, t0, 1, centre, phase, true);
      }
    }

    // presents
    {
      const gifts = [
        [0.3, 0.66, 0.24, 0.28, 0.24, 0.4, RED, GOLD],
        [1.5, 0.62, 0.18, 0.2, 0.18, 0.9, BLUE, GOLD],
        [2.6, 0.68, 0.3, 0.22, 0.26, 0.1, GOLD, RED],
        [3.7, 0.64, 0.2, 0.3, 0.2, 0.6, GREEN_BOX, RED],
        [4.7, 0.6, 0.26, 0.18, 0.3, 0.3, SILVER, RED],
        [5.6, 0.7, 0.22, 0.24, 0.22, 1.1, RED, SILVER],
      ];
      for (let g = 0; g < gifts.length; g++) {
        const [ang, dist, w, h, d, rot, body, rib] = gifts[g];
        const m = Math.floor(5000 * (w * h + w * d + h * d));
        const [bp, bn, ribbon] = box(rng, w, h, d, m);
        const cr = Math.cos(rot);
        const sr = Math.sin(rot);
        const npts = bp.length / 3;
        const pos = new Float32Array(npts * 3);
        const nrm = new Float32Array(npts * 3);
        const col = new Float32Array(npts * 3);
        const kind = new Int32Array(npts);
        const centre = new Float32Array(npts * 3);
        const baseX = Math.cos(ang) * dist;
        const baseZ = Math.sin(ang) * dist;
        const t0 = T_GIFTS[0] + (g * (T_GIFTS[1] - T_GIFTS[0])) / gifts.length;

        for (let i = 0; i < npts; i++) {
          const x = bp[i * 3];
          const y = bp[i * 3 + 1];
          const z = bp[i * 3 + 2];
          // rotate around Y
          const rx = x * cr - z * sr;
          const rz = x * sr + z * cr;
          pos[i * 3] = rx + baseX;
          pos[i * 3 + 1] = y + GROUND;
          pos[i * 3 + 2] = rz + baseZ;

          const nx = bn[i * 3];
          const ny = bn[i * 3 + 1];
          const nz = bn[i * 3 + 2];
          nrm[i * 3] = nx * cr - nz * sr;
          nrm[i * 3 + 1] = ny;
          nrm[i * 3 + 2] = nx * sr + nz * cr;

          if (ribbon[i]) {
            col[i * 3] = rib[0];
            col[i * 3 + 1] = rib[1];
            col[i * 3 + 2] = rib[2];
            kind[i] = RIBBON;
          } else {
            col[i * 3] = body[0];
            col[i * 3 + 1] = body[1];
            col[i * 3 + 2] = body[2];
            kind[i] = GIFT;
          }
          centre[i * 3] = baseX;
          centre[i * 3 + 1] = GROUND;
          centre[i * 3 + 2] = baseZ;
        }
        q.add(pos, nrm, col, kind, t0, 1, centre, 0, true);
      }
    }

    const P = q.freeze(rng);
    P.sky = sky_layout(rng, cols, rows);
    return P;
  }

  function sky_layout(rng, cols, rows) {
    const n = Math.floor((cols * rows) / 20);
    const f = Math.floor((cols * rows) / 60);
    const starsY = new Int32Array(n);
    const starsX = new Int32Array(n);
    const starsCh = new Int32Array(n);
    const starsPh = new Float32Array(n);
    const chars = [46, 46, 39, 46, 43, 42]; // ..'.+*
    for (let i = 0; i < n; i++) {
      starsY[i] = Math.floor(rng() * rows);
      starsX[i] = Math.floor(rng() * cols);
      starsCh[i] = chars[Math.floor(rng() * chars.length)];
      starsPh[i] = rng() * TAU;
    }
    const flakesY = new Float32Array(f);
    const flakesX = new Float32Array(f);
    const flakesSp = new Float32Array(f);
    const flakesPh = new Float32Array(f);
    const flakesNear = new Uint8Array(f);
    for (let i = 0; i < f; i++) {
      flakesY[i] = rng() * rows;
      flakesX[i] = rng() * cols;
      flakesSp[i] = 1.5 + rng() * 2.5;
      flakesPh[i] = rng() * TAU;
      flakesNear[i] = rng() < 0.35 ? 1 : 0;
    }
    return {
      starsY,
      starsX,
      starsCh,
      starsPh,
      flakesY,
      flakesX,
      flakesSp,
      flakesPh,
      flakesNear,
    };
  }

  // ---------- projection & drawing ----------
  function view(x, y, z, theta, out) {
    const c = Math.cos(theta);
    const s = Math.sin(theta);
    const rx = x * c - z * s;
    const rz = x * s + z * c;
    const ct = Math.cos(TILT);
    const st = Math.sin(TILT);
    out[0] = rx;
    out[1] = y * ct + rz * st;
    out[2] = rz * ct - y * st;
  }

  function project(x, y, z, theta, cols, rows, aspect, out) {
    const c = Math.cos(theta);
    const s = Math.sin(theta);
    const rx = x * c - z * s;
    const rz = x * s + z * c;
    const ct = Math.cos(TILT);
    const st = Math.sin(TILT);
    const vy = y * ct + rz * st;
    const vz = rz * ct - y * st;
    const f = CAM / (CAM + vz);
    out[0] = cols / 2 + rx * f * SCALE * aspect;
    out[1] = rows * HORIZON - vy * f * SCALE;
    out[2] = vz;
  }

  // ---------- frame generation ----------
  function frame(P, t, cols, rows, aspect) {
    const Ncells = cols * rows;
    const char = new Int32Array(Ncells);
    const colR = new Float32Array(Ncells);
    const colG = new Float32Array(Ncells);
    const colB = new Float32Array(Ncells);
    const emitR = new Float32Array(Ncells);
    const emitG = new Float32Array(Ncells);
    const emitB = new Float32Array(Ncells);

    const theta = SPIN * t;
    const info = { rotation: ((theta * DEG) % 360 + 360) % 360, lights: 0, letters: 0 };

    // back sky (stars)
    {
      const sky = P.sky;
      for (let i = 0; i < sky.starsY.length; i++) {
        const y = sky.starsY[i];
        const x = sky.starsX[i];
        if (y < 0 || y >= rows || x < 0 || x >= cols) continue;
        const idx = y * cols + x;
        const b = 0.18 + 0.5 * Math.pow(Math.sin(t * hz(1.3, LOOP) + sky.starsPh[i]), 8);
        char[idx] = sky.starsCh[i];
        colR[idx] = b * 0.8;
        colG[idx] = b * 0.85;
        colB[idx] = b * 1.0;
      }
    }

    // points
    const N = P.N;
    const tmp = [0, 0, 0];
    const depthBuf = new Float32Array(Ncells);
    depthBuf.fill(-1e9);

    let lightCount = 0;

    for (let i = 0; i < N; i++) {
      if (t < P.t0[i]) continue;
      const age = t - P.t0[i];
      let s = 1;
      if (P.pop[i]) s = pop(age / 0.35);

      const cx = P.centre[i * 3];
      const cy = P.centre[i * 3 + 1];
      const cz = P.centre[i * 3 + 2];
      const px = cx + (P.pos[i * 3] - cx) * s;
      const py = cy + (P.pos[i * 3 + 1] - cy) * s;
      const pz = cz + (P.pos[i * 3 + 2] - cz) * s;

      project(px, py, pz, theta, cols, rows, aspect, tmp);
      const sx = tmp[0];
      const sy = tmp[1];
      const depth = tmp[2];
      const ix = Math.floor(sx);
      const iy = Math.floor(sy);
      if (ix < 0 || ix >= cols || iy < 0 || iy >= rows) continue;

      // normal in view space
      view(P.nrm[i * 3], P.nrm[i * 3 + 1], P.nrm[i * 3 + 2], theta, tmp);
      const lam = clamp(tmp[0] * LIGHT[0] + tmp[1] * LIGHT[1] + tmp[2] * LIGHT[2], 0, 1);

      let kind = P.kind[i];
      let r = P.col[i * 3];
      let g = P.col[i * 3 + 1];
      let b = P.col[i * 3 + 2];
      let bright = (0.45 + 0.55 * lam) * P.gain[i];
      let glow = 0;

      // growth frost
      if (kind === NEEDLE && age < 0.25) {
        kind = SNOW;
        r = g = b = 0.95;
        bright = 0.9;
      }

      if (kind === BULB) {
        const tw = 0.5 + 0.5 * Math.sin(t * hz(4, LOOP) + P.phase[i]);
        bright = 0.35 + 0.65 * tw;
        r *= 0.55 + 0.6 * tw;
        g *= 0.55 + 0.6 * tw;
        b *= 0.55 + 0.6 * tw;
        glow = 0.35 * tw;
        lightCount++;
      } else if (kind === GARLAND) {
        const shimmer = 0.75 + 0.45 * Math.pow(Math.sin(t * hz(6, LOOP) + P.phase[i]), 2);
        r *= shimmer;
        g *= shimmer;
        b *= shimmer;
        glow = 0.12;
      } else if (kind === BALL) {
        const spec = Math.pow(lam, 14);
        r += spec * (1 - r);
        g += spec * (1 - g);
        b += spec * (1 - b);
        bright += spec;
        glow = 0.3 * spec;
      }

      const level = clamp(Math.floor((bright + 0.12 * P.jitter[i]) * LEVELS), 0, LEVELS - 1);
      const code = RAMP[kind * LEVELS + level];

      const idx = iy * cols + ix;
      // closer points win (larger depth is nearer in this convention? original used -depth sort far→near)
      // original: order = argsort(-depth) so larger depth first? Wait, project z increases away?
      // In original: f = CAM / (CAM + v[:,2]), so larger z is farther. They sort -depth so far first, then near overwrites.
      if (depth > depthBuf[idx]) {
        depthBuf[idx] = depth;
        char[idx] = code;
        colR[idx] = clamp(r * bright, 0, 1); // wait, rgb already has some brightness in original
        // Original: rgb = P.col * b, then level from b, but they write rgb (which includes b)
        // Looking back: rgb = P.col[vis] * b[:, None] then later char from level of b, col = rgb
        // So color already includes the brightness factor.
        colR[idx] = clamp(r * (kind === BULB || kind === GARLAND || kind === BALL ? 1 : bright / (0.45 + 0.55 * lam || 1)), 0, 1);
        // Simpler correct approach matching original more closely:
        // They compute b = (0.45+0.55*lam)*gain, rgb = col * b, then write rgb and use b for level.
        // I already have bright = b. For bulbs etc I overwrote bright and rgb.
        // Let's recompute cleanly.
      }
    }

    // Re-do the point loop more carefully to match original shading
    depthBuf.fill(-1e9);
    lightCount = 0;
    for (let i = 0; i < N; i++) {
      if (t < P.t0[i]) continue;
      const age = t - P.t0[i];
      let s = P.pop[i] ? pop(age / 0.35) : 1;

      const cx = P.centre[i * 3];
      const cy = P.centre[i * 3 + 1];
      const cz = P.centre[i * 3 + 2];
      const px = cx + (P.pos[i * 3] - cx) * s;
      const py = cy + (P.pos[i * 3 + 1] - cy) * s;
      const pz = cz + (P.pos[i * 3 + 2] - cz) * s;

      project(px, py, pz, theta, cols, rows, aspect, tmp);
      const ix = Math.floor(tmp[0]);
      const iy = Math.floor(tmp[1]);
      const depth = tmp[2];
      if (ix < 0 || ix >= cols || iy < 0 || iy >= rows) continue;

      view(P.nrm[i * 3], P.nrm[i * 3 + 1], P.nrm[i * 3 + 2], theta, tmp);
      const lam = clamp(tmp[0] * LIGHT[0] + tmp[1] * LIGHT[1] + tmp[2] * LIGHT[2], 0, 1);

      let kind = P.kind[i];
      let cr = P.col[i * 3];
      let cg = P.col[i * 3 + 1];
      let cb = P.col[i * 3 + 2];
      let b = (0.45 + 0.55 * lam) * P.gain[i];
      let glow = 0;

      if (kind === NEEDLE && age < 0.25) {
        kind = SNOW;
        cr = cg = cb = 0.95;
        b = 0.9;
      }

      if (kind === BULB) {
        const tw = 0.5 + 0.5 * Math.sin(t * hz(4, LOOP) + P.phase[i]);
        b = 0.35 + 0.65 * tw;
        cr = P.col[i * 3] * (0.55 + 0.6 * tw);
        cg = P.col[i * 3 + 1] * (0.55 + 0.6 * tw);
        cb = P.col[i * 3 + 2] * (0.55 + 0.6 * tw);
        glow = 0.35 * tw;
        lightCount++;
      } else if (kind === GARLAND) {
        const shimmer = 0.75 + 0.45 * Math.pow(Math.sin(t * hz(6, LOOP) + P.phase[i]), 2);
        cr *= shimmer;
        cg *= shimmer;
        cb *= shimmer;
        glow = 0.12;
      } else if (kind === BALL) {
        const spec = Math.pow(lam, 14);
        cr += spec * (1 - cr);
        cg += spec * (1 - cg);
        cb += spec * (1 - cb);
        b += spec;
        glow = 0.3 * spec;
      }

      const level = clamp(Math.floor((b + 0.12 * P.jitter[i]) * LEVELS), 0, LEVELS - 1);
      const code = RAMP[kind * LEVELS + level];

      const idx = iy * cols + ix;
      // near overwrites far: larger depth is farther in this projection, so we want smaller depth to win
      // Original sorted by -depth (far first), then wrote, so near (smaller? ) overwrites.
      // In their project: v[:,2] after view, f = CAM/(CAM+v2), larger v2 → smaller f → farther.
      // They argsort(-depth) → large depth (far) first, then small depth (near) overwrites. Perfect.
      // So condition for write: if this point is nearer than current, i.e. depth < depthBuf (if we store depth)
      // Initialize depthBuf to +inf, write if depth < depthBuf[idx]
      if (depth < depthBuf[idx] || depthBuf[idx] < -1e8) {
        depthBuf[idx] = depth;
        char[idx] = code;
        colR[idx] = clamp(cr, 0, 1); // original multiplies by b already for most, but for bulbs they set rgb = col*(..)
        // In original for normal: rgb = col * b
        // For bulb: rgb = col * (0.55+0.6*tw)
        // For garland: rgb = col * shimmer (no *b)
        // For ball: rgb = col*b + spec*(1-col*b) roughly
        // To match better:
        if (kind === BULB) {
          // already set
        } else if (kind === GARLAND) {
          // already * shimmer
        } else {
          colR[idx] = clamp(cr * b, 0, 1);
          colG[idx] = clamp(cg * b, 0, 1);
          colB[idx] = clamp(cb * b, 0, 1);
        }
        if (kind === BULB || kind === GARLAND || kind === BALL) {
          colR[idx] = clamp(cr, 0, 1);
          colG[idx] = clamp(cg, 0, 1);
          colB[idx] = clamp(cb, 0, 1);
        }
        emitR[idx] = colR[idx] * glow;
        emitG[idx] = colG[idx] * glow;
        emitB[idx] = colB[idx] * glow;
      }
    }

    info.lights = Math.floor(lightCount / BULB_PTS);

    // star
    if (t > T_STAR) {
      const on = smooth((t - T_STAR) / 0.3);
      const flash = Math.pow(clamp(1 - (t - T_STAR) / 1.0, 0, 1), 2);
      project(0, Y1 + 0.09, 0, theta, cols, rows, aspect, tmp);
      drawStar(char, colR, colG, colB, emitR, emitG, emitB, cols, rows, tmp[0], tmp[1], (1.6 + 1.6 * on) * (1 + 0.05 * Math.sin(t * hz(3, LOOP))), t, aspect, flash);
    }

    // front snowflakes
    {
      const sky = P.sky;
      for (let i = 0; i < sky.flakesY.length; i++) {
        if (!sky.flakesNear[i]) continue;
        let speed = sky.flakesSp[i] * 1.6;
        if (LOOP) speed = Math.max(1, Math.round((speed * LOOP) / rows)) * (rows / LOOP);
        const y = Math.floor((sky.flakesY[i] + t * speed) % rows);
        const x = Math.floor((sky.flakesX[i] + 1.5 * Math.sin(t * hz(0.8, LOOP) + sky.flakesPh[i])) % cols);
        if (y < 0 || y >= rows || x < 0 || x >= cols) continue;
        const idx = y * cols + x;
        char[idx] = 42; // *
        colR[idx] = colG[idx] = colB[idx] = 0.85;
      }
      // far flakes (behind, already somewhat drawn, but re-draw for correctness)
      for (let i = 0; i < sky.flakesY.length; i++) {
        if (sky.flakesNear[i]) continue;
        let speed = sky.flakesSp[i];
        if (LOOP) speed = Math.max(1, Math.round((speed * LOOP) / rows)) * (rows / LOOP);
        const y = Math.floor((sky.flakesY[i] + t * speed) % rows);
        const x = Math.floor((sky.flakesX[i] + 1.5 * Math.sin(t * hz(0.8, LOOP) + sky.flakesPh[i])) % cols);
        if (y < 0 || y >= rows || x < 0 || x >= cols) continue;
        const idx = y * cols + x;
        if (char[idx] === 0) {
          char[idx] = 46; // .
          colR[idx] = colG[idx] = colB[idx] = 0.5;
        }
      }
    }

    // message
    if (t > T_MSG) {
      const row = rows - 3;
      const x0 = Math.floor((cols - MESSAGE.length) / 2);
      const shown = Math.min(MESSAGE.length, Math.floor((t - T_MSG) * 10));
      // clear band
      for (let dy = -1; dy <= 1; dy++) {
        for (let dx = -3; dx < shown + 3; dx++) {
          const xx = x0 + dx;
          const yy = row + dy;
          if (yy >= 0 && yy < rows && xx >= 0 && xx < cols) {
            const idx = yy * cols + xx;
            char[idx] = 0;
            colR[idx] = colG[idx] = colB[idx] = 0;
            emitR[idx] = emitG[idx] = emitB[idx] = 0;
          }
        }
      }
      for (let i = 0; i < shown; i++) {
        const ch = MESSAGE[i];
        if (ch === " ") continue;
        const tint = [RED, GOLD, GREEN_BOX][(Math.floor(i / 2) + Math.floor(t * 1.5)) % 3];
        const power = 0.85 + 0.25 * Math.sin(t * hz(5, LOOP) + i);
        const xx = x0 + i;
        if (xx >= 0 && xx < cols && row >= 0 && row < rows) {
          const idx = row * cols + xx;
          char[idx] = ch.charCodeAt(0);
          colR[idx] = clamp(tint[0] * power, 0, 1);
          colG[idx] = clamp(tint[1] * power, 0, 1);
          colB[idx] = clamp(tint[2] * power, 0, 1);
          emitR[idx] = colR[idx] * 0.4;
          emitG[idx] = colG[idx] * 0.4;
          emitB[idx] = colB[idx] * 0.4;
        }
      }
    }

    let letters = 0;
    for (let i = 0; i < Ncells; i++) if (char[i] > 0) letters++;
    info.letters = letters;

    return { char, colR, colG, colB, emitR, emitG, emitB, info };
  }

  function drawStar(char, colR, colG, colB, emitR, emitG, emitB, cols, rows, cx, cy, size, t, aspect, flash) {
    const tintR = GOLD[0] + (1 - GOLD[0]) * flash;
    const tintG = GOLD[1] + (1 - GOLD[1]) * flash;
    const tintB = GOLD[2] + (1 - GOLD[2]) * flash;
    const power = 1 + 1.6 * flash;

    // rays
    for (let k = 0; k < 8; k++) {
      const ang = (k * TAU) / 8 + 0.1 * Math.sin(t * hz(0.7, LOOP));
      if (Math.sin(ang) < -0.3) continue;
      const reach = size * (3.0 + 0.8 * Math.sin(t * hz(3, LOOP) + k * 1.3)) * (k % 2 === 0 ? 1.0 : 0.6) * (1 + flash);
      const chs = "-/|\\";
      const ch = chs[Math.round((((ang * DEG) % 180) + 180) % 180 / 45) % 4];
      for (let s = size * 1.1; s < reach; s += 0.5) {
        const y = Math.round(cy - Math.sin(ang) * s);
        const x = Math.round(cx + Math.cos(ang) * s * aspect);
        if (y < 0 || y >= rows || x < 0 || x >= cols) continue;
        const idx = y * cols + x;
        const p = power * 0.6 * (1 - s / reach);
        char[idx] = ch.charCodeAt(0);
        colR[idx] = clamp(tintR * p, 0, 1);
        colG[idx] = clamp(tintG * p, 0, 1);
        colB[idx] = clamp(tintB * p, 0, 1);
        emitR[idx] = colR[idx] * 0.5;
        emitG[idx] = colG[idx] * 0.5;
        emitB[idx] = colB[idx] * 0.5;
      }
    }

    // filled star
    const R = size;
    const rIn = size * 0.45;
    const spin = 0.15 * Math.sin(t * hz(0.9, LOOP));
    const y0 = Math.floor(cy - R) - 1;
    const y1 = Math.floor(cy + R) + 2;
    const x0 = Math.floor(cx - R * aspect) - 1;
    const x1 = Math.floor(cx + R * aspect) + 2;
    for (let y = y0; y < y1; y++) {
      for (let x = x0; x < x1; x++) {
        if (y < 0 || y >= rows || x < 0 || x >= cols) continue;
        const dx = (x - cx) / aspect;
        const dy = y - cy;
        const r = Math.hypot(dx, dy);
        let a = (Math.atan2(dx, -dy) - spin) % (TAU / 5);
        if (a < 0) a += TAU / 5;
        a /= TAU / 5;
        if (r <= rIn + (R - rIn) * Math.abs(a - 0.5) * 2) {
          const idx = y * cols + x;
          const ch = r < R * 0.35 ? 64 : r < R * 0.65 ? 35 : 42; // @ # *
          const p = power * (1.15 - 0.3 * r / R);
          char[idx] = ch;
          colR[idx] = clamp(tintR * p, 0, 1);
          colG[idx] = clamp(tintG * p, 0, 1);
          colB[idx] = clamp(tintB * p, 0, 1);
          emitR[idx] = colR[idx] * 0.9;
          emitG[idx] = colG[idx] * 0.9;
          emitB[idx] = colB[idx] * 0.9;
        }
      }
    }
  }

  // ---------- Canvas renderer ----------
  const CW = 9;
  const CH = 15;
  const FONT_PX = 14;
  const COLS = 88;
  const ROWS = 74;
  const BG = [6 / 255, 8 / 255, 14 / 255];

  class Renderer {
    constructor(container) {
      this.cols = COLS;
      this.rows = ROWS;
      this.cw = CW;
      this.ch = CH;
      this.w = COLS * CW;
      this.h = ROWS * CH;
      this.hudH = 0; // no HUD when used as background

      this.canvas = document.createElement("canvas");
      this.canvas.width = this.w;
      this.canvas.height = this.h;
      // Full-screen fixed background, shifted to the right
      this.canvas.id = "ascii-christmas-tree";
      this.canvas.style.cssText = [
        "position:fixed",
        "top:5%",
        "left:0",
        "width:100vw",
        "height:95%",
        "object-fit:contain",
        "object-position: right center",
        "background:transparent",
        "z-index:0",
        "pointer-events:none",
        "image-rendering:pixelated"
      ].join(";");
      this.ctx = this.canvas.getContext("2d", { alpha: true });

      // glyph atlas
      const atlasCanvas = document.createElement("canvas");
      atlasCanvas.width = 95 * CW;
      atlasCanvas.height = CH;
      const actx = atlasCanvas.getContext("2d");
      actx.fillStyle = "#000";
      actx.fillRect(0, 0, atlasCanvas.width, atlasCanvas.height);
      actx.font = `${FONT_PX}px "Menlo","Consolas","DejaVu Sans Mono","Courier New",monospace`;
      actx.textAlign = "center";
      actx.textBaseline = "middle";
      actx.fillStyle = "#fff";
      for (let code = 32; code < 127; code++) {
        actx.fillText(String.fromCharCode(code), (code - 32) * CW + CW / 2, CH / 2);
      }
      this.atlas = atlasCanvas;

      // offscreen for the art
      this.art = document.createElement("canvas");
      this.art.width = this.w;
      this.art.height = this.h;
      this.artCtx = this.art.getContext("2d", { alpha: true });

      // bloom temp
      this.bloom = document.createElement("canvas");
      this.bloom.width = this.cols;
      this.bloom.height = this.rows;
      this.bloomCtx = this.bloom.getContext("2d", { alpha: true });

      // Put canvas behind page content
      if (container) {
        container.innerHTML = "";
        container.appendChild(this.canvas);
      } else {
        document.body.prepend(this.canvas);
      }
      // Only z-index — do NOT touch position of fixed navbar
      if (!document.getElementById("christmas-tree-layer-fix")) {
        const style = document.createElement("style");
        style.id = "christmas-tree-layer-fix";
        style.textContent = [
          "#ascii-christmas-tree{",
          "position:fixed!important;top:5%!important;left:0;width:100vw;height:95%!important;",
          "z-index:0!important;pointer-events:none!important;background:transparent!important;",
          "}",
          "#mainNav,.navbar.fixed-top{z-index:1030!important;}",
          "#main-content,main{position:relative;z-index:1;}",
        ].join("");
        document.head.appendChild(style);
      }
    }

    draw(data, t, paused) {
      const { char, colR, colG, colB, emitR, emitG, emitB, info } = data;
      const ctx = this.artCtx;
      ctx.clearRect(0, 0, this.w, this.h);
      const img = ctx.createImageData(this.w, this.h);
      const px = img.data;

      // fully transparent base
      for (let i = 0; i < px.length; i += 4) {
        px[i] = 0;
        px[i + 1] = 0;
        px[i + 2] = 0;
        px[i + 3] = 0;
      }

      // draw glyphs (simple nearest: sample atlas)
      // For speed we approximate by filling a colored rect modulated by a crude glyph mask,
      // or we can use drawImage per cell (slow). Better: use the atlas with globalComposite.
      // Fast path: for each cell, if char, fill a solid color rect (loses glyph shape but fast),
      // then we can improve later. For fidelity we do per-pixel from atlas.

      // Pre-extract atlas luminance
      if (!this.atlasData) {
        const actx = this.atlas.getContext("2d");
        this.atlasData = actx.getImageData(0, 0, 95 * CW, CH).data;
      }
      const ad = this.atlasData;

      // 2) colored glyph on top of outline
      for (let row = 0; row < this.rows; row++) {
        for (let col = 0; col < this.cols; col++) {
          const idx = row * this.cols + col;
          const code = char[idx];
          if (code < 32 || code > 126) continue;
          const gi = code - 32;
          const cr = colR[idx];
          const cg = colG[idx];
          const cb = colB[idx];
          for (let dy = 0; dy < CH; dy++) {
            for (let dx = 0; dx < CW; dx++) {
              const aidx = (dy * (95 * CW) + gi * CW + dx) * 4;
              const lum = ad[aidx] / 255;
              if (lum < 0.08) continue;
              const pidx = ((row * CH + dy) * this.w + (col * CW + dx)) * 4;
              px[pidx]     = Math.min(255, cr * 255 * lum);
              px[pidx + 1] = Math.min(255, cg * 255 * lum);
              px[pidx + 2] = Math.min(255, cb * 255 * lum);
              px[pidx + 3] = Math.min(255, lum * 255);
            }
          }
        }
      }
      ctx.putImageData(img, 0, 0);

      // simple bloom
      const bctx = this.bloomCtx;
      bctx.clearRect(0, 0, this.cols, this.rows);
      const bimg = bctx.createImageData(this.cols, this.rows);
      const bp = bimg.data;
      for (let i = 0; i < this.cols * this.rows; i++) {
        const er = emitR[i], eg = emitG[i], eb = emitB[i];
        const a = Math.max(er, eg, eb);
        if (a < 0.01) {
          bp[i * 4 + 3] = 0;
          continue;
        }
        bp[i * 4]     = Math.min(255, er * 255 * 1.5);
        bp[i * 4 + 1] = Math.min(255, eg * 255 * 1.5);
        bp[i * 4 + 2] = Math.min(255, eb * 255 * 1.5);
        bp[i * 4 + 3] = Math.min(255, a * 200);
      }
      bctx.putImageData(bimg, 0, 0);
      bctx.filter = "blur(1.5px)";
      bctx.drawImage(this.bloom, 0, 0);
      bctx.filter = "none";

      // final composite – fully transparent where nothing is drawn
      this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
      this.ctx.drawImage(this.art, 0, 0);
      this.ctx.globalCompositeOperation = "lighter";
      this.ctx.drawImage(this.bloom, 0, 0, this.cols, this.rows, 0, 0, this.w, this.h);
      this.ctx.globalCompositeOperation = "source-over";
	  
	  // Кириллица — атлас только ASCII
	const mctx = this.ctx;
	mctx.font = `${FONT_PX}px "Menlo","Consolas","DejaVu Sans Mono","Courier New",monospace`;
	mctx.textAlign = "center";
	mctx.textBaseline = "middle";
	mctx.shadowColor = "transparent";
	mctx.shadowBlur = 0;

	for (let row = 0; row < this.rows; row++) {
	  for (let col = 0; col < this.cols; col++) {
		const idx = row * this.cols + col;
		const code = char[idx];
		if (code <= 126) continue; // ASCII уже нарисован атласом
		const r = Math.round(colR[idx] * 255);
		const g = Math.round(colG[idx] * 255);
		const b = Math.round(colB[idx] * 255);
		mctx.fillStyle = `rgb(${r},${g},${b})`;
		mctx.fillText(
		  String.fromCharCode(code),
		  col * this.cw + this.cw / 2,
		  row * this.ch + this.ch / 2
		);
	  }
	}
	  
    }
  }

  // ---------- main ----------
  // Tree is fully decorated after ~14s; after that it only spins.
  const T_FULL = 14.0;
  // sessionStorage = once per browser session (after login / new tab session)
  const STORAGE_KEY = "christmas_tree_seen_2026";

  function start() {
	  
	  
	  // ==== ДАТА/МЕСЯЦ ПОКАЗА ====
	  const now = new Date();
	  const m = now.getMonth() + 1;   // 1..12
	  const d = now.getDate();        // 1..31

	  const START_MONTH = 12, START_DAY = 1;   // с 1 декабря
	  const END_MONTH   = 1,  END_DAY   = 14;  // по 14 января

	  // сезон = декабрь ИЛИ первая половина января
	  const inSeason =
		(m === START_MONTH && d >= START_DAY) ||
		(m === END_MONTH   && d <= END_DAY);

	  if (!inSeason) return;   // вне сезона — ёлку не строим, канвас не создаём
	  // ============================
  
    const host = document.getElementById("christmas-tree");
    const R = new Renderer(host);
    const P = build(COLS, ROWS);
    const aspect = CH / CW;

    let seen = false;
    try {
      seen = sessionStorage.getItem(STORAGE_KEY) === "1";
    } catch (e) {}

    // First page in session → growth from 0.
    // Mark immediately so any next page load shows the full tree.
    let t = seen ? T_FULL : 0;
    if (!seen) {
      try {
        sessionStorage.setItem(STORAGE_KEY, "1");
      } catch (e) {}
    }

    let last = performance.now();

    function tick(now) {
      const dt = Math.min(0.05, (now - last) / 1000);
      last = now;
      t += dt;
      const data = frame(P, t, COLS, ROWS, aspect);
      R.draw(data, t, false);
      requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start);
  } else {
    start();
  }
})();
