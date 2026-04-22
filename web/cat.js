'use strict';

(function () {
  const _ = null;
  const K = '#3d1808';  // 가장 어두운 외곽선
  const D = '#6b3010';  // 진한 갈색
  const Y = '#f0b000';  // 바나나 노란색
  const H = '#ffd84d';  // 밝은 노란색 (하이라이트)
  const O = '#c07010';  // 주황갈색 (그림자)
  const F = '#7a5840';  // 고양이 털
  const G = '#b0a8a0';  // 회색 주둥이
  const W = '#d8d0cc';  // 밝은 주둥이
  const L = '#9a7860';  // 밝은 갈색 털
  const S = '#c0c0d8';  // 발 아래 그림자
  const E = '#2a1408';  // 눈

  // 몸통 (모든 프레임 공통, row 0-8 + 13-19)
  const BODY = [
    [_,_,_,_,_,_,_,K,K,_,_,_,_,_,_,_],   //  0 바나나 끝
    [_,_,_,_,_,_,K,D,D,K,_,_,_,_,_,_],   //  1
    [_,_,_,_,_,_,K,D,D,K,_,_,_,_,_,_],   //  2
    [_,_,_,_,_,K,D,D,D,K,_,_,_,_,_,_],   //  3
    [_,_,_,_,K,Y,O,D,K,_,_,_,_,_,_,_],   //  4 몸통 시작
    [_,_,_,K,Y,H,Y,O,D,K,_,_,_,_,_,_],   //  5
    [_,_,K,Y,H,H,Y,Y,O,D,K,_,_,_,_,_],   //  6
    [_,K,Y,H,H,Y,Y,Y,Y,O,D,K,_,_,_,_],   //  7
    [K,Y,H,Y,F,F,F,F,F,Y,Y,O,K,_,_,_],   //  8 얼굴 시작
    // rows 9-12: 눈 (프레임별로 교체)
    [K,Y,Y,F,L,G,G,G,L,F,Y,Y,O,K,_,_],   //  9
    [K,Y,Y,F,E,E,G,E,E,F,Y,Y,O,K,_,_],   // 10 눈
    [K,Y,Y,F,G,W,G,W,G,F,Y,Y,O,K,_,_],   // 11
    [K,Y,Y,F,F,G,G,G,F,F,Y,Y,O,K,_,_],   // 12
    [K,Y,Y,Y,F,F,F,F,F,Y,Y,Y,O,K,_,_],   // 13 턱
    [K,Y,Y,Y,Y,Y,Y,Y,Y,Y,Y,Y,O,K,_,_],   // 14 몸통
    [K,F,Y,Y,Y,Y,Y,Y,Y,Y,Y,F,O,K,_,_],   // 15 팔 위치
    [_,K,F,O,Y,Y,Y,Y,Y,O,F,K,_,_,_,_],   // 16 팔 아래
    [_,_,K,O,D,Y,Y,Y,Y,D,K,_,_,_,_,_],   // 17 하단 몸통
    [_,_,K,Y,Y,Y,Y,Y,Y,Y,K,_,_,_,_,_],   // 18
    [_,_,_,K,Y,Y,Y,Y,Y,K,_,_,_,_,_,_],   // 19
  ];

  // 기본 눈 (슬픈 눈)
  const EYES_NORMAL = BODY.slice(9, 13);

  // 호버 시 눈 (살짝 덜 슬픈)
  const EYES_HAPPY = [
    [K,Y,Y,F,L,G,G,G,L,F,Y,Y,O,K,_,_],
    [K,Y,Y,F,G,E,W,E,G,F,Y,Y,O,K,_,_],
    [K,Y,Y,F,G,G,G,G,G,F,Y,Y,O,K,_,_],
    [K,Y,Y,F,F,G,G,G,F,F,Y,Y,O,K,_,_],
  ];

  // 기본 다리 (서 있는 포즈)
  const LEGS_SIT = [
    [_,_,_,K,F,Y,F,F,Y,F,K,_,_,_,_,_],
    [_,_,_,K,F,F,K,K,F,F,K,_,_,_,_,_],
    [_,_,S,S,S,S,S,S,S,S,S,_,_,_,_,_],
    [_,_,_,S,S,S,S,S,S,_,_,_,_,_,_,_],
  ];

  // 걷기 A - 다리 벌린 자세
  const LEGS_A = [
    [_,_,K,F,Y,F,_,_,F,Y,F,K,_,_,_,_],
    [_,K,F,F,_,_,_,_,_,_,F,F,K,_,_,_],
    [_,S,S,S,_,_,_,_,_,_,S,S,S,_,_,_],
    [_,_,S,S,_,_,_,_,_,_,S,S,_,_,_,_],
  ];

  // 걷기 B - 다리 모은 자세
  const LEGS_B = [
    [_,_,_,_,K,F,F,F,F,K,_,_,_,_,_,_],
    [_,_,_,_,K,F,K,K,F,K,_,_,_,_,_,_],
    [_,_,_,S,S,S,S,S,S,S,_,_,_,_,_,_],
    [_,_,_,_,S,S,S,S,_,_,_,_,_,_,_,_],
  ];

  function makeFrame(eyes, legs) {
    return [...BODY.slice(0, 9), ...eyes, ...BODY.slice(13), ...legs];
  }

  const NORMAL = makeFrame(EYES_NORMAL, LEGS_SIT);
  const HAPPY  = makeFrame(EYES_HAPPY,  LEGS_SIT);
  const WALK_A = makeFrame(EYES_NORMAL, LEGS_A);
  const WALK_B = makeFrame(EYES_NORMAL, LEGS_B);

  const SCALE   = 3;
  const SPR_W   = 16;
  const SPR_H   = 24;
  const PX_W    = SPR_W * SCALE;
  const PX_H    = SPR_H * SCALE;
  const BUBBLE_H = 22;
  const PAD     = 4;
  const CW      = PX_W + PAD * 2;
  const CH      = PX_H + BUBBLE_H;
  const MARGIN_X  = 12;
  const FROM_BOTTOM = 10;
  const VERT_JITTER = 28;

  const PURRS = ['으아아…', '바나나~', '냥…'];

  const canvas = document.createElement('canvas');
  canvas.width  = CW;
  canvas.height = CH;
  canvas.setAttribute('aria-hidden', 'true');
  Object.assign(canvas.style, {
    position: 'fixed',
    zIndex: '40',
    pointerEvents: 'none',
    imageRendering: 'pixelated',
    left: '0',
    top:  '0',
  });
  document.body.appendChild(canvas);
  const ctx = canvas.getContext('2d');

  const hit = document.createElement('div');
  hit.setAttribute('aria-hidden', 'true');
  Object.assign(hit.style, {
    position: 'fixed',
    width:  PX_W + 'px',
    height: PX_H + 'px',
    zIndex: '41',
    cursor: 'default',
  });
  document.body.appendChild(hit);

  let cx = 40, vx = 0, vy = 0;
  let tx = 100, ty = 0;
  let state = 'wander', sitTimer = 0;
  let hovered = false, hoverPurr = PURRS[0];
  let bobClock = 0, walkClock = 0;
  let facingLeft = false;
  let lastT = 0;

  function bandY() {
    const maxCy = window.innerHeight - CH - FROM_BOTTOM;
    const minCy = Math.max(MARGIN_X, maxCy - VERT_JITTER);
    return { minCy, maxCy };
  }

  function newTarget() {
    const { minCy, maxCy } = bandY();
    tx = MARGIN_X + Math.random() * Math.max(8, window.innerWidth - CW - MARGIN_X * 2);
    ty = minCy + Math.random() * Math.max(0, maxCy - minCy);
  }

  let cy = bandY().maxCy;

  function rrect(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function drawBubble(text) {
    ctx.save();
    ctx.font = '600 10px system-ui, "Segoe UI", sans-serif';
    const tw  = ctx.measureText(text).width;
    const bw  = tw + 14, bh = 18;
    const bx  = CW / 2 - bw / 2, by = 2;
    const tip = CW / 2;
    ctx.shadowColor = 'rgba(0,0,0,0.12)';
    ctx.shadowBlur  = 4;
    ctx.fillStyle   = '#fffdf8';
    rrect(bx, by, bw, bh, 5);
    ctx.fill();
    ctx.restore();
    ctx.strokeStyle = '#d8cfc0';
    ctx.lineWidth   = 1;
    rrect(bx, by, bw, bh, 5);
    ctx.stroke();
    ctx.fillStyle = '#fffdf8';
    ctx.beginPath();
    ctx.moveTo(tip - 4, by + bh - 1);
    ctx.lineTo(tip + 4, by + bh - 1);
    ctx.lineTo(tip, by + bh + 5);
    ctx.closePath();
    ctx.fill();
    ctx.strokeStyle = '#d8cfc0';
    ctx.beginPath();
    ctx.moveTo(tip - 4, by + bh);
    ctx.lineTo(tip, by + bh + 5);
    ctx.lineTo(tip + 4, by + bh);
    ctx.stroke();
    ctx.fillStyle = '#3a2a1a';
    ctx.fillText(text, bx + 7, by + bh - 5);
  }

  function drawSprite(frame) {
    for (let r = 0; r < frame.length; r++) {
      for (let c = 0; c < frame[r].length; c++) {
        const col = frame[r][c];
        if (!col) continue;
        ctx.fillStyle = col;
        ctx.fillRect(PAD + c * SCALE, BUBBLE_H + r * SCALE, SCALE, SCALE);
      }
    }
  }

  function render() {
    ctx.clearRect(0, 0, CW, CH);
    if (hovered) drawBubble(hoverPurr);

    let frame;
    if (hovered) {
      frame = HAPPY;
    } else if (state === 'sit') {
      frame = NORMAL;
    } else {
      frame = (Math.floor(walkClock / 0.2) % 2 === 0) ? WALK_A : WALK_B;
    }

    if (facingLeft) {
      ctx.save();
      ctx.translate(CW, 0);
      ctx.scale(-1, 1);
      drawSprite(frame);
      ctx.restore();
    } else {
      drawSprite(frame);
    }
  }

  function syncLayout(cyVal, bob) {
    const left = Math.round(cx);
    const top  = Math.round(cyVal + bob);
    canvas.style.left = left + 'px';
    canvas.style.top  = top  + 'px';
    hit.style.left    = left + PAD + 'px';
    hit.style.top     = top  + BUBBLE_H + 'px';
  }

  function tick(now) {
    const dt = Math.min((now - lastT) / 1000, 0.05);
    lastT = now;
    const { minCy, maxCy } = bandY();

    if (state === 'sit') {
      vx *= 0.85; vy *= 0.85;
      sitTimer -= dt;
      if (sitTimer <= 0) { state = 'wander'; newTarget(); }
    } else {
      const dx = tx - cx, dy = ty - cy;
      const dist = Math.hypot(dx, dy) || 1;
      if (dist < 6) {
        state = 'sit';
        sitTimer = 0.8 + Math.random() * 1.2;
        vx = vy = 0;
      } else {
        const spd = 38;
        vx += ((dx / dist) * spd - vx) * dt * 5;
        vy += ((dy / dist) * spd - vy) * dt * 5;
      }
    }

    cx += vx * dt;
    cy += vy * dt;
    cx = Math.max(MARGIN_X, Math.min(window.innerWidth - CW - MARGIN_X, cx));
    cy = Math.max(minCy, Math.min(maxCy, cy));

    const moving = Math.abs(vx) > 1.5 || Math.abs(vy) > 1.5;
    bobClock  += dt * (moving ? 7 : 2);
    walkClock += dt;
    if (Math.abs(vx) > 1) facingLeft = vx < 0;

    const bob = Math.sin(bobClock) * (moving ? 2 : 0.6);
    syncLayout(cy, bob);
    render();
    requestAnimationFrame(tick);
  }

  hit.addEventListener('mouseenter', () => {
    hovered   = true;
    hoverPurr = PURRS[(Math.random() * PURRS.length) | 0];
  });
  hit.addEventListener('mouseleave', () => { hovered = false; });
  window.addEventListener('resize', () => {
    const { minCy, maxCy } = bandY();
    cy = Math.min(maxCy, Math.max(minCy, cy));
    cx = Math.min(window.innerWidth - CW - MARGIN_X, Math.max(MARGIN_X, cx));
  });

  newTarget();
  lastT = performance.now();
  syncLayout(cy, 0);
  render();
  requestAnimationFrame(tick);
})();
