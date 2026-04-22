'use strict';

(function () {
  const _ = null;
  const K = '#2a2a3a';   // 외곽선
  const G = '#7a7e8e';   // 회색 몸통
  const L = '#a8acbe';   // 밝은 회색 하이라이트
  const W = '#e8e8f0';   // 얼굴/배 흰색
  const R = '#cc2222';   // 붉은 목걸이
  const B = '#d4a520';   // 금 방울
  const E = '#1a1a2a';   // 눈/코

  // 앉기/기본 포즈
  const SIT = [
    [_,_,K,K,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,G,G,K,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,L,G,G,K,_,_,_,_,_,_,_,_,_,_],
    [K,G,W,E,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,W,W,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,G,G,G,G,K,_,_,_,_,_,K,_,_,_],
    [_,K,R,R,R,R,K,_,_,_,_,K,G,K,_,_],
    [_,K,G,G,B,G,G,G,G,G,K,G,G,K,_,_],
    [_,K,G,L,G,G,G,G,G,K,_,K,G,K,_,_],
    [_,_,K,G,G,G,G,G,K,_,_,_,K,K,_,_],
    [_,_,K,G,G,G,G,K,_,_,_,_,_,_,_,_],
    [_,_,_,K,K,_,K,K,_,_,_,_,_,_,_,_],
    [_,_,_,K,_,_,K,_,_,_,_,_,_,_,_,_],
    [_,_,K,K,_,_,K,K,_,_,_,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
  ];

  // 걷기 프레임 A (앞다리 앞으로)
  const WALK_A = [
    [_,_,K,K,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,G,G,K,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,L,G,G,K,_,_,_,_,_,_,_,_,_,_],
    [K,G,W,E,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,W,W,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,G,G,G,G,K,_,_,_,_,_,K,_,_,_],
    [_,K,R,R,R,R,K,_,_,_,_,K,G,K,_,_],
    [_,K,G,G,B,G,G,G,G,G,K,G,G,K,_,_],
    [_,K,G,L,G,G,G,G,G,K,_,K,G,K,_,_],
    [_,_,K,G,G,G,G,G,K,_,_,_,K,K,_,_],
    [_,_,K,G,G,G,G,K,_,_,_,_,_,_,_,_],
    [_,K,K,_,_,_,K,K,_,_,_,_,_,_,_,_],
    [K,K,_,_,_,_,K,_,_,_,_,_,_,_,_,_],
    [K,_,_,_,_,K,K,_,_,_,_,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
  ];

  // 걷기 프레임 B (앞다리 뒤로)
  const WALK_B = [
    [_,_,K,K,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,G,G,K,_,_,_,_,_,_,_,_,_,_,_],
    [_,K,L,G,G,K,_,_,_,_,_,_,_,_,_,_],
    [K,G,W,E,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,W,W,G,G,K,_,_,_,_,_,_,_,_,_],
    [K,G,G,G,G,G,K,_,_,_,_,_,K,_,_,_],
    [_,K,R,R,R,R,K,_,_,_,_,K,G,K,_,_],
    [_,K,G,G,B,G,G,G,G,G,K,G,G,K,_,_],
    [_,K,G,L,G,G,G,G,G,K,_,K,G,K,_,_],
    [_,_,K,G,G,G,G,G,K,_,_,_,K,K,_,_],
    [_,_,K,G,G,G,G,K,_,_,_,_,_,_,_,_],
    [_,_,_,K,K,_,_,K,K,_,_,_,_,_,_,_],
    [_,_,_,K,_,_,_,_,K,K,_,_,_,_,_,_],
    [_,_,K,K,_,_,_,_,_,K,K,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
    [_,_,_,_,_,_,_,_,_,_,_,_,_,_,_,_],
  ];

  const SCALE = 2;
  const SPR = 16;
  const PX = SPR * SCALE;
  const BUBBLE_H = 22;
  const PAD = 4;
  const CW = PX + PAD * 2;
  const CH = PX + BUBBLE_H;
  const MARGIN_X = 12;
  const FROM_BOTTOM = 10;
  const VERT_JITTER = 28;

  const PURRS = ['그르릉~', '골골~', '냥…'];

  const canvas = document.createElement('canvas');
  canvas.width = CW;
  canvas.height = CH;
  canvas.setAttribute('aria-hidden', 'true');
  Object.assign(canvas.style, {
    position: 'fixed',
    zIndex: '40',
    pointerEvents: 'none',
    imageRendering: 'pixelated',
    left: '0',
    top: '0',
  });
  document.body.appendChild(canvas);
  const ctx = canvas.getContext('2d');

  const hit = document.createElement('div');
  hit.setAttribute('aria-hidden', 'true');
  Object.assign(hit.style, {
    position: 'fixed',
    width: PX + 'px',
    height: PX + 'px',
    zIndex: '41',
    cursor: 'default',
  });
  document.body.appendChild(hit);

  let cx = 40;
  let vx = 0;
  let vy = 0;
  let tx = 100;
  let ty = 0;
  let state = 'wander';
  let sitTimer = 0;
  let hovered = false;
  let hoverPurr = PURRS[0];
  let bobClock = 0;
  let walkClock = 0;
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
    const tw = ctx.measureText(text).width;
    const bw = tw + 14;
    const bh = 18;
    const bx = CW / 2 - bw / 2;
    const by = 2;
    const tip = CW / 2;

    ctx.shadowColor = 'rgba(0,0,0,0.12)';
    ctx.shadowBlur = 4;
    ctx.fillStyle = '#fffdf8';
    rrect(bx, by, bw, bh, 5);
    ctx.fill();
    ctx.restore();

    ctx.strokeStyle = '#d8cfc0';
    ctx.lineWidth = 1;
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
    if (state === 'sit' || hovered) {
      frame = SIT;
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
    const top = Math.round(cyVal + bob);
    canvas.style.left = left + 'px';
    canvas.style.top = top + 'px';
    hit.style.left = left + PAD + 'px';
    hit.style.top = top + BUBBLE_H + 'px';
  }

  function tick(now) {
    const dt = Math.min((now - lastT) / 1000, 0.05);
    lastT = now;

    const { minCy, maxCy } = bandY();

    if (state === 'sit') {
      vx *= 0.85;
      vy *= 0.85;
      sitTimer -= dt;
      if (sitTimer <= 0) {
        state = 'wander';
        newTarget();
      }
    } else {
      const dx = tx - cx;
      const dy = ty - cy;
      const dist = Math.hypot(dx, dy) || 1;
      if (dist < 6) {
        state = 'sit';
        sitTimer = 0.8 + Math.random() * 1.2;
        vx = 0;
        vy = 0;
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
    bobClock += dt * (moving ? 7 : 2);
    walkClock += dt;

    if (Math.abs(vx) > 1) {
      facingLeft = vx < 0;
    }

    const bob = Math.sin(bobClock) * (moving ? 2 : 0.6);
    syncLayout(cy, bob);
    render();
    requestAnimationFrame(tick);
  }

  hit.addEventListener('mouseenter', () => {
    hovered = true;
    hoverPurr = PURRS[(Math.random() * PURRS.length) | 0];
  });
  hit.addEventListener('mouseleave', () => {
    hovered = false;
  });

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
