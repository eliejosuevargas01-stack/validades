#!/usr/bin/env node
'use strict';

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const UI_PATH = path.join(__dirname, 'loadtest-ui.html');

let sseClients = new Set();
let current = null;

function sendSse(event, data) {
  const payload = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const res of sseClients) {
    res.write(payload);
  }
}

function safeJson(req) {
  return new Promise((resolve) => {
    let body = '';
    req.on('data', (chunk) => { body += chunk; });
    req.on('end', () => {
      try { resolve(body ? JSON.parse(body) : {}); }
      catch (_) { resolve({}); }
    });
  });
}

function createRunner(opts) {
  const target = new URL(opts.url);
  const agent = target.protocol === 'https:'
    ? new https.Agent({ keepAlive: true, maxSockets: opts.concurrency * 4 })
    : new http.Agent({ keepAlive: true, maxSockets: opts.concurrency * 4 });

  let sent = 0;
  let completed = 0;
  let ok = 0;
  let err = 0;
  let inflight = 0;
  let stopped = false;
  const latencies = [];
  const startedAt = Date.now();

  function requestOnce(seq) {
    return new Promise((resolve) => {
      const body = JSON.stringify({ action: opts.action });
      const started = Date.now();
      const options = {
        protocol: target.protocol,
        hostname: target.hostname,
        port: target.port || (target.protocol === 'https:' ? 443 : 80),
        path: target.pathname + target.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Content-Length': Buffer.byteLength(body)
        },
        timeout: opts.timeoutMs,
        agent
      };

      const transport = target.protocol === 'https:' ? https : http;
      const req = transport.request(options, (res) => {
        res.on('data', () => {});
        res.on('end', () => {
          const ms = Date.now() - started;
          latencies.push(ms);
          if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) ok += 1;
          else err += 1;
          resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, ms, seq });
        });
      });

      req.on('timeout', () => {
        req.destroy(new Error('timeout'));
      });

      req.on('error', () => {
        const ms = Date.now() - started;
        latencies.push(ms);
        err += 1;
        resolve({ ok: false, status: 'ERR', ms, seq });
      });

      req.write(body);
      req.end();
    });
  }

  async function run() {
    const queue = Array.from({ length: opts.total }, (_, i) => i + 1);
    sendSse('log', { message: `Iniciando: ${opts.total} reqs | conc ${opts.concurrency} | timeout ${opts.timeoutMs}ms` });

    const workers = Array.from({ length: opts.concurrency }, async () => {
      while (queue.length && !stopped) {
        const seq = queue.shift();
        sent += 1;
        inflight += 1;
        const res = await requestOnce(seq);
        inflight -= 1;
        completed += 1;
        if (!res.ok) {
          sendSse('log', { message: `#${res.seq} falhou (${res.status}) em ${res.ms}ms` });
        }
        if (completed % 50 === 0 || completed === opts.total) {
          sendSse('progress', { sent, completed, total: opts.total, ok, err, inflight });
        }
      }
    });

    await Promise.all(workers);

    latencies.sort((a, b) => a - b);
    const p = (q) => latencies[Math.min(latencies.length - 1, Math.floor(q * latencies.length))] || 0;
    const avg = latencies.reduce((s, v) => s + v, 0) / (latencies.length || 1);
    const duration = ((Date.now() - startedAt) / 1000).toFixed(1);

    sendSse('summary', {
      ok,
      err,
      duration,
      p50: p(0.5),
      p90: p(0.9),
      p95: p(0.95),
      p99: p(0.99),
      avg: Math.round(avg)
    });

    current = null;
  }

  return {
    run,
    stop: () => { stopped = true; sendSse('log', { message: 'Teste interrompido.' }); },
    status: () => ({ sent, completed, total: opts.total, ok, err, inflight })
  };
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  if (url.pathname === '/') {
    const html = fs.readFileSync(UI_PATH, 'utf8');
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(html);
    return;
  }

  if (url.pathname === '/events') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive'
    });
    res.write('\n');
    sseClients.add(res);
    req.on('close', () => sseClients.delete(res));
    return;
  }

  if (url.pathname === '/start' && req.method === 'POST') {
    if (current) {
      res.writeHead(409, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Já existe um teste em execução.' }));
      return;
    }
    const body = await safeJson(req);
    const opts = {
      url: body.url || 'https://myn8n.seommerce.shop/webhook/planilha-atualizada',
      action: body.action || 'atualizar_agenda',
      total: Math.max(1, Number(body.total || 1000)),
      concurrency: Math.max(1, Number(body.concurrency || 50)),
      timeoutMs: Math.max(1000, Number(body.timeoutMs || 30000))
    };
    current = createRunner(opts);
    current.run();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, opts }));
    return;
  }

  if (url.pathname === '/stop' && req.method === 'POST') {
    if (current) current.stop();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true }));
    return;
  }

  if (url.pathname === '/status') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(current ? current.status() : { idle: true }));
    return;
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Not found');
});

const PORT = Number(process.env.PORT || 8787);
server.listen(PORT, () => {
  console.log(`Loadtest UI em http://localhost:${PORT}`);
});
