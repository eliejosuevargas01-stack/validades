#!/usr/bin/env node
'use strict';

const http = require('http');
const https = require('https');
const { URL } = require('url');

const defaults = {
  url: 'https://myn8n.seommerce.shop/webhook/planilha-atualizada',
  action: 'atualizar_agenda',
  total: 1000,
  concurrency: 50,
  timeoutMs: 30000
};

function getArg(name, fallback) {
  const idx = process.argv.indexOf(`--${name}`);
  if (idx === -1) return fallback;
  const val = process.argv[idx + 1];
  if (!val || val.startsWith('--')) return fallback;
  return val;
}

const url = getArg('url', defaults.url);
const action = getArg('action', defaults.action);
const total = Math.max(1, Number(getArg('total', defaults.total)) || defaults.total);
const concurrency = Math.max(1, Number(getArg('concurrency', defaults.concurrency)) || defaults.concurrency);
const timeoutMs = Math.max(1000, Number(getArg('timeout', defaults.timeoutMs)) || defaults.timeoutMs);

const target = new URL(url);
const agent = target.protocol === 'https:'
  ? new https.Agent({ keepAlive: true, maxSockets: concurrency * 4 })
  : new http.Agent({ keepAlive: true, maxSockets: concurrency * 4 });

let sent = 0;
let completed = 0;
let ok = 0;
let err = 0;
let inflight = 0;

const latencies = [];
const startedAt = Date.now();

function now() { return Date.now(); }

function requestOnce(seq) {
  return new Promise((resolve) => {
    const body = JSON.stringify({ action });
    const started = now();
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
      timeout: timeoutMs,
      agent
    };

    const transport = target.protocol === 'https:' ? https : http;
    const req = transport.request(options, (res) => {
      res.on('data', () => {});
      res.on('end', () => {
        const ms = now() - started;
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
      const ms = now() - started;
      latencies.push(ms);
      err += 1;
      resolve({ ok: false, status: 'ERR', ms, seq });
    });

    req.write(body);
    req.end();
  });
}

function printProgress() {
  const pct = ((completed / total) * 100).toFixed(1);
  process.stdout.write(`\rEnviadas: ${sent} | Concluídas: ${completed}/${total} (${pct}%) | OK: ${ok} | ERRO: ${err} | Em voo: ${inflight}   `);
}

async function run() {
  console.log(`Load test (Node) -> ${url}`);
  console.log(`Total: ${total} | Concorrência: ${concurrency} | Timeout: ${timeoutMs} ms | Action: ${action}`);

  const queue = Array.from({ length: total }, (_, i) => i + 1);

  const workers = Array.from({ length: concurrency }, async () => {
    while (queue.length) {
      const seq = queue.shift();
      sent += 1;
      inflight += 1;
      const res = await requestOnce(seq);
      inflight -= 1;
      completed += 1;
      if (completed % 50 === 0 || completed === total) printProgress();
      if (res.status === 'ERR' && completed % 200 === 0) {
        // keep stdout readable; avoid per-request logging
        // console.log(`\nFalha #${seq}`);
      }
    }
  });

  await Promise.all(workers);
  printProgress();
  console.log('\n');

  latencies.sort((a, b) => a - b);
  const p = (q) => latencies[Math.min(latencies.length - 1, Math.floor(q * latencies.length))] || 0;
  const avg = latencies.reduce((s, v) => s + v, 0) / (latencies.length || 1);
  const duration = ((now() - startedAt) / 1000).toFixed(1);

  console.log(`Resumo: OK ${ok} | ERRO ${err} | Duração ${duration}s`);
  console.log(`Latência (ms): p50 ${p(0.5)} | p90 ${p(0.9)} | p95 ${p(0.95)} | p99 ${p(0.99)} | média ${Math.round(avg)}`);
}

run().catch((err) => {
  console.error('Falha no load test:', err);
  process.exit(1);
});
